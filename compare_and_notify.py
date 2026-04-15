from __future__ import annotations

import argparse
import asyncio
import csv
import json
import sys
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import compare_jinshuju as compare
from notify import (
    DEFAULT_CONFIG_PATH,
    DEFAULT_MESSAGE_TEMPLATE,
    NapCatWsClient,
    Recipient,
    ResultRow,
    build_message,
    load_ws_server_config,
    response_error_text,
    truncate_preview,
    verify_group_member,
    write_results,
)
from playwright.sync_api import sync_playwright


DEFAULT_OUTPUT_ROOT = Path(r"C:\Users\Theta\napcat_notify\runs")


@dataclass
class SkippedRecipient:
    name: str
    college: str
    qq: str
    source: str
    reason: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compare Jinshuju registrations and precheck/send QQ notifications")
    parser.add_argument("--form-title", help="Form title on the Jinshuju home page")
    parser.add_argument("--entries-url", help="Direct Jinshuju entries URL")
    parser.add_argument("--qualification-file", required=True, type=Path, help="Qualification Excel file path")
    parser.add_argument("--name-field-label", default="姓名", help="Name column label in the export file")
    parser.add_argument("--college-field-label", default="学院", help="College column label in the export file")
    parser.add_argument("--qq-field-label", default="QQ号", help="QQ column label in the export file")
    parser.add_argument("--created-after", help="Only include submissions created after this timestamp")
    parser.add_argument("--download-format", choices=("xlsx", "csv"), default=compare.DEFAULT_DOWNLOAD_FORMAT)
    parser.add_argument("--profile-dir", type=Path, default=compare.DEFAULT_PROFILE_DIR, help="Persistent browser profile directory")
    parser.add_argument("--output-root", type=Path, default=DEFAULT_OUTPUT_ROOT, help="Root output directory")
    parser.add_argument("--home-url", default=compare.DEFAULT_HOME_URL, help="Jinshuju home URL")
    parser.add_argument("--headless", action="store_true", help="Run browser automation in headless mode")
    parser.add_argument(
        "--notify-target",
        choices=("qualified_not_registered", "registered_not_qualified"),
        default="qualified_not_registered",
        help="Which comparison result set should be converted into recipients",
    )
    parser.add_argument("--message-template", default=DEFAULT_MESSAGE_TEMPLATE, help="Notification message template; may contain {name}")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG_PATH, help="NapCat OneBot11 config path")
    parser.add_argument("--group-id", type=int, help="If provided, use group temporary sessions for precheck/send")
    parser.add_argument("--delay", type=float, default=2.0, help="Delay in seconds between notification requests")
    parser.add_argument("--limit", type=int, help="Only process the first N recipients")
    parser.add_argument("--send", action="store_true", help="Actually send messages; default only performs precheck and dry-run")
    args = parser.parse_args()
    if not args.form_title and not args.entries_url:
        parser.error("--form-title and --entries-url: provide at least one")
    return args


def make_run_dir(output_root: Path) -> Path:
    stamp = datetime.now().strftime("compare-notify-%Y%m%d-%H%M%S")
    run_dir = output_root / stamp
    run_dir.mkdir(parents=True, exist_ok=False)
    return run_dir


def write_skipped(path: Path, rows: list[SkippedRecipient]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=["name", "college", "qq", "source", "reason"])
        writer.writeheader()
        for row in rows:
            writer.writerow(asdict(row))


def write_recipients(path: Path, recipients: list[Recipient]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=["QQ号", "姓名", "学院", "消息"])
        writer.writeheader()
        for recipient in recipients:
            writer.writerow(
                {
                    "QQ号": recipient.qq,
                    "姓名": recipient.name,
                    "学院": recipient.college,
                    "消息": recipient.message or "",
                }
            )


def build_source_label(record: Any) -> str:
    if isinstance(record, compare.QualificationRecord):
        return f"{record.source}#{record.sheet}:{record.row_number}"
    return f"{record.source}#{record.row_number}"


def build_recipients(records: list[Any], template: str) -> tuple[list[Recipient], list[SkippedRecipient]]:
    recipients: list[Recipient] = []
    skipped: list[SkippedRecipient] = []
    for record in records:
        qq = compare.normalize_qq(getattr(record, "qq", ""))
        name = compare.normalize_name(getattr(record, "name", "")) or str(getattr(record, "name", "")).strip()
        college = str(getattr(record, "college", "")).strip()
        source = build_source_label(record)
        if not qq:
            skipped.append(
                SkippedRecipient(
                    name=name,
                    college=college,
                    qq=str(getattr(record, "qq", "")),
                    source=source,
                    reason="missing_or_invalid_qq",
                )
            )
            continue
        recipients.append(Recipient(qq=qq, name=name, college=college, message=template))
    return recipients, skipped


def print_notify_preview(recipients: list[Recipient], mode: str, group_id: int | None, notify_target: str) -> None:
    print()
    print(f"Notify target: {notify_target}")
    print(f"Notify mode: {mode}")
    if group_id:
        print(f"Delivery mode: group temporary session (group_id={group_id})")
    else:
        print("Delivery mode: private chat")
    print(f"Recipients: {len(recipients)}")
    print()
    for index, recipient in enumerate(recipients, start=1):
        print(f"[{index:02d}] {recipient.name} ({recipient.qq}) - {recipient.college}")
        print(f"     {truncate_preview(build_message(recipient), 120)}")


async def run_notify(
    *,
    config_path: Path,
    recipients: list[Recipient],
    notify_dir: Path,
    send: bool,
    group_id: int | None,
    delay: float,
) -> dict[str, Any]:
    ws_config = load_ws_server_config(config_path)
    mode = "send" if send else "precheck"
    rows: list[ResultRow] = []

    async with NapCatWsClient(
        host=ws_config["host"],
        port=int(ws_config["port"]),
        token=str(ws_config["token"]),
    ) as client:
        login_response = await client.request("get_login_info")
        if login_response.get("status") != "ok" or login_response.get("retcode") != 0:
            raise RuntimeError("get_login_info failed: " + response_error_text(login_response))
        login_info = login_response.get("data") or {}
        print()
        print(f"NapCat precheck passed: {login_info.get('user_id')} / {login_info.get('nickname')}")

        for index, recipient in enumerate(recipients, start=1):
            message = build_message(recipient)
            sent_at = datetime.now().isoformat(timespec="seconds")
            try:
                if group_id:
                    is_member, member_error = await verify_group_member(client, group_id, int(recipient.qq))
                    if not is_member:
                        rows.append(
                            ResultRow(
                                qq=recipient.qq,
                                name=recipient.name,
                                college=recipient.college,
                                mode=mode,
                                status="api_failed",
                                message_id="",
                                error=f"group_member_check_failed: {member_error}",
                                sent_at=sent_at,
                                message_preview=truncate_preview(message, 120),
                            )
                        )
                        print(f"[{index:02d}] Precheck failed: {recipient.name} ({recipient.qq}) -> {member_error}")
                        if index < len(recipients):
                            await asyncio.sleep(delay)
                        continue

                if not send:
                    rows.append(
                        ResultRow(
                            qq=recipient.qq,
                            name=recipient.name,
                            college=recipient.college,
                            mode=mode,
                            status="dry_run",
                            message_id="",
                            error="",
                            sent_at="",
                            message_preview=truncate_preview(message, 120),
                        )
                    )
                    print(f"[{index:02d}] Prechecked: {recipient.name} ({recipient.qq})")
                    if index < len(recipients):
                        await asyncio.sleep(delay)
                    continue

                params: dict[str, Any] = {
                    "user_id": int(recipient.qq),
                    "message": message,
                }
                if group_id:
                    params["group_id"] = int(group_id)
                response = await client.request("send_private_msg", params)
                if response.get("status") == "ok" and response.get("retcode") == 0:
                    data = response.get("data") or {}
                    rows.append(
                        ResultRow(
                            qq=recipient.qq,
                            name=recipient.name,
                            college=recipient.college,
                            mode=mode,
                            status="sent",
                            message_id=str(data.get("message_id", "")),
                            error="",
                            sent_at=sent_at,
                            message_preview=truncate_preview(message, 120),
                        )
                    )
                    print(f"[{index:02d}] Sent: {recipient.name} ({recipient.qq})")
                else:
                    error = response_error_text(response)
                    rows.append(
                        ResultRow(
                            qq=recipient.qq,
                            name=recipient.name,
                            college=recipient.college,
                            mode=mode,
                            status="api_failed",
                            message_id="",
                            error=error,
                            sent_at=sent_at,
                            message_preview=truncate_preview(message, 120),
                        )
                    )
                    print(f"[{index:02d}] Send failed: {recipient.name} ({recipient.qq}) -> {error}")
            except Exception as exc:
                rows.append(
                    ResultRow(
                        qq=recipient.qq,
                        name=recipient.name,
                        college=recipient.college,
                        mode=mode,
                        status="transport_failed",
                        message_id="",
                        error=str(exc),
                        sent_at=sent_at,
                        message_preview=truncate_preview(message, 120),
                    )
                )
                print(f"[{index:02d}] Transport failed: {recipient.name} ({recipient.qq}) -> {exc}")
            if index < len(recipients):
                await asyncio.sleep(delay)

    write_results(notify_dir, rows)
    summary = {
        "generated_at": datetime.now().astimezone().isoformat(),
        "notify_dir": str(notify_dir),
        "mode": mode,
        "group_id": group_id or 0,
        "total": len(rows),
        "dry_run_total": sum(1 for row in rows if row.status == "dry_run"),
        "sent_total": sum(1 for row in rows if row.status == "sent"),
        "api_failed_total": sum(1 for row in rows if row.status == "api_failed"),
        "transport_failed_total": sum(1 for row in rows if row.status == "transport_failed"),
    }
    (notify_dir / "notify_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return summary


def print_notify_summary(summary: dict[str, Any]) -> None:
    print()
    print(f"Notify output: {summary['notify_dir']}")
    print(f"Mode: {summary['mode']}")
    print(f"Total: {summary['total']}")
    print(f"dry_run: {summary['dry_run_total']}")
    print(f"sent: {summary['sent_total']}")
    print(f"api_failed: {summary['api_failed_total']}")
    print(f"transport_failed: {summary['transport_failed_total']}")


def main() -> int:
    args = parse_args()
    created_after = compare.parse_user_datetime(args.created_after) if args.created_after else None
    run_dir = make_run_dir(args.output_root)
    notify_dir = run_dir / "notify"
    notify_dir.mkdir(parents=True, exist_ok=True)

    qualifications = compare.load_qualification_records(args.qualification_file)
    args.profile_dir.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        context, active_profile_dir = compare.launch_context_with_fallback(p, args.profile_dir, args.headless)
        try:
            export_path = compare.export_jinshuju_file(
                context=context,
                run_dir=run_dir,
                form_title=args.form_title,
                entries_url=args.entries_url,
                home_url=args.home_url,
                download_format=args.download_format,
            )
        finally:
            context.close()

    form_records = compare.load_form_records_from_export(
        export_path,
        name_label=args.name_field_label,
        college_label=args.college_field_label,
        qq_label=args.qq_field_label,
        created_after=created_after,
    )
    deduped_qualifications, qualification_duplicates = compare.dedupe_qualifications(qualifications)
    deduped_form_records, form_duplicates = compare.dedupe_form_records(form_records)
    matched, qualified_not_registered, registered_not_qualified = compare.compare_records(
        deduped_qualifications,
        deduped_form_records,
    )

    compare_summary = {
        "generated_at": datetime.now().astimezone().isoformat(),
        "output_dir": str(run_dir),
        "export_file": str(export_path),
        "form_title": args.form_title or "",
        "entries_url": args.entries_url or "",
        "qualification_file": str(args.qualification_file),
        "name_field_label": args.name_field_label,
        "college_field_label": args.college_field_label,
        "qq_field_label": args.qq_field_label,
        "created_after": args.created_after or "",
        "download_format": args.download_format,
        "profile_dir": str(active_profile_dir),
        "qualification_total": len(deduped_qualifications),
        "form_total": len(deduped_form_records),
        "matched_total": len(matched),
        "qualified_not_registered_total": len(qualified_not_registered),
        "registered_not_qualified_total": len(registered_not_qualified),
        "qualification_duplicates_total": len(qualification_duplicates),
        "form_duplicates_total": len(form_duplicates),
    }
    compare.write_outputs(
        run_dir,
        matched=matched,
        qualified_not_registered=qualified_not_registered,
        registered_not_qualified=registered_not_qualified,
        qualification_duplicates=qualification_duplicates,
        form_duplicates=form_duplicates,
        summary=compare_summary,
    )
    compare.print_summary(compare_summary)

    if args.notify_target == "qualified_not_registered":
        source_records: list[Any] = qualified_not_registered
    else:
        source_records = registered_not_qualified

    recipients, skipped = build_recipients(source_records, args.message_template)
    if args.limit is not None:
        recipients = recipients[: args.limit]
    write_recipients(run_dir / "notify_candidates.csv", recipients)
    write_skipped(run_dir / "notify_skipped.csv", skipped)

    if not recipients:
        print()
        print("No recipients generated from the selected comparison result.")
        return 0

    print_notify_preview(
        recipients,
        mode="send" if args.send else "precheck",
        group_id=args.group_id,
        notify_target=args.notify_target,
    )
    if skipped:
        print()
        print(f"Skipped {len(skipped)} records without valid QQ; see {run_dir / 'notify_skipped.csv'}")

    try:
        notify_summary = asyncio.run(
            run_notify(
                config_path=args.config,
                recipients=recipients,
                notify_dir=notify_dir,
                send=args.send,
                group_id=args.group_id,
                delay=args.delay,
            )
        )
    except KeyboardInterrupt:
        print("Interrupted by user")
        return 130
    except Exception as exc:
        print(f"Notification stage failed: {exc}", file=sys.stderr)
        return 1

    print_notify_summary(notify_summary)
    if not args.send:
        rerun_parts = [
            "python .\\compare_and_notify.py",
            f'--qualification-file "{args.qualification_file}"',
        ]
        if args.form_title:
            rerun_parts.append(f'--form-title "{args.form_title}"')
        if args.entries_url:
            rerun_parts.append(f'--entries-url "{args.entries_url}"')
        rerun_parts.append(f'--profile-dir "{args.profile_dir}"')
        if args.group_id:
            rerun_parts.append(f"--group-id {args.group_id}")
        if args.headless:
            rerun_parts.append("--headless")
        rerun_parts.append("--send")
        print()
        print("If the precheck looks correct, rerun with:")
        print(" ".join(rerun_parts))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
