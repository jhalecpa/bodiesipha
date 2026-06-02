"""Main Click CLI for GTD Email Tool."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import click
from rich.console import Console
from rich.prompt import Prompt

from .auth import (
    CONFIG_DIR,
    CONFIG_FILE,
    authenticate_device_flow,
    is_authenticated,
    is_configured,
    load_config,
    save_config,
)
from .database import (
    GTDCategory,
    add_project,
    count_by_category,
    get_emails_by_category,
    get_past_decisions,
    get_projects,
    get_unprocessed_emails,
    initialize_db,
    save_weekly_review,
)
from .display import (
    console,
    print_capture_summary,
    print_email_card,
    print_email_list,
    print_error,
    print_header,
    print_info,
    print_review_dashboard,
    print_rule,
    print_success,
    print_warning,
    prompt_clarify_choice,
    prompt_confirm,
    prompt_due_date,
    prompt_next_action,
    prompt_notes,
    prompt_project_name,
    prompt_waiting_for,
)
from .gtd_processor import GTDProcessor


# ---------------------------------------------------------------------------
# Helper to ensure configuration / authentication before API calls
# ---------------------------------------------------------------------------

def _require_config() -> None:
    """Exit with a friendly message if the tool hasn't been set up yet."""
    if not is_configured():
        print_error(
            "GTD Email Tool has not been configured. Run [bold]gtd setup[/bold] first."
        )
        sys.exit(1)


def _require_auth() -> None:
    """Exit if there's no valid cached token."""
    _require_config()
    if not is_authenticated():
        print_error(
            "No valid authentication token found. Run [bold]gtd auth[/bold] to sign in."
        )
        sys.exit(1)


# ---------------------------------------------------------------------------
# CLI group
# ---------------------------------------------------------------------------

@click.group()
@click.version_option("0.1.0", prog_name="gtd")
def cli() -> None:
    """GTD Email Manager — Getting Things Done for your Outlook/Hotmail inbox.

    Run 'gtd setup' first to configure your Microsoft credentials, then
    'gtd auth' to sign in, and 'gtd capture' to fetch your emails.
    """
    # Ensure the local database is always initialised
    initialize_db()


# ---------------------------------------------------------------------------
# gtd setup
# ---------------------------------------------------------------------------

@cli.command("setup")
def cmd_setup() -> None:
    """Configure Microsoft Graph API credentials."""
    print_header("GTD Setup — Microsoft Graph API Credentials")

    console.print(
        "[bold]To connect GTD Email to your Outlook/Hotmail account you need to register\n"
        "a free application in the Microsoft Azure portal.[/bold]\n"
    )
    console.print(
        "Steps:\n"
        "  1. Go to [link=https://portal.azure.com]https://portal.azure.com[/link]\n"
        "  2. Search for [bold]App registrations[/bold] and click [bold]New registration[/bold]\n"
        "  3. Name it anything (e.g. 'GTD Email Tool')\n"
        "  4. Under [bold]Supported account types[/bold] choose\n"
        "     [italic]'Accounts in any organizational directory and personal Microsoft accounts'[/italic]\n"
        "  5. No Redirect URI needed (we use device flow)\n"
        "  6. After creation, copy the [bold]Application (client) ID[/bold]\n"
        "  7. Under [bold]API permissions[/bold] add:\n"
        "     Mail.ReadWrite, Mail.Send (Microsoft Graph, Delegated)\n"
    )
    print_rule()

    existing: dict = {}
    if CONFIG_FILE.exists():
        try:
            existing = load_config()
        except Exception:
            pass

    client_id = Prompt.ask(
        "[bold]Application (client) ID[/bold]",
        default=existing.get("client_id", ""),
    )
    tenant_id = Prompt.ask(
        "[bold]Tenant ID[/bold] (leave blank for personal Hotmail/Outlook accounts)",
        default=existing.get("tenant_id", "consumers"),
    )
    email = Prompt.ask(
        "[bold]Your email address[/bold]",
        default=existing.get("email", "jhalecpa@hotmail.com"),
    )

    console.print(
        "\n[bold]Optional: Anthropic API key for AI-powered bulk classification.[/bold]\n"
        "Get one at [link=https://console.anthropic.com]console.anthropic.com[/link]. "
        "Leave blank to skip.\n"
    )
    anthropic_key = Prompt.ask(
        "[bold]Anthropic API key[/bold] (optional, for [bold cyan]gtd ai-clarify[/bold cyan])",
        default=existing.get("anthropic_api_key", ""),
    )

    config = {
        "client_id": client_id.strip(),
        "tenant_id": tenant_id.strip() or "consumers",
        "email": email.strip(),
    }
    if anthropic_key.strip():
        config["anthropic_api_key"] = anthropic_key.strip()
    save_config(config)

    print_success(f"Configuration saved to {CONFIG_FILE}")
    console.print(
        "\nNext step: run [bold cyan]gtd auth[/bold cyan] to sign in with your Microsoft account."
    )


# ---------------------------------------------------------------------------
# gtd auth
# ---------------------------------------------------------------------------

@cli.command("auth")
def cmd_auth() -> None:
    """Authenticate with your Microsoft account via device flow."""
    _require_config()
    print_header("Microsoft Authentication")

    config = load_config()
    print_info(f"Signing in as: {config.get('email', 'unknown')}")

    try:
        token = authenticate_device_flow(config)
        print_success("Authentication successful! Token cached for future use.")
        console.print(
            "\nYou can now run [bold cyan]gtd capture[/bold cyan] to fetch your emails."
        )
    except RuntimeError as exc:
        print_error(str(exc))
        sys.exit(1)


# ---------------------------------------------------------------------------
# gtd capture
# ---------------------------------------------------------------------------

@cli.command("capture")
@click.option("--max", "max_emails", default=50, show_default=True,
              help="Maximum number of emails to fetch.")
@click.option("--all", "fetch_all", is_flag=True, default=False,
              help="Fetch every email in your inbox (may take a while for large inboxes).")
def cmd_capture(max_emails: int, fetch_all: bool) -> None:
    """Fetch new emails from your Outlook inbox."""
    _require_auth()
    print_header("Capture")

    if fetch_all:
        print_info("Fetching [bold]all[/bold] inbox emails — this may take a few minutes...")
        max_emails = -1

    processor = GTDProcessor()
    try:
        new_count, total_fetched = processor.capture(max_emails=max_emails)
        print_capture_summary(new_count, total_fetched)
    except Exception as exc:
        print_error(f"Capture failed: {exc}")
        sys.exit(1)

    unprocessed = len(processor.get_unprocessed())
    if unprocessed:
        print_info(
            f"You have [bold]{unprocessed}[/bold] unprocessed emails. "
            "Run [bold cyan]gtd clarify[/bold cyan] to process them."
        )


# ---------------------------------------------------------------------------
# gtd clarify
# ---------------------------------------------------------------------------

@cli.command("clarify")
@click.option("--offline", is_flag=True, default=False,
              help="Process emails without syncing changes back to Outlook.")
def cmd_clarify(offline: bool) -> None:
    """Interactively process unprocessed emails from your inbox."""
    if not offline:
        _require_auth()
    print_header("Clarify — Process Your Inbox")

    processor = GTDProcessor(offline=offline)
    unprocessed = processor.get_unprocessed()

    if not unprocessed:
        print_success("Your inbox is empty — nothing to clarify!")
        return

    print_info(f"{len(unprocessed)} email(s) to process.")
    print_rule()

    processed_count = 0

    for i, row in enumerate(unprocessed, 1):
        email = processor.format_email_summary(row)

        console.print(f"\n[dim]Email {i} of {len(unprocessed)}[/dim]")
        print_email_card(email, index=i)

        category = prompt_clarify_choice()
        if category is None:
            print_info("Clarify session ended. Returning to main menu.")
            break

        if category == GTDCategory.INBOX:
            # User chose to skip / keep in inbox
            print_info("Skipped — keeping in Inbox.")
            continue

        # Gather additional metadata depending on category
        next_action = ""
        waiting_for = ""
        project_name = ""
        due_date = ""
        notes = ""

        if category == GTDCategory.NEXT_ACTION:
            next_action = prompt_next_action()
            due_date = prompt_due_date()
            notes = prompt_notes()

        elif category == GTDCategory.WAITING:
            waiting_for = prompt_waiting_for()
            notes = prompt_notes()

        elif category == GTDCategory.PROJECT:
            project_name = prompt_project_name()
            next_action = prompt_next_action()
            notes = prompt_notes()

        elif category == GTDCategory.CALENDAR:
            due_date = prompt_due_date()
            notes = prompt_notes()

        elif category in (GTDCategory.REFERENCE, GTDCategory.SOMEDAY):
            notes = prompt_notes()

        elif category == GTDCategory.TRASH:
            if not prompt_confirm("Delete this email permanently in Outlook?"):
                category = GTDCategory.INBOX  # Abort, keep in inbox
                print_info("Aborted — keeping in Inbox.")
                continue

        try:
            processor.categorize(
                message_id=email["message_id"],
                category=category,
                next_action=next_action,
                project_name=project_name,
                due_date=due_date,
                notes=notes,
                waiting_for=waiting_for,
                sync_to_outlook=not offline,
            )
            cat_name = GTDCategory.DISPLAY_NAMES.get(category, category)
            print_success(f"Moved to: {cat_name}")
            processed_count += 1
        except Exception as exc:
            print_error(f"Failed to categorize: {exc}")

        print_rule()

    print_info(f"Processed {processed_count} email(s) this session.")


# ---------------------------------------------------------------------------
# gtd ai-clarify
# ---------------------------------------------------------------------------

@cli.command("ai-clarify")
@click.option("--max", "max_emails", default=500, show_default=True,
              help="Maximum number of inbox emails to classify.")
@click.option("--dry-run", is_flag=True, default=False,
              help="Show what AI would classify without saving anything.")
@click.option("--offline", is_flag=True, default=False,
              help="Skip syncing changes back to Outlook.")
@click.option("--model", default="claude-haiku-4-5-20251001", show_default=True,
              help="Claude model to use for classification.")
def cmd_ai_clarify(max_emails: int, dry_run: bool, offline: bool, model: str) -> None:
    """Bulk-classify inbox emails with AI (Claude) using the GTD methodology.

    The AI learns from your past manual decisions and applies the same patterns
    to new emails. Use --dry-run to preview before committing.
    """
    if not offline:
        _require_auth()

    config = load_config()
    api_key = config.get("anthropic_api_key", "")
    if not api_key:
        print_error(
            "No Anthropic API key configured. Run [bold]gtd setup[/bold] and enter your key, "
            "or set the [bold]ANTHROPIC_API_KEY[/bold] environment variable."
        )
        import os
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            sys.exit(1)

    try:
        import anthropic as _anthropic_check  # noqa: F401
    except ImportError:
        print_error(
            "The [bold]anthropic[/bold] package is not installed. "
            "Run: [bold]pip install anthropic[/bold]"
        )
        sys.exit(1)

    from .ai_classifier import classify_all

    print_header("AI Clarify — Bulk GTD Classification")

    processor = GTDProcessor(offline=offline)
    unprocessed = processor.get_unprocessed()

    if not unprocessed:
        print_success("No unprocessed emails — inbox is clean!")
        return

    emails = [processor.format_email_summary(r) for r in unprocessed[:max_emails]]
    total = len(emails)
    print_info(f"Classifying [bold]{total}[/bold] emails with Claude...")

    # Load past decisions for few-shot learning
    past_rows = get_past_decisions(max_per_category=4)
    if past_rows:
        past_examples = [processor.format_email_summary(r) for r in past_rows]
        print_info(
            f"Using [bold]{len(past_examples)}[/bold] of your past decisions as examples."
        )
    else:
        past_examples = []
        print_info("No past decisions yet — AI will use general GTD rules.")

    print_rule()

    from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeElapsedColumn

    results: list[tuple[dict, dict]] = []
    errors: list[str] = []

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("{task.completed}/{task.total} emails"),
        TimeElapsedColumn(),
        console=console,
    ) as progress:
        task = progress.add_task("Classifying...", total=total)

        def on_progress(done: int, _total: int) -> None:
            progress.update(task, completed=done)

        try:
            results = classify_all(
                emails,
                api_key=api_key,
                examples=past_examples,
                model=model,
                progress_callback=on_progress,
            )
        except Exception as exc:
            print_error(f"AI classification failed: {exc}")
            sys.exit(1)

    # Tally results
    from collections import Counter
    tally: Counter = Counter()
    for _email, result in results:
        tally[result["category"]] += 1

    # Show summary table
    from rich.table import Table
    from rich import box as rbox

    table = Table(title="AI Classification Summary", box=rbox.ROUNDED)
    table.add_column("Category", style="bold")
    table.add_column("Count", justify="right")
    table.add_column("% of inbox", justify="right")
    for cat in GTDCategory.ALL:
        if cat == GTDCategory.INBOX:
            continue
        count = tally.get(cat, 0)
        if count == 0:
            continue
        pct = f"{count / total * 100:.0f}%"
        table.add_row(GTDCategory.DISPLAY_NAMES.get(cat, cat), str(count), pct)
    if tally.get(GTDCategory.INBOX, 0):
        table.add_row(
            "[dim]Inbox (needs review)[/dim]",
            str(tally[GTDCategory.INBOX]),
            f"{tally[GTDCategory.INBOX] / total * 100:.0f}%",
        )
    console.print(table)

    if dry_run:
        print_warning("Dry run — no changes saved. Re-run without [bold]--dry-run[/bold] to apply.")
        return

    # --- Bulk delete prompt for trash ---
    trash_pairs = [(e, r) for e, r in results if r["category"] == GTDCategory.TRASH]
    delete_trash = False
    if trash_pairs:
        print_rule()
        console.print(f"\n[bold red]Trash ({len(trash_pairs)} emails)[/bold red] — sample:\n")
        for email, result in trash_pairs[:10]:
            console.print(
                f"  [dim]•[/dim] {email['subject'][:60]:<60}  "
                f"[dim]{email['sender'][:30]}[/dim]\n"
                f"    [italic dim]{result.get('reasoning', '')}[/italic dim]"
            )
        if len(trash_pairs) > 10:
            console.print(f"  [dim]... and {len(trash_pairs) - 10} more[/dim]")
        console.print()
        delete_trash = prompt_confirm(
            f"Permanently delete these [bold]{len(trash_pairs)}[/bold] emails from Outlook?"
        )
        if not delete_trash:
            print_info("Trash emails will be skipped (left in inbox).")

    # --- Confirm remaining classifications ---
    non_trash = [(e, r) for e, r in results if r["category"] != GTDCategory.TRASH]
    non_inbox = [(e, r) for e, r in non_trash if r["category"] != GTDCategory.INBOX]
    if non_inbox:
        print_rule()
        if not prompt_confirm(f"Apply [bold]{len(non_inbox)}[/bold] other classifications?"):
            print_info("Aborted — no changes made.")
            return

    # --- Apply ---
    applied = 0
    deleted = 0
    skipped = 0
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("{task.completed}/{task.total}"),
        console=console,
    ) as progress:
        task = progress.add_task("Applying...", total=total)
        for email, result in results:
            cat = result["category"]

            if cat == GTDCategory.INBOX:
                skipped += 1
                progress.advance(task)
                continue

            if cat == GTDCategory.TRASH and not delete_trash:
                skipped += 1
                progress.advance(task)
                continue

            try:
                processor.categorize(
                    message_id=email["message_id"],
                    category=cat,
                    next_action=result.get("next_action", ""),
                    notes=result.get("reasoning", "") + (
                        (" | " + result["notes"]) if result.get("notes") else ""
                    ),
                    sync_to_outlook=not offline,
                )
                if cat == GTDCategory.TRASH:
                    deleted += 1
                else:
                    applied += 1
            except Exception as exc:
                errors.append(f"{email['subject'][:40]}: {exc}")
            progress.advance(task)

    parts = []
    if applied:
        parts.append(f"{applied} classified")
    if deleted:
        parts.append(f"{deleted} deleted")
    if skipped:
        parts.append(f"{skipped} left in inbox")
    print_success("Done! " + ", ".join(parts) + ".")
    if errors:
        print_warning(f"{len(errors)} error(s) encountered:")
        for e in errors[:5]:
            console.print(f"  [red]•[/red] {e}")


# ---------------------------------------------------------------------------
# gtd next-actions
# ---------------------------------------------------------------------------

@cli.command("next-actions")
def cmd_next_actions() -> None:
    """Show your Next Actions list."""
    _require_auth()
    print_header("Next Actions")

    emails = [
        GTDProcessor(offline=True).format_email_summary(row)
        for row in get_emails_by_category(GTDCategory.NEXT_ACTION)
    ]
    print_email_list(emails, "Next Actions", GTDCategory.NEXT_ACTION)

    if emails:
        console.print(
            "\n[dim]Tip: reply/forward/archive emails directly in Outlook. "
            "Re-run [bold]gtd capture[/bold] + [bold]gtd clarify[/bold] to keep in sync.[/dim]"
        )


# ---------------------------------------------------------------------------
# gtd waiting
# ---------------------------------------------------------------------------

@cli.command("waiting")
def cmd_waiting() -> None:
    """Show the Waiting For list."""
    _require_auth()
    print_header("Waiting For")

    p = GTDProcessor(offline=True)
    emails = [p.format_email_summary(r) for r in get_emails_by_category(GTDCategory.WAITING)]
    print_email_list(emails, "Waiting For", GTDCategory.WAITING)


# ---------------------------------------------------------------------------
# gtd projects
# ---------------------------------------------------------------------------

@cli.command("projects")
@click.option("--add", "add_project_flag", is_flag=True, default=False,
              help="Add a new project.")
def cmd_projects(add_project_flag: bool) -> None:
    """Show (or add) Projects."""
    _require_auth()
    print_header("Projects")

    if add_project_flag:
        name = Prompt.ask("[magenta]Project name[/magenta]")
        description = Prompt.ask("[dim]Description (optional)[/dim]", default="")
        pid = add_project(name, description)
        print_success(f"Project '{name}' added (ID {pid}).")
        return

    # Show project emails
    p = GTDProcessor(offline=True)
    emails = [p.format_email_summary(r) for r in get_emails_by_category(GTDCategory.PROJECT)]
    print_email_list(emails, "Projects", GTDCategory.PROJECT)

    # Show named projects from projects table
    projects = get_projects()
    if projects:
        from rich.table import Table
        from rich import box as rbox
        table = Table(title="Named Projects", box=rbox.ROUNDED)
        table.add_column("ID", style="dim", width=6)
        table.add_column("Name", style="bold magenta")
        table.add_column("Description")
        table.add_column("Status", width=10)
        table.add_column("Created", width=12)
        for proj in projects:
            table.add_row(
                str(proj["id"]),
                proj["name"],
                proj["description"] or "",
                proj["status"],
                (proj["created_at"] or "")[:10],
            )
        console.print(table)


# ---------------------------------------------------------------------------
# gtd someday
# ---------------------------------------------------------------------------

@cli.command("someday")
def cmd_someday() -> None:
    """Show the Someday / Maybe list."""
    _require_auth()
    print_header("Someday / Maybe")

    p = GTDProcessor(offline=True)
    emails = [p.format_email_summary(r) for r in get_emails_by_category(GTDCategory.SOMEDAY)]
    print_email_list(emails, "Someday / Maybe", GTDCategory.SOMEDAY)


# ---------------------------------------------------------------------------
# gtd reference
# ---------------------------------------------------------------------------

@cli.command("reference")
def cmd_reference() -> None:
    """Show the Reference archive."""
    _require_auth()
    print_header("Reference")

    p = GTDProcessor(offline=True)
    emails = [p.format_email_summary(r) for r in get_emails_by_category(GTDCategory.REFERENCE)]
    print_email_list(emails, "Reference", GTDCategory.REFERENCE)


# ---------------------------------------------------------------------------
# gtd review
# ---------------------------------------------------------------------------

@cli.command("review")
@click.option("--save", is_flag=True, default=False,
              help="Save this review snapshot to the database.")
def cmd_review(save: bool) -> None:
    """Show the weekly review dashboard."""
    _require_auth()

    counts = count_by_category()
    print_review_dashboard(counts)

    # Tips
    inbox_count = counts.get(GTDCategory.INBOX, 0)
    next_count = counts.get(GTDCategory.NEXT_ACTION, 0)
    waiting_count = counts.get(GTDCategory.WAITING, 0)

    if inbox_count:
        print_warning(
            f"You have {inbox_count} unprocessed emails. Run [bold]gtd clarify[/bold]."
        )
    if next_count:
        print_info(
            f"You have {next_count} Next Actions. Run [bold]gtd next-actions[/bold] to see them."
        )
    if waiting_count:
        print_info(
            f"You have {waiting_count} items you're waiting on. Run [bold]gtd waiting[/bold]."
        )

    if save:
        notes = Prompt.ask("[dim]Review notes (optional)[/dim]", default="")
        rid = save_weekly_review(notes=notes)
        print_success(f"Review #{rid} saved.")


# ---------------------------------------------------------------------------
# gtd engage  (alias for next-actions with interactive mark-done)
# ---------------------------------------------------------------------------

@cli.command("engage")
@click.option("--offline", is_flag=True, default=False,
              help="Work through list without syncing back to Outlook.")
def cmd_engage(offline: bool) -> None:
    """Work through your Next Actions list interactively."""
    if not offline:
        _require_auth()
    print_header("Engage — Work Through Next Actions")

    p = GTDProcessor(offline=offline)
    next_actions = get_emails_by_category(GTDCategory.NEXT_ACTION)

    if not next_actions:
        print_success("No Next Actions — you're all caught up!")
        return

    for i, row in enumerate(next_actions, 1):
        email = p.format_email_summary(row)
        console.print(f"\n[dim]Action {i} of {len(next_actions)}[/dim]")
        print_email_card(email, index=i)

        choice = Prompt.ask(
            "[bold]What do you want to do?[/bold]\n"
            "  [bold green]d[/bold green] Done (archive)\n"
            "  [bold yellow]s[/bold yellow] Skip\n"
            "  [bold red]q[/bold red] Quit\n"
            "[bold]Choice[/bold]",
            choices=["d", "s", "q"],
            default="s",
            show_choices=False,
        )
        if choice == "q":
            break
        elif choice == "d":
            try:
                p.mark_action_done(email["message_id"])
                print_success("Marked as done and archived.")
            except Exception as exc:
                print_error(f"Could not archive: {exc}")
        elif choice == "s":
            print_info("Skipped.")

        print_rule()

    print_info("Engage session complete.")


# ---------------------------------------------------------------------------
# gtd status  (quick at-a-glance summary)
# ---------------------------------------------------------------------------

@cli.command("status")
def cmd_status() -> None:
    """Show a quick summary of your GTD buckets."""
    initialize_db()
    counts = count_by_category()

    print_header("GTD Status")
    total = sum(counts.values())
    for cat in GTDCategory.ALL:
        count = counts.get(cat, 0)
        if count == 0:
            continue
        name = GTDCategory.DISPLAY_NAMES.get(cat, cat)
        bar = "#" * min(count, 40)
        console.print(f"  {name:<25} {count:>4}  [dim]{bar}[/dim]")
    console.print(f"\n  {'Total':<25} {total:>4}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    cli()


if __name__ == "__main__":
    main()
