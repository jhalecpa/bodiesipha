"""Rich terminal display components for GTD Email Tool."""

from __future__ import annotations

from datetime import datetime
from typing import Any

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.text import Text
from rich import box
from rich.columns import Columns
from rich.rule import Rule
from rich.prompt import Prompt, Confirm

console = Console()


# ---------------------------------------------------------------------------
# Color / style constants
# ---------------------------------------------------------------------------

CATEGORY_STYLES = {
    "inbox": "bold cyan",
    "trash": "dim red",
    "reference": "steel_blue",
    "someday": "dark_orange",
    "waiting": "yellow",
    "next_action": "bold green",
    "project": "bold magenta",
    "calendar": "bright_blue",
}

CATEGORY_ICONS = {
    "inbox": "[?]",
    "trash": "[X]",
    "reference": "[R]",
    "someday": "[S]",
    "waiting": "[W]",
    "next_action": "[>]",
    "project": "[P]",
    "calendar": "[C]",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _relative_date(iso_str: str) -> str:
    """Return a human-friendly relative date string."""
    if not iso_str:
        return "unknown"
    try:
        dt = datetime.fromisoformat(iso_str.replace("Z", "+00:00"))
        now = datetime.now(dt.tzinfo)
        delta = now - dt
        days = delta.days
        if days == 0:
            return "Today"
        elif days == 1:
            return "Yesterday"
        elif days < 7:
            return f"{days}d ago"
        elif days < 30:
            return f"{days // 7}w ago"
        else:
            return iso_str[:10]
    except Exception:
        return iso_str[:10]


# ---------------------------------------------------------------------------
# Header / Footer
# ---------------------------------------------------------------------------

def print_header(title: str = "GTD Email Manager") -> None:
    """Print the application header."""
    console.print()
    console.print(
        Panel(
            f"[bold white]{title}[/bold white]\n"
            "[dim]Getting Things Done for jhalecpa@hotmail.com[/dim]",
            style="blue",
            expand=False,
        )
    )
    console.print()


def print_rule(title: str = "") -> None:
    """Print a horizontal rule."""
    console.print(Rule(title, style="dim"))


# ---------------------------------------------------------------------------
# Email display
# ---------------------------------------------------------------------------

def print_email_card(email: dict, index: int = 0) -> None:
    """Print a single email as a rich panel."""
    subj = email.get("subject", "(no subject)")
    sender = email.get("sender", "Unknown")
    sender_email = email.get("sender_email", "")
    received = _relative_date(email.get("received_at", ""))
    preview = email.get("body_preview", "")[:300]
    has_attach = "(attachment) " if email.get("has_attachments") else ""
    is_read = email.get("is_read", False)
    category = email.get("category", "inbox")

    read_marker = "" if is_read else "[bold]NEW[/bold] "
    cat_icon = CATEGORY_ICONS.get(category, "[?]")
    cat_style = CATEGORY_STYLES.get(category, "white")

    title_text = (
        f"[{cat_style}]{cat_icon}[/{cat_style}] "
        f"#{index} {read_marker}"
        f"[bold]{subj}[/bold]"
    )

    body_text = (
        f"[dim]From:[/dim] {sender} <{sender_email}>\n"
        f"[dim]Received:[/dim] {received}  {has_attach}\n\n"
        f"{preview}"
    )

    if email.get("next_action"):
        body_text += f"\n\n[green]Next Action:[/green] {email['next_action']}"
    if email.get("waiting_for"):
        body_text += f"\n[yellow]Waiting For:[/yellow] {email['waiting_for']}"
    if email.get("project_name"):
        body_text += f"\n[magenta]Project:[/magenta] {email['project_name']}"
    if email.get("due_date"):
        body_text += f"\n[red]Due:[/red] {email['due_date']}"
    if email.get("notes"):
        body_text += f"\n[cyan]Notes:[/cyan] {email['notes']}"

    console.print(Panel(body_text, title=title_text, border_style=cat_style))


def print_email_list(emails: list[dict], title: str, category: str = "") -> None:
    """Print a table of emails for a GTD bucket."""
    if not emails:
        console.print(f"[dim]No items in {title}.[/dim]")
        return

    cat_style = CATEGORY_STYLES.get(category, "white")
    table = Table(
        title=f"[{cat_style}]{title}[/{cat_style}] ({len(emails)} items)",
        box=box.ROUNDED,
        show_lines=False,
        expand=True,
    )
    table.add_column("#", style="dim", width=4, no_wrap=True)
    table.add_column("Subject", style="bold", ratio=3)
    table.add_column("From", ratio=2)
    table.add_column("Date", width=12)
    table.add_column("Extra", ratio=2, style="dim")

    for i, email in enumerate(emails, 1):
        subj = email.get("subject", "(no subject)")
        sender = email.get("sender") or email.get("sender_email", "Unknown")
        received = _relative_date(email.get("received_at", ""))
        extra = (
            email.get("next_action")
            or email.get("waiting_for")
            or email.get("project_name")
            or ""
        )
        read_style = "" if email.get("is_read") else "bold"
        table.add_row(
            str(i),
            Text(subj[:70], style=read_style),
            sender[:40],
            received,
            extra[:50],
        )

    console.print(table)


# ---------------------------------------------------------------------------
# Review dashboard
# ---------------------------------------------------------------------------

def print_review_dashboard(counts: dict[str, int]) -> None:
    """Print the weekly review dashboard with counts per bucket."""
    print_header("GTD Weekly Review Dashboard")

    buckets = [
        ("inbox",       "Inbox (Unprocessed)",  "cyan"),
        ("next_action", "Next Actions",          "green"),
        ("waiting",     "Waiting For",           "yellow"),
        ("project",     "Projects",              "magenta"),
        ("someday",     "Someday / Maybe",        "dark_orange"),
        ("reference",   "Reference",             "steel_blue"),
        ("calendar",    "Calendar",              "bright_blue"),
        ("trash",       "Trash",                 "red"),
    ]

    panels = []
    for key, label, color in buckets:
        count = counts.get(key, 0)
        badge = f"[{color}]{count:>4}[/{color}]"
        icon = CATEGORY_ICONS.get(key, "[?]")
        panels.append(
            Panel(
                f"{badge}\n[dim]{icon}[/dim]",
                title=f"[{color}]{label}[/{color}]",
                border_style=color,
                expand=True,
                width=26,
            )
        )

    console.print(Columns(panels, equal=True, expand=True))
    console.print()


def print_capture_summary(new_count: int, total_fetched: int) -> None:
    """Print the result of a capture operation."""
    console.print(
        f"[green]Capture complete.[/green] "
        f"Fetched [bold]{total_fetched}[/bold] emails, "
        f"[bold cyan]{new_count}[/bold cyan] new."
    )


# ---------------------------------------------------------------------------
# Interactive clarify prompts
# ---------------------------------------------------------------------------

CLARIFY_MENU = """
[bold]Categorise this email:[/bold]
  [bold cyan]i[/bold cyan]  Keep in Inbox (skip)
  [bold red]t[/bold red]  Trash / Delete
  [bold steel_blue]r[/bold steel_blue]  Reference (keep, non-actionable)
  [bold dark_orange]s[/bold dark_orange]  Someday / Maybe
  [bold yellow]w[/bold yellow]  Waiting For
  [bold green]n[/bold green]  Next Action
  [bold magenta]p[/bold magenta]  Project
  [bold bright_blue]c[/bold bright_blue]  Calendar
  [bold dim]q[/bold dim]  Quit clarify session
"""

CLARIFY_KEYS = {
    "i": "inbox",
    "t": "trash",
    "r": "reference",
    "s": "someday",
    "w": "waiting",
    "n": "next_action",
    "p": "project",
    "c": "calendar",
    "q": None,  # quit signal
}


def prompt_clarify_choice() -> str | None:
    """
    Show the clarify menu and return the selected GTD category key,
    or None if the user wants to quit.
    """
    console.print(CLARIFY_MENU)
    while True:
        choice = Prompt.ask(
            "[bold]Choice[/bold]",
            choices=list(CLARIFY_KEYS.keys()),
            default="i",
            show_choices=False,
        )
        category = CLARIFY_KEYS.get(choice)
        if choice == "q":
            return None
        return category


def prompt_next_action() -> str:
    """Prompt user for a next action description."""
    return Prompt.ask("[green]Next Action[/green] (brief description)", default="")


def prompt_waiting_for() -> str:
    """Prompt user for who/what they are waiting for."""
    return Prompt.ask("[yellow]Waiting For[/yellow] (who/what)", default="")


def prompt_project_name() -> str:
    """Prompt user for a project name."""
    return Prompt.ask("[magenta]Project Name[/magenta]", default="")


def prompt_due_date() -> str:
    """Prompt user for a due date (optional)."""
    return Prompt.ask("[red]Due Date[/red] (YYYY-MM-DD, optional)", default="")


def prompt_notes() -> str:
    """Prompt user for additional notes."""
    return Prompt.ask("[cyan]Notes[/cyan] (optional)", default="")


def prompt_confirm(question: str) -> bool:
    """Ask a yes/no question and return the bool result."""
    return Confirm.ask(question)


def print_success(message: str) -> None:
    console.print(f"[bold green]{message}[/bold green]")


def print_error(message: str) -> None:
    console.print(f"[bold red]Error:[/bold red] {message}")


def print_warning(message: str) -> None:
    console.print(f"[bold yellow]Warning:[/bold yellow] {message}")


def print_info(message: str) -> None:
    console.print(f"[cyan]{message}[/cyan]")
