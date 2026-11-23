"""Prompts. Kept in one place so every question behaves the same way."""
import sys


def ask(prompt):
    try:
        return input(prompt).strip()
    except (EOFError, KeyboardInterrupt):
        print()
        sys.exit("Cancelled. Nothing was written.")


def confirm(prompt, assume_yes=False):
    if assume_yes:
        print(prompt + "y (--yes)")
        return True
    return ask(prompt).lower() in ("y", "yes")
