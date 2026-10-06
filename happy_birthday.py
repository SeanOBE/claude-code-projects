"""Create a "Happy Birthday" folder on the Desktop containing a text file saying hello Seán."""
import sys
from pathlib import Path


def find_desktop() -> Path:
    home = Path.home()
    for candidate in (home / "Desktop", home / "OneDrive" / "Desktop"):
        if candidate.is_dir():
            return candidate
    return home / "Desktop"


def main(desktop: Path) -> Path:
    folder = desktop / "Happy Birthday"
    folder.mkdir(parents=True, exist_ok=True)
    note = folder / "hello.txt"
    note.write_text("hello Seán", encoding="utf-8-sig")  # BOM so Notepad shows the á correctly
    return note


if __name__ == "__main__":
    print(main(Path(sys.argv[1]) if len(sys.argv) > 1 else find_desktop()))
