"""French -> Dutch vocabulary trainer that uses an Excel file as its database.

The Excel file (default: words.xlsx) must have two columns with a header row:

    | Frans   | Nederlands     |
    |---------|----------------|
    | bonjour | hallo / goedendag |
    | le chat | de kat         |

Several accepted Dutch answers can be separated with "/" or ";".
If the file does not exist yet, a sample file is created automatically.

Usage:
    python translation_game.py              # uses words.xlsx
    python translation_game.py mywords.xlsx # uses another file
"""

import random
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook

DEFAULT_DB = "words.xlsx"
FRENCH_HEADER = "Frans"
DUTCH_HEADER = "Nederlands"

SAMPLE_WORDS = [
    ("bonjour", "hallo / goedendag"),
    ("merci", "dank je / dank u / bedankt"),
    ("le chat", "de kat"),
    ("le chien", "de hond"),
    ("la maison", "het huis"),
    ("la voiture", "de auto / de wagen"),
    ("le livre", "het boek"),
    ("l'eau", "het water"),
    ("le pain", "het brood"),
    ("la pomme", "de appel"),
    ("rouge", "rood"),
    ("grand", "groot"),
    ("petit", "klein"),
    ("manger", "eten"),
    ("boire", "drinken"),
    ("aujourd'hui", "vandaag"),
    ("demain", "morgen"),
    ("l'école", "de school"),
    ("le jardin", "de tuin"),
    ("la fenêtre", "het raam / het venster"),
]


def create_sample_db(path: Path) -> None:
    """Create an Excel file with a header row and some example words."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Woorden"
    ws.append([FRENCH_HEADER, DUTCH_HEADER])
    for french, dutch in SAMPLE_WORDS:
        ws.append([french, dutch])
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 35
    wb.save(path)


def load_words(path: Path) -> list[tuple[str, list[str]]]:
    """Read (french, [accepted dutch answers]) pairs from the first sheet."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    words = []
    for row_number, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_number == 1:
            continue  # header row
        if not row or len(row) < 2 or row[0] is None or row[1] is None:
            continue  # skip empty / incomplete rows
        french = str(row[0]).strip()
        answers = [
            a.strip()
            for a in str(row[1]).replace(";", "/").split("/")
            if a.strip()
        ]
        if french and answers:
            words.append((french, answers))
    wb.close()
    return words


def normalize(text: str) -> str:
    """Make comparison forgiving for case and extra spaces."""
    return " ".join(text.lower().split())


def is_correct(answer: str, accepted: list[str]) -> bool:
    return normalize(answer) in {normalize(a) for a in accepted}


def play(words: list[tuple[str, list[str]]]) -> None:
    print("\n=== Frans -> Nederlands ===")
    print("Typ de Nederlandse vertaling onder het Franse woord.")
    print("Typ 'q' om te stoppen.\n")

    correct = 0
    asked = 0
    mistakes: list[tuple[str, list[str]]] = []
    deck: list[tuple[str, list[str]]] = []

    while True:
        # Go through all words in random order before repeating any.
        if not deck:
            deck = words[:]
            random.shuffle(deck)

        french, accepted = deck.pop()
        print(f"Frans:      {french}")
        try:
            answer = input("Nederlands: ")
        except (EOFError, KeyboardInterrupt):
            print()
            break

        if answer.strip().lower() == "q":
            break

        asked += 1
        if is_correct(answer, accepted):
            correct += 1
            print("✔ Juist!\n")
        else:
            mistakes.append((french, accepted))
            print(f"✘ Fout. Juist antwoord: {' / '.join(accepted)}\n")

    print("\n=== Resultaat ===")
    if asked == 0:
        print("Geen vragen beantwoord.")
        return
    print(f"Score: {correct}/{asked} ({correct / asked:.0%})")
    if mistakes:
        print("\nOm te herhalen:")
        seen = set()
        for french, accepted in mistakes:
            if french not in seen:
                seen.add(french)
                print(f"  {french} = {' / '.join(accepted)}")


def main() -> None:
    db_path = Path(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_DB)

    if not db_path.exists():
        create_sample_db(db_path)
        print(f"'{db_path}' bestond niet; een voorbeeldbestand is aangemaakt.")
        print("Open het in Excel om je eigen woorden toe te voegen.")

    words = load_words(db_path)
    if not words:
        print(f"Geen woorden gevonden in '{db_path}'.")
        print(f"Zet '{FRENCH_HEADER}' in kolom A en '{DUTCH_HEADER}' in kolom B.")
        sys.exit(1)

    print(f"{len(words)} woorden geladen uit '{db_path}'.")
    play(words)


if __name__ == "__main__":
    main()
