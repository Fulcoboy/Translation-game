# Translation game (Frans → Nederlands)

A small Python quiz: it shows a random French word from an Excel file and you type the Dutch translation underneath.

## Setup

```bash
pip install -r requirements.txt
python translation_game.py            # uses words.xlsx
python translation_game.py other.xlsx # use a different file
```

## The Excel database

`words.xlsx` (first sheet), with a header row:

| Frans   | Nederlands         |
|---------|--------------------|
| bonjour | hallo / goedendag  |
| le chat | de kat             |

- Column A: French word, column B: Dutch translation.
- Multiple accepted answers: separate them with `/` or `;`.
- Answers are checked case-insensitively and ignore extra spaces.
- If the file does not exist, a sample file with 20 words is created.

Every word is asked once (in random order) before any repeats. Type `q` to stop and see your score plus the words to review.
