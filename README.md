# Bollywood Actor Filmography

Career studies of five Hindi-film actors. Each one is a seven-sheet Excel workbook generated from JSON research files. The focus is the actor's romantic screen persona: who he was paired with, how each relationship works, and how that changes across his career.

The JSON under `research/` is the source of truth. The workbooks are built from it and are never edited by hand.

## Actors

| Actor | Films | Release years | Romances | Collaborators | Cast rows | Awards | Sources |
| --- | ---: | --- | ---: | ---: | ---: | ---: | ---: |
| Aamir Khan | 57 | 1973–2026 | 37 | 47 | 169 | 143 | 121 |
| Shah Rukh Khan | 97 | 1992–2023 | 68 | 79 | 342 | 291 | 186 |
| Amitabh Bachchan | 260 | 1969–2025 | 95 | 306 | 1,332 | 266 | 890 |
| Shashi Kapoor | 154 | 1948–1998 | 76 | 374 | 777 | 156 | 385 |
| Rishi Kapoor | 160 | 1955–2022 | 90 | 561 | 1,116 | 294 | 468 |

Counts are as of 8 October 2026. "Films" means released feature films the actor appears in, including cameos, voice roles and child roles. Each actor's `meta.json` has a `scope_note` listing what was left out and why.

## The workbook

Seven sheets, joined by stable IDs:

| Sheet | One row per | ID |
| --- | --- | --- |
| Overview | (dashboard: career totals, score averages, career phases, charts) | |
| Films | film | `FILM-001` |
| Romances | meaningful romantic pairing | `PAIR-001` |
| Collaborators | director, producer, banner, composer or major writer | `PERSON-001` |
| Ensemble Cast | supporting performance | `CAST-001` |
| Awards | award win or nomination | `AWARD-001` |
| Sources | source | `SOURCE-001` |

`Romances` is the main analytical sheet. Each pairing is scored from 0 to 10 on pair friction, external opposition, lover-boy characterization, romance centrality, relationship health, mutuality, female agency and chemistry, and each score has a written explanation beside it. A pairing gets a row only when the relationship matters to the story, not because an actress is the top-billed woman.

Two of those scores are kept deliberately separate. Pair friction is conflict that starts between the partners. External opposition is whatever is imposed on them from outside: family, class, a villain, fate. For all five actors the second runs well above the first:

| Actor | Pair friction | External opposition | Lover-boy |
| --- | ---: | ---: | ---: |
| Aamir Khan | 3.86 | 6.22 | 5.54 |
| Shah Rukh Khan | 4.34 | 6.26 | 5.34 |
| Amitabh Bachchan | 3.23 | 6.11 | 3.79 |
| Shashi Kapoor | 3.62 | 6.26 | 4.70 |
| Rishi Kapoor | 3.39 | 6.81 | 5.17 |

These are means over each actor's `Romances` rows.

## Layout

```
research/<actor-slug>/
  meta.json                   actor, research cutoff, scope note, career phases
  films.json                  one JSON array per sheet, one object per row
  romances.json
  collaborators.json
  ensemble_cast.json
  awards.json
  sources.json
  collaborator_roster.json    frozen list that fixes the PERSON ids
  extract_config.json         scope decisions, kept so extraction can be re-run
<Actor Name> Filmography.xlsx generated workbooks, one per actor
.claude/skills/bollywood-actor-filmography/
  SKILL.md                    the full specification
  scripts/                    build, extraction and merge scripts
```

`collaborator_roster.json` and `extract_config.json` exist only for Amitabh Bachchan, Shashi Kapoor and Rishi Kapoor. The Aamir Khan and Shah Rukh Khan runs predate the scripts that use them.

`SKILL.md` is the reference for everything this README summarizes: the column list for every sheet, the controlled vocabularies, the scoring rubrics, the commercial-data rules and the research workflow.

## Rebuilding a workbook

The build script needs Python with openpyxl. The current workbooks were built with Python 3.14 and openpyxl 3.1.5.

```bash
python3 -m venv ~/.venvs/filmography
~/.venvs/filmography/bin/pip install openpyxl
```

Then, from the repository root:

```bash
PY=~/.venvs/filmography/bin/python
SCRIPT=.claude/skills/bollywood-actor-filmography/scripts/build_workbook.py

$PY $SCRIPT validate research/shashi-kapoor
$PY $SCRIPT build research/shashi-kapoor -o "Shashi Kapoor Filmography.xlsx"
$PY $SCRIPT schema romances     # JSON keys and allowed values for one sheet
```

`validate` checks IDs, cross-sheet references, controlled vocabularies, score ranges, duplicate films and basic consistency of the commercial figures. `build` runs the same checks and then writes the workbook, including the Overview formulas and charts.

To change anything in a workbook, edit the JSON and rebuild.

All five actors validate with 0 errors. Each also reports between 5 and 13 warnings. These are known gaps in the research, not defects: films with no Wikipedia article and so no director or language, chemistry scores left blank, award rows with no ceremony year.

## Adding an actor

The research workflow is a [Claude Code](https://claude.com/claude-code) skill. Open the repository in Claude Code and run `/bollywood-actor-filmography` with the actor's name. A full career is split into five phases, one per session:

1. Filmography, Film IDs and the fields that can be extracted mechanically.
2. Films and romances, in batches of about 30 films.
3. Collaborators.
4. Ensemble cast.
5. Awards, sources, `validate` and `build`.

The other scripts in `scripts/` support those phases. `extract_wiki.py` fills mechanical fields from Wikipedia wikitext, and the `merge_*.py` scripts apply a batch of hand-written rows to the research files. `SKILL.md` documents each one.

## Limits

* **The scores are interpretive.** They are readings of each film against a written rubric, not measurements.
* **Chemistry is not comparable across all five actors.** For Aamir Khan and Shah Rukh Khan it is scored only where a source that was read assesses the pairing or how it was received: 11 of 37 pairings and 7 of 68. The rest are left blank with a "Not scored" note. For the other three every pairing is scored, and most of those scores have no such source behind them: 77 of 95 for Amitabh Bachchan, 68 of 76 for Shashi Kapoor and 88 of 90 for Rishi Kapoor. Those rows begin their chemistry explanation with "Interpretive reading".
* **Most sources are Wikipedia.** Of the 2,050 source rows, 1,596 are Wikipedia articles. The `Sources` sheet grades each source from A to D: 1,600 are tier C, 284 tier B and 166 tier D, and none is tier A. Nearly all the tier B press and trade sourcing is in the Amitabh Bachchan and Shashi Kapoor workbooks.
* **Box-office data is thin for older films.** Shashi Kapoor has no worldwide gross for any of his 154 films, and Rishi Kapoor has one for 11 of 160. Figures are nominal and in the currency the source reported.
* **Overview cells can look blank in a previewer.** openpyxl saves formulas without calculated values, so file previewers such as Quick Look and Google Drive may show the Overview figures empty. A spreadsheet application such as Excel or LibreOffice calculates them when it opens the file.
