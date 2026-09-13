---
name: bollywood-actor-filmography
description: Research a male Bollywood actor’s complete released acting filmography and create a seven-sheet Excel workbook focused on his roles, romantic pairings, relationship dynamics, collaborators, ensemble casts, commercial performance, awards, and sources. Use for complete career studies or updates to an existing actor-filmography workbook.
metadata:
  short-description: Build a Bollywood actor filmography workbook
---

# Bollywood Actor Filmography

Create or update a professionally formatted `.xlsx` workbook examining a selected male Bollywood actor’s complete released acting filmography.

Record the research as JSON files and generate the workbook with the bundled build script. Never hand-edit the `.xlsx`; change the JSON and rebuild.

## Build tooling

The script is `scripts/build_workbook.py` in this skill's directory. Run it with the interpreter that has openpyxl installed:

```bash
PY=~/.venvs/filmography/bin/python
SCRIPT=.claude/skills/bollywood-actor-filmography/scripts/build_workbook.py
$PY $SCRIPT init research/<actor-slug> --actor "<Actor Name>" --cutoff YYYY-MM-DD
$PY $SCRIPT schema films            # JSON keys, types and allowed values for one sheet
$PY $SCRIPT validate research/<actor-slug>
$PY $SCRIPT build research/<actor-slug> -o "<Actor Name> Filmography.xlsx"
```

Research files, one directory per actor:

* `meta.json`: `actor`, `research_cutoff`, `scope_note`, and `career_phases` (a list of objects with `phase`, `start_year`, `end_year`, `summary`)
* `films.json`, `romances.json`, `collaborators.json`, `ensemble_cast.json`, `awards.json`, `sources.json`: JSON arrays with one object per row, using the keys printed by `schema`
* `unreleased.json` (optional): unreleased or abandoned films, only when the user asks for them; IDs `UNREL-001`; no commercial figures

Data conventions:

* Omit a key or use `null` when a value is unknown. Never write placeholders such as "N/A" or "TBD".
* Separate multiple values with `; ` or use a JSON array (Source IDs, Film IDs, names, opposition sources).
* Scores are numbers from 0 to 10. Years are numbers. Dates are `YYYY-MM-DD`, or `YYYY-MM` / `YYYY` when only that is known.
* Money is a number in the unit named in `currency` (for example `"currency": "INR crore"`), or text for a range or disputed figure (`"₹60–65 crore"`).
* Leave derived values blank: `film`, `release_year` and the actor's character in linked sheets, `release_year` and `decade` in `Films`, and `collaboration_count` are filled from the other data.

The script produces the sheet order, headers, number formats, frozen panes, filters, data validation, score colour scales, hyperlinks, row heights, chronological sorting, and the Overview formulas and charts. It checks IDs, cross-sheet references, controlled vocabularies, score ranges, duplicate films and basic commercial-data consistency. Fix every error and review every warning before delivery.

The project’s central analytical focus is the actor’s romantic screen persona, especially:

* Principal female co-stars
* Nature and development of each romantic relationship
* Pair friction
* External opposition
* Lover-boy characterization
* Romance centrality
* Relationship health
* Female-character agency
* Changes in these patterns across the actor’s career

## Required input

Determine or obtain:

* Actor’s name
* Career cutoff date, if specified
* Whether to include only released films or also unreleased projects
* Whether the request is for a full career, selected period or workbook update

If the user names only the actor, default to all released feature-film acting appearances through the current date.

Do not ask for details that can be reasonably inferred.

## Filmography scope

Include released feature films in which the actor has:

* A lead role
* A co-lead or ensemble-lead role
* A supporting role
* An antagonistic role
* An extended special appearance
* A cameo
* A performance as himself
* A voice role or narration
* A child-actor role
* A role in a language other than Hindi
* A role in one segment of an anthology

Identify the role type explicitly.

Exclude by default:

* Television programs
* Advertisements
* Standalone music videos
* Documentaries featuring only interview footage
* Films produced or directed by the actor in which he does not act
* Announced projects that have not been released
* Unverified rumored appearances

Record significant unreleased or abandoned films in a clearly marked section at the bottom of the `Films` sheet only when the user requests them. Never mix their commercial figures with released films.

Use theatrical release year as the primary year. Preserve alternate or disputed dates in a note.

## Workbook contract

Create exactly these seven sheets, in this order:

1. `Overview`
2. `Films`
3. `Romances`
4. `Collaborators`
5. `Ensemble Cast`
6. `Awards`
7. `Sources`

Use stable IDs to connect sheets:

* `FILM-001`
* `PAIR-001`
* `PERSON-001`
* `CAST-001`
* `AWARD-001`
* `SOURCE-001`

One film must have only one master row in `Films`. A film may have multiple rows in `Romances`, `Ensemble Cast`, `Awards`, and `Sources`.

## Sheet 1: Overview

Create a concise career dashboard supported by formulas or calculated source data from the other sheets.

Include:

* Actor
* Career period
* Research cutoff date
* Total released acting films
* Lead and co-lead films
* Supporting roles
* Cameos and special appearances
* Films with central romance
* Films with secondary romance
* Films without meaningful romance
* Distinct principal female co-stars
* Most frequent romantic co-star
* Most frequent director
* Most frequent producer or production banner
* Most frequent music director
* Average pair-friction score
* Average external-opposition score
* Average lover-boy score
* Average romance-centrality score
* Highest-grossing film, using comparable worldwide-gross data
* Major awards total
* Source-confidence summary

Add a compact career-phase analysis, such as:

* Early romantic-hero phase
* Experimental or transitional phase
* Mature lead phase
* Character-actor or prestige phase

Include only useful visualizations. Prefer:

* Lover-boy score by release year
* Actress collaboration frequency
* Pair friction versus external opposition
* Romance centrality by decade

Do not overload the dashboard with decorative charts.

## Sheet 2: Films

Use one row per film.

Required columns:

* Film ID
* Title
* Release Date
* Release Year
* Decade
* Language
* Genre
* Secondary Genre
* Runtime
* Director
* Producer or Producers
* Production Banner
* Writer or Writers
* Source Material or Remake Status
* Song Composer
* Background-Score Composer
* Lyricist
* Actor’s Character
* Alternate Identity or Double Role
* Role Type
* Billing Position
* Character Occupation
* Character Social Background
* Role Gist
* Principal Motivation
* Principal Conflict
* Character Arc
* Moral Alignment
* Character Outcome
* Principal Female Lead or Leads
* Romance Presence
* Romance Centrality Score
* Overall Relationship Classification
* Budget
* Domestic Nett
* Domestic Gross
* Overseas Gross
* Worldwide Gross
* Currency
* Box-Office Verdict
* Annual Box-Office Rank
* Critical Reception Summary
* Recognition or Legacy
* Source IDs
* Confidence Grade
* Research Notes

Keep `Role Gist` to approximately 40–80 words. Describe the actor’s character and arc, not merely the general film plot.

Use these `Role Type` values consistently:

* Lead
* Co-lead
* Ensemble lead
* Supporting
* Antagonist
* Extended special appearance
* Cameo
* Self appearance
* Voice role
* Narrator
* Child role

Use these `Romance Presence` values:

* None
* Implied
* Minor
* Secondary
* Major
* Central

Use these broad relationship classifications when applicable:

* Frictionless idealized romance
* United couple versus external world
* Friendship to love
* Rivals to lovers
* Playful conflict masking attraction
* Genuine incompatibility gradually resolved
* Established relationship
* Marriage under strain
* Relationship damaged by mistrust
* Relationship manipulated by outsiders
* Forbidden love
* Unrequited love
* Love triangle
* Tragic lovers
* Unequal or coercive relationship
* Obsessive pursuit
* Romance subordinate to another genre
* Multiple romantic relationships
* No meaningful romance

## Sheet 3: Romances

This is the project’s primary analytical sheet.

Use one row per meaningful romantic pairing. If the actor has multiple romantic interests in one film, create a separate row for each.

Do not create a pairing row merely because an actress is the highest-billed woman. A narratively meaningful romantic or marital connection must exist.

Required columns:

* Pairing ID
* Film ID
* Film
* Release Year
* Actor’s Character
* Actress
* Actress’s Character
* Actress Role Prominence
* Principal Female Lead
* Relationship Category
* Initial Relationship
* How They Meet
* Who Initiates
* Courtship Pattern
* Emotional Dynamic
* Pair’s Principal Conflict
* Pair Friction Score
* Pair Friction Explanation
* External Opposition Score
* External Opposition Source
* External Opposition Explanation
* Lover-Boy Score
* Lover-Boy Explanation
* Romance Centrality Score
* Relationship Health Score
* Mutuality Score
* Female Agency Score
* Chemistry Score
* Chemistry Explanation
* Playful Banter
* Serious Conflict
* Jealousy or Possessiveness
* Deception or Mistaken Identity
* Separation or Estrangement
* Sacrifice
* Couple Functions as Team
* Romantic Transformation of Actor’s Character
* Transformation of Female Character
* Relationship Outcome
* Marriage or Long-Term Union
* Happy, Tragic or Unresolved Ending
* Major Romantic Songs
* Relationship Analysis
* Contemporary Critical Reading
* Source IDs
* Confidence Grade

Use `Yes`, `No`, `Partial`, `Unclear`, or `Not applicable` consistently for categorical relationship fields.

### Pair friction: 0–10

Measure conflict originating between the partners. Do not count purely external obstacles.

* `0`: Completely aligned; no meaningful interpersonal conflict
* `1–2`: Negligible disagreement or brief hesitation
* `3–4`: Mild personality conflict or temporary misunderstanding
* `5–6`: Recurring arguments, conflicting goals or substantial mistrust
* `7–8`: Strong antagonism, serious rupture or damaging distrust
* `9`: Relationship dominated by hostility, coercion or betrayal
* `10`: Direct enemies or fundamentally destructive relationship

Distinguish playful banter from serious friction.

### External opposition: 0–10

Measure obstacles imposed from outside the couple.

* `0`: No meaningful external obstacle
* `1–2`: Mild inconvenience or disapproval
* `3–4`: Noticeable family, social or circumstantial resistance
* `5–6`: Sustained opposition affecting the relationship
* `7–8`: Severe family, class, social, criminal or political interference
* `9`: Forced separation, persecution or extreme danger
* `10`: External opposition causes or nearly causes death, permanent separation or tragedy

Identify the source:

* Family
* Class
* Caste
* Religion
* Community
* Existing relationship
* Villain
* Crime
* War or politics
* Illness
* Distance
* Fate
* Other

### Actor as lover boy: 0–10

Measure how strongly the role conforms to the lover-boy archetype: a romantic male protagonist whose identity, motivation and appeal are substantially organized around courtship, devotion or sacrifice for love.

* `0`: No romantic identity
* `1–2`: Incidental romantic interest
* `3–4`: Romance is present but another function dominates
* `5–6`: Significant romantic heroism shared with other major functions
* `7–8`: Clearly a romantic hero
* `9`: Romance is the character’s overwhelming motivation
* `10`: Defining pure lover-boy role

Consider:

* Romantic screen time
* Courtship activity
* Emotional openness
* Devotion
* Rebellion or sacrifice for love
* Romantic songs
* Whether the role shaped the actor’s romantic-star image

Do not equate lover-boy score with romance centrality. A film can centre on a relationship while portraying the man as a husband, obsessive pursuer or emotionally unavailable partner rather than a lover boy.

### Additional scores

Use 0–10 for:

* `Romance Centrality`: importance of romance to the film
* `Relationship Health`: consent, respect, trust and emotional safety
* `Mutuality`: comparable emotional participation by both partners
* `Female Agency`: independent goals, choices and narrative consequences
* `Chemistry`: credibility or appeal of the performers’ interaction

Treat chemistry as an interpretive assessment and provide a short explanation.

Evaluate coercion, stalking, possessiveness and deception from both:

1. The film’s intended contemporary framing
2. A modern critical perspective

## Sheet 4: Collaborators

Include directors, principal producers, production banners, song composers and background-score composers. Include major writers when their contribution is particularly important.

Use one row per collaborator or composing team.

Required columns:

* Person ID
* Name
* Profession or Function
* Birth Year
* Death Year
* Years Active
* Industry or Primary Language
* Career Background
* Debut or Early Breakthrough
* Major Works
* Recurring Genres or Themes
* Creative Style
* Major Awards
* Historical Importance
* Relationship with Selected Actor
* Number of Collaborations
* Collaboration Film IDs
* Most Important Collaboration
* Extended Biography
* Source IDs
* Confidence Grade

Write approximately:

* 150–300 words for directors
* 100–250 words for principal producers
* 150–300 words for music directors or composing teams
* 100–200 words for important writers

Do not create long biographies for minor contractual co-producers unless their involvement is historically relevant.

Separate song composers from background-score composers when they differ.

## Sheet 5: Ensemble Cast

Include major actors who materially affect the plot, protagonist, romance or cultural identity of the film.

Use one row per supporting performance.

Required columns:

* Cast ID
* Film ID
* Film
* Release Year
* Actor
* Character
* Cast Function
* Importance Level
* Relationship to Protagonist
* Relationship to Romantic Plot
* Character Gist
* Performance Recognition
* Source IDs

Use these `Cast Function` values where appropriate:

* Co-lead
* Supporting
* Antagonist
* Mentor
* Family
* Friend
* Comic support
* Romantic rival
* Authority figure
* Cameo
* Self appearance

Use `Major`, `Medium`, or `Minor` for importance.

Do not reproduce the complete end credits. Include the meaningful ensemble.

## Sheet 6: Awards

Use one row per award result.

Required columns:

* Award ID
* Film ID
* Film
* Ceremony Year
* Award Organization
* Category
* Recipient
* Recipient Type
* Win or Nomination
* Competitive or Honorary
* Notes
* Source IDs
* Confidence Grade

Cover, where applicable:

* National Film Awards
* Filmfare Awards
* Major critics’ awards
* International festivals
* Academy Award submissions or nominations
* Performance awards
* Direction and writing awards
* Music and playback-singing awards
* Important technical awards

Do not describe a nomination as a win. Keep film-festival selection separate from an award.

Record non-award recognition—cult status, re-releases, major critical lists or social influence—in `Films`, not here.

## Sheet 7: Sources

Use one row per source.

Required columns:

* Source ID
* Film IDs Supported
* Pairing IDs Supported
* Source Title
* Publisher
* Author
* Publication Date
* URL
* Access Date
* Claims Supported
* Source Type
* Reliability Tier
* Notes

Use these reliability tiers:

* `A`: Official credits, government records, award bodies, primary interviews
* `B`: Reputable newspapers, magazines, trade publications or academic sources
* `C`: Established reference databases used with corroboration
* `D`: Fan-maintained or weak secondary sources used only when stronger evidence is unavailable

Use Wikipedia and IMDb for orientation and cross-checking, not as the sole source for disputed commercial, biographical or award claims.

## Commercial-data rules

Never conflate:

* Budget
* Domestic nett
* Domestic gross
* Overseas gross
* Worldwide gross
* Distributor share

Keep the original reported currency and nominal amount.

If sources disagree:

* Preserve the important competing figures in `Research Notes`.
* Use a range when appropriate.
* Explain whether one figure is nett and another is gross.
* Assign a lower confidence grade.
* Do not invent an exact compromise figure.

Use inflation-adjusted values only if requested. Keep them separate from nominal figures and explain the adjustment method.

Use these confidence grades:

* `High`: supported by authoritative or multiple strong sources
* `Medium`: credible but incompletely corroborated
* `Low`: estimated, disputed or supported only by weak sources
* `Unknown`: no reliable data located

## Research workflow

For a new actor:

1. Establish the complete released acting filmography.
2. Reconcile alternate titles, delayed releases, languages and cameo claims.
3. Assign Film IDs chronologically.
4. Research film and role facts.
5. Create separate romance records.
6. Apply the scoring rubrics consistently.
7. Add recurring collaborators and extended biographies.
8. Add meaningful ensemble members.
9. Add awards and commercial information.
10. Populate sources.
11. Run `validate`; fix every error and review every warning.
12. Run `build`. It generates the Overview from the detail data, re-opens the saved workbook for a structural check, and prints the value each Overview formula should produce.
13. openpyxl saves formulas without calculated values. Spreadsheet applications calculate them on open, but file previewers may show those cells blank. When a spreadsheet application or LibreOffice is available, open the workbook for a visual check and confirm the Overview matches the build report.

For a large filmography, use a pilot when the user requests one or when the scoring model has not previously been validated. Select contrasting films:

* Pure romance
* Romantic comedy
* Action or mission-driven film with secondary romance
* Ensemble film
* Film without meaningful romance

After the pilot, continue to the complete filmography unless the user asked to review the pilot first.

## Analytical standards

Distinguish:

* Pair friction from external opposition
* Romantic pairing from merely having a female co-star
* Lover-boy identity from general romantic involvement
* Actress billing from narrative importance
* Playful hostility from genuine incompatibility
* Film-intended romance from ethically healthy romance
* Actor chemistry from character relationship health
* Commercial success from artistic or cultural recognition
* Production year from theatrical release year
* Fact from interpretation

Do not force romance into films where it is absent.

If a character has a spouse who appears only briefly, record the relationship in `Films`; create a `Romances` row only if the relationship has meaningful narrative or character value.

For multiple heroines or love triangles, create one pairing row per woman and explain the relationships independently.

## Workbook presentation

Create a professional research workbook with:

* Frozen header rows
* Filters on all data tables
* Wrapped text for narrative fields
* Appropriate column widths
* Consistent number formats
* Distinct formatting for facts, scores and interpretive analysis
* Conditional formatting for 0–10 scores
* Hyperlinked source URLs
* Chronological sorting in `Films`
* Clear data validation for repeated categories
* No clipped headings or narrative cells
* No merged cells inside data tables

Use formulas for Overview metrics where practical. Ensure formulas contain no obvious errors.

Keep the workbook readable. Extended prose belongs in designated narrative columns rather than being repeated across multiple sheets.

## Final validation

Before delivery, verify:

* Every released acting film is represented once in `Films`.
* Each meaningful romantic interest has a separate `Romances` row.
* Films without romance are not given artificial pairing records.
* Pair friction excludes purely external conflict.
* Lover-boy scores follow the same standard across decades.
* Major collaborators have biographies.
* Major supporting actors are included without dumping full credits.
* Budget and gross measures are correctly labeled.
* Disputed data has confidence grades and notes.
* Award wins and nominations are distinguished.
* Source IDs resolve correctly.
* Overview figures reconcile with the detail sheets.
* Important formulas calculate correctly.
* The workbook passes visual inspection.

Deliver the final `.xlsx` workbook with a short summary of its scope, important limitations and the most significant preliminary pattern found.
