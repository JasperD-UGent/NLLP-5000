# NLLP-5000 (compilation)

## Prerequisites

To run the `process_input_data.py` script, you need access to the original Routledge dictionary (more specifically, to the `Frequency-index.txt` file that you can download after having purchased the dictionary). The dictionary can be purchased on the [Routledge webshop](https://www.routledge.com/A-Frequency-Dictionary-of-Dutch-Core-Vocabulary-for-Learners/Tiberius-Schoonheim/p/book/9780415523806). To download the `Frequency-index.txt` file, follow the instructions in the "Support material" section on that page.

**IMPORTANT**: As explained in the NLP4CALL paper (see main `README.md`), the `Frequency-index.txt` file contains a few errors and typos. These issues will need to be manually fixed before you can run the script. All necessary edits are listed in Appendix A of the paper.

To run the `build_nllp5000_resource.py` script, you first need to annotate (i.e. verify and complete) the unverified XLSX spreadsheet returned by the `process_input_data.py` script (located in the automatically created `output/inputData_procsd` folder).

## Requirements

### Python

See `requirements_Py313.txt`. Tested with Python 3.13.5. The requirements can be installed automatically in one go by running the following command in your shell (preferably in a fresh virtual environment):

```shell
$ python -m pip install -r requirements_Py313.txt
```

## Spreadsheet column names

Below you can find a detailed overview and description of the columns used in the spreadsheet returned by the `process_input_data.py` script, presented per part of speech (each part of speech has a separate sheet).

Abbreviations and symbols used in the tables:

- N/A: not applicable (e.g., used in "Stands for ..." column when "Column name" is self-explanatory)
- Upward arrow: relevant information is specified in the rows above

Detailed information on the UD POS tags and morphological features can be found on the following web pages:

- POS tags: <https://universaldependencies.org/u/pos/index.html>
- Morphological features: <https://universaldependencies.org/u/feat/index.html>

### Adjectives (`adj`)

| Column name | Stands for ... | Description | UD POS tag | UD morphological features |
| --- | --- | --- | --- | --- |
| ST_standaard | Stellende trap (standaard) | Positive, first degree; uninflected, standard form | ADJ | Degree=Pos |
| ST_verbogen | Stellende trap (verbogen) | Positive, first degree; inflected form with *buigings-e* | ADJ | Degree=Pos |
| ST_iets | Stellende trap (iets) | Positive, first degree; inflected form with *buigings-s* needed when the adjective is combined with "iets" | ADJ | Degree=Pos |
| VT_standaard | Vergrotende trap (standaard) | Comparative, second degree; uninflected, standard form | ADJ | Degree=Cmp |
| VT_verbogen | Vergrotende trap (verbogen) | Comparative, second degree; inflected form with *buigings-e* | ADJ | Degree=Cmp |
| VT_iets | Vergrotende trap (iets) | Comparative, second degree; inflected form with *buigings-s* needed when the adjective is combined with "iets" | ADJ | Degree=Cmp |
| OT_standaard | Overtreffende trap (standaard) | Superlative, third degree; uninflected, standard form | ADJ | Degree=Sup |
| OT_verbogen | Overtreffende trap (verbogen) | Superlative, third degree; inflected form with *buigings-e* | ADJ | Degree=Sup |

### Adverbs (`adv`)

| Column name | Stands for ... | Description | UD POS tag | UD morphological features |
| --- | --- | --- | --- | --- |
| regulier | N/A | Column indicating if the adverb is a regular adverb | ADV | None |
| negatief | N/A | Column indicating if the adverb encodes negation | PART (if the word is a particle) or ADV (if the word is a regular adverb) | Polarity=Neg (if the word is a particle) or PronType=Neg (if the word is a regular adverb) |
| partikel | N/A | Column indicating if the adverb is a particle | PART (if the word encodes negation) or ADP (if the word does not encode negation, in which case it is a verbal particle) | PartType=Neg (if the word encodes negation) or PartType=Vbp (if the word is a verbal particle) |
| onveranderlijke_vorm | N/A | Fixed form of the adverb | &uparrow; | &uparrow; |

### Articles (`art`)

| Column name | Stands for ... | Description | UD POS tag | UD morphological features |
| --- | --- | --- | --- | --- |
| bepaald | N/A | Column indicating if the article is definite | DET | Definite=Def \| PronType=Art |
| onbepaald | N/A | Column indicating if the article is indefinite | DET | Definite=Ind \| PronType=Art |
| nominatief | N/A | Uninflected, standard nominative form | &uparrow; | Case=Nom |
| accusatief | N/A | Uninflected, standard accusative form | &uparrow; | Case=Acc |
| genitief | N/A | Uninflected, standard genitive form | &uparrow; | Case=Gen |
| datief | N/A | Uninflected, standard dative form | &uparrow; | Case=Dat |
| verbogen | N/A | Semicolon-separated string of inflected forms per case (e.g., "nominatief=ene;accusatief=ene") | &uparrow; | (see nominatief / accusatief / genitief / datief) |

### Conjunctions (`conj`)

| Column name | Stands for ... | Description | UD POS tag | UD morphological features |
| --- | --- | --- | --- | --- |
| nevenschikkend | N/A | Column indicating if the conjunction is coordinating | CCONJ | None |
| onderschikkend | N/A | Column indicating if the conjunction is subordinating | SCONJ | None |
| onveranderlijke_vorm | N/A | Fixed form of the conjunction | &uparrow; | &uparrow; |

### Interjections (`interj`)

| Column name | Stands for ... | Description | UD POS tag | UD morphological features |
| --- | --- | --- | --- | --- |
| onveranderlijke_vorm | N/A | Fixed form of the interjection | INTJ | None |

### Nouns (`noun`)

| Column name | Stands for ... | Description | UD POS tag | UD morphological features |
| --- | --- | --- | --- | --- |
| M | Mannelijk | Column indicating if the noun is masculine | NOUN | Gender=Masc |
| O | Onzijdig | Column indicating if the noun is neuter | NOUN | Gender=Neut |
| V | Vrouwelijk | Column indicating if the noun is feminine | NOUN | Gender=Fem |
| M-O | Mannelijk-onzijdig | Column indicating if the noun is both masculine and neuter | NOUN | Gender=Masc,Neut |
| M-O-V | Mannelijk-onzijdig-vrouwelijk | Column indicating if the noun is masculine, neuter, as well as feminine | NOUN | Gender=Fem,Masc,Neut |
| M-V | Mannelijk-vrouwelijk | Column indicating if the noun is both masculine and feminine | NOUN | Gender=Fem,Masc |
| O-V | Onzijdig-vrouwelijk | Column indicating if the noun is both neuter and feminine | NOUN | Gender=Fem,Neut |
| eigennaam | N/A | Column indicating if the noun is a proper noun | PROPN | &uparrow; |
| standaardvorm_ENK | standaardvorm (enkelvoud) | Standard form in singular | &uparrow; | Number=Sing |
| standaardvorm_MV | standaardvorm (meervoud) | Standard form in plural | &uparrow; | Number=Plur |
| verkleinwoord_ENK | verkleinwoord (enkelvoud) | Diminutive form in singular | &uparrow; | Degree=Dim \| Number=Sing |
| verkleinwoord_MV | verkleinwoord (meervoud) | Diminutive form in plural | &uparrow; | Degree=Dim \| Number=Plur |

### Numerals (`num`)

| Column name | Stands for ... | Description | UD POS tag | UD morphological features |
| --- | --- | --- | --- | --- |
| hoofdtelwoord | N/A | Column indicating if the numeral is a cardinal number | NUM | NumType=Card |
| rangtelwoord | N/A | Column indicating if the numeral is an ordinal number | NUM | NumType=Ord |
| bepaald | N/A | Column indicating if the numeral is definite | &uparrow; | Definite=Def |
| onbepaald | N/A | Column indicating if the numeral is indefinite | &uparrow; | Definite=Ind |
| standaardvorm | N/A | Uninflected, standard form of the numeral | &uparrow; | &uparrow; |
| verbogen_e | N/A | Inflected form with *buigings-e* | &uparrow; | &uparrow; |
| verbogen_n | N/A | Inflected form with *buigings-n* | &uparrow; | &uparrow; |

### Prepositions (`prep`)

| Column name | Stands for ... | Description | UD POS tag | UD morphological features |
| --- | --- | --- | --- | --- |
| achterzetsel | N/A | Column indicating if the word is a postposition | ADP | AdpType=Post |
| omzetsel | N/A | Column indicating if the word is a circumposition | ADP | AdpType=Circ |
| voorzetsel | N/A | Column indicating if the word is a preposition | ADP | AdpType=Prep |
| onveranderlijke_vorm | N/A | Fixed form of the word | &uparrow; | &uparrow; |

### Pronouns (`pron`)

| Column name | Stands for ... | Description | UD POS tag | UD morphological features |
| --- | --- | --- | --- | --- |
| aanwijzend | N/A | Column indicating if the word is a demonstrative pronoun | PRON | PronType=Dem |
| betrekkelijk | N/A | Column indicating if the word is a relative pronoun | PRON | PronType=Rel |
| bezittelijk | N/A | Column indicating if the word is a possessive pronoun | PRON | Poss=Yes \| PronType=Prs |
| onbepaald | N/A | Column indicating if the word is an indefinite pronoun | PRON | PronType=Ind |
| persoonlijk | N/A | Column indicating if the word is a personal pronoun | PRON | PronType=Prs |
| wederkerend | N/A | Column indicating if the word is a reflexive pronoun | PRON | PronType=Prs \| Reflex=Yes |
| wederkerig | N/A | Column indicating if the word is a reciprocal pronoun | PRON | PronType=Rcp |
| vragend | N/A | Column indicating if the word is an interrogative pronoun | PRON | PronType=Int |
| getal | N/A | Number of the word (if applicable) | &uparrow; | Number=Sing or Number=Plur |
| persoon | N/A | Person of the word (if applicable) | &uparrow; | Person=1 or Person=2 or Person=3 |
| geslacht | N/A | Gender of the word (if applicable) | &uparrow; | Gender=Fem or Gender=Masc or Gender=Neut or Gender=Fem,Masc or Gender=Fem,Neut or Gender=Fem,Masc,Neut or Gender=Masc,Neut |
| nominatief | N/A | Uninflected, standard nominative form | &uparrow; | Case=Nom |
| accusatief | N/A | Uninflected, standard accusative form | &uparrow; | Case=Acc |
| genitief | N/A | Uninflected, standard genitive form | &uparrow; | Case=Gen |
| datief | N/A | Uninflected, standard dative form | &uparrow; | Case=Dat |
| verbogen_e | N/A | Semicolon-separated string of inflected forms with *buigings-e* per case | &uparrow; | (see nominatief / accusatief / genitief / datief) |
| verbogen_n | N/A | Semicolon-separated string of inflected forms with *buigings-n* per case | &uparrow; | (see nominatief / accusatief / genitief / datief) |

### Verbs (`verb`)

| Column name | Stands for ... | Description | UD POS tag | UD morphological features |
| --- | --- | --- | --- | --- |
| zelfstandig | N/A | Column indicating if the verb is a lexical verb | VERB | None |
| hulp_regulier | N/A | Column indicating if the verb is an auxiliary verb (regular) | AUX | None |
| hulp_koppel | N/A | Column indicating if the verb is an auxiliary verb (copula) | AUX | VerbType=Cop |
| hulp_modaal | N/A | Column indicating if the verb is an auxiliary verb (modal) | AUX | VerbType=Mod |
| INF | Infinitive | Infinitive form | &uparrow; | VerbForm=Inf |
| TT_1-2_inversie | Tegenwoordige tijd 1e persoon enkelvoud en 2e persoon enkelvoud met inversie | Present tense first person singular and second person singular with inversion | &uparrow; | Mood=Ind \| Number=Sing \| Person=1 \| Tense=Pres \| VerbForm=Fin and Mood=Ind \| Number=Sing \| Person=2 \| Tense=Pres \| VerbForm=Fin |
| TT_2_je | Tegenwoordige tijd 2e persoon enkelvoud in je-vorm | Present tense second person singular in *je* form | &uparrow; | Mood=Ind \| Number=Sing \| Person=2 \| Tense=Pres \| VerbForm=Fin |
| TT_3 | Tegenwoordige tijd 3e persoon enkelvoud | Present tense third person singular | &uparrow; | Mood=Ind \| Number=Sing \| Person=3 \| Tense=Pres \| VerbForm=Fin |
| TT_4-5-6 | Tegenwoordige tijd 1e, 2e en 3e persoon meervoud | Present tense first, second, and third person plural | &uparrow; | Mood=Ind \| Number=Plur \| Person=1 \| Tense=Pres \| VerbForm=Fin and Mood=Ind \| Number=Plur \| Person=2 \| Tense=Pres \| VerbForm=Fin and Mood=Ind \| Number=Plur \| Person=3 \| Tense=Pres \| VerbForm=Fin |
| VT_1-2-3 | Verleden tijd 1e, 2e en 3e persoon enkelvoud | Past tense first, second, and third person singular | &uparrow; | Mood=Ind \| Number=Sing \| Person=1 \| Tense=Past \| VerbForm=Fin and Mood=Ind \| Number=Sing \| Person=2 \| Tense=Past \| VerbForm=Fin and Mood=Ind \| Number=Sing \| Person=3 \| Tense=Past \| VerbForm=Fin |
| VT_4-5-6 | Verleden tijd 1e, 2e en 3e persoon meervoud | Past tense first, second, and third person plural | &uparrow; | Mood=Ind \| Number=Plur \| Person=1 \| Tense=Past \| VerbForm=Fin and Mood=Ind \| Number=Plur \| Person=2 \| Tense=Past \| VerbForm=Fin and Mood=Ind \| Number=Plur \| Person=3 \| Tense=Past \| VerbForm=Fin |
| gebiedende_wijs | N/A | Imperative form | &uparrow; | Mood=Imp \| VerbForm=Fin |
| aanvoegende_wijs | N/A | Conjunctive/subjunctive form | &uparrow; | Mood=Sub \| VerbForm=Fin |
| TW_deelwoord | Tegenwoordig deelwoord | Present participle | &uparrow; | Tense=Pres \| VerbForm=Part |
| VT_deelwoord | Voltooid deelwoord | Past participle | &uparrow; | Tense=Past \| VerbForm=Part |