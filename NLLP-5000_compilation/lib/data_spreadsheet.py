from lib import config
from lib.io_utils import dump_json, load_json
import csv
import logging
from openpyxl import Workbook
import os
import re
import string


def _apply_rules_xlsx(headword_orig: str, pos_full: str, pos_new: str, row: int) -> dict:
    """Apply set of predefined rules to automatically populate inflected word forms in spreadsheet to be annotated.

    :param headword_orig: Original headword from Routledge dictionary.
    :param pos_full: Original, full POS tag from Routledge dictionary.
    :param pos_new: New POS tag.
    :param row: Row number.
    :return: Dictionary containing cell IDs as keys and cell content as values.
    """
    d_cells_to_vals = {}

    if pos_new in ["adj", "interj"]:
        d_cells_to_vals[f"F{row}"] = headword_orig

    if pos_new == "adj":

        if headword_orig[-3:-1] in ["aa", "ee", "oo", "uu"]:
            d_cells_to_vals[f"G{row}"] = f"{headword_orig[:-2]}{headword_orig[-1]}e"
            d_cells_to_vals[f"H{row}"] = f"{headword_orig}s"
            d_cells_to_vals[f"I{row}"] = f"{headword_orig[:-2]}{headword_orig[-1]}er"
            d_cells_to_vals[f"J{row}"] = f"{headword_orig[:-2]}{headword_orig[-1]}ere"
            d_cells_to_vals[f"K{row}"] = f"{headword_orig[:-2]}{headword_orig[-1]}ers"
        else:
            d_cells_to_vals[f"G{row}"] = f"{headword_orig}e"
            d_cells_to_vals[f"H{row}"] = f"{headword_orig}s" if not headword_orig.endswith(("s", "sch")) else headword_orig
            d_cells_to_vals[f"I{row}"] = f"{headword_orig}er"
            d_cells_to_vals[f"J{row}"] = f"{headword_orig}ere"
            d_cells_to_vals[f"K{row}"] = f"{headword_orig}ers"

        if headword_orig.endswith("sch"):  # words ending in "sch" do not have a superlative form
            d_cells_to_vals[f"L{row}"] = "N/A"
            d_cells_to_vals[f"M{row}"] = "N/A"
        else:
            d_cells_to_vals[f"L{row}"] = f"{headword_orig}st" if not headword_orig.endswith("s") else f"{headword_orig}t"
            d_cells_to_vals[f"M{row}"] = f"{headword_orig}ste" if not headword_orig.endswith("s") else f"{headword_orig}te"

    elif pos_new == "adv":
        d_cells_to_vals[f"I{row}"] = headword_orig

    elif pos_new == "art":
        d_cells_to_vals[f"H{row}"] = headword_orig

    elif pos_new == "conj":
        d_cells_to_vals[f"H{row}"] = headword_orig

    elif pos_new == "noun":

        if pos_full == "noun, de(m)":
            d_cells_to_vals[f"F{row}"] = "x"
        elif pos_full == "noun, het":
            d_cells_to_vals[f"G{row}"] = "x"
        elif pos_full == "noun, de(f)":
            d_cells_to_vals[f"H{row}"] = "x"
        elif pos_full in ["noun, de(m)/het", "noun, het/de(m)", "noun, de(m)|het"]:
            d_cells_to_vals[f"I{row}"] = "x"
        elif pos_full == "noun, de":
            d_cells_to_vals[f"K{row}"] = "x"
        elif pos_full == "noun, de(f)/het":
            d_cells_to_vals[f"L{row}"] = "x"
        else:
            pass

        d_cells_to_vals[f"N{row}"] = headword_orig

        if headword_orig[-3:-1] in ["aa", "ee", "oo", "uu"]:
            d_cells_to_vals[f"O{row}"] = f"{headword_orig[:-2]}{headword_orig[-1]}en"
        else:
            d_cells_to_vals[f"O{row}"] = f"{headword_orig}en" if not headword_orig.endswith("e") else f"{headword_orig}n"

        if headword_orig.endswith("ing"):
            d_cells_to_vals[f"P{row}"] = f"{headword_orig[:-1]}kje"
            d_cells_to_vals[f"Q{row}"] = f"{headword_orig[:-1]}kjes"
        elif headword_orig.endswith(("d", "f", "g", "h", "k", "p", "s", "t")):
            d_cells_to_vals[f"P{row}"] = f"{headword_orig}je"
            d_cells_to_vals[f"Q{row}"] = f"{headword_orig}jes"
        else:
            d_cells_to_vals[f"P{row}"] = f"{headword_orig}tje"
            d_cells_to_vals[f"Q{row}"] = f"{headword_orig}tjes"

    elif pos_new == "num":
        d_cells_to_vals[f"J{row}"] = headword_orig

    elif pos_new == "prep":
        d_cells_to_vals[f"I{row}"] = headword_orig

    elif pos_new == "verb":
        d_cells_to_vals[f"J{row}"] = headword_orig
        d_cells_to_vals[f"N{row}"] = headword_orig

    else:
        pass

    return d_cells_to_vals


def _load_d_headers_spreadsheet() -> dict:
    """Get the headers to be used in the spreadsheet.

    :return: A dictionary containing the headers per POS.
    """
    with open(config.DIREC_SPREADSHEET_HEADERS / "common.txt", mode="r", encoding="utf-8") as f:
        l_common_headers = [header.strip() for header in f.readlines()]

    l_pos = [
        doc.replace(".txt", "").replace("specific_", "")
        for doc in os.listdir(config.DIREC_SPREADSHEET_HEADERS) if doc.startswith("specific_")
    ]
    d_headers = {}

    for pos in l_pos:

        with open(config.DIREC_SPREADSHEET_HEADERS / f"specific_{pos}.txt", mode="r", encoding="utf-8") as f:
            l_headers_pos = [header.strip() for header in f.readlines()]

        d_headers[pos] = l_common_headers + l_headers_pos

    return d_headers


def _load_d_mapping_pos_orig_to_new(d_headwords_orig: dict) -> dict:
    """Map part-of-speech (POS) tags as used in the Routledge resource to new POS tag (e.g., mapping masculine nouns,
    feminine nouns, and neuter nouns to the overarching tag 'noun').

    :param d_headwords_orig: Structured dictionary of the raw frequency dictionary.
    :return: The mapping dictionary.
    """
    d_mapping = {}

    for word in d_headwords_orig:

        for pos in d_headwords_orig[word]["POS_tags"]:

            if pos not in d_mapping:

                if pos in {"art", "conj", "prep", "verb", "pron", "adv", "num", "adj", "interj"}:
                    d_mapping[pos] = pos
                elif "noun" in pos:
                    d_mapping[pos] = "noun"
                else:
                    raise ValueError(f"Invalid POS tag: {pos} (headword = {word}).")

    return d_mapping


def convert_raw_input_into_dic(path_input_data: str, logger: logging.Logger) -> dict:
    """Convert raw input data into a structured Python dictionary.

    :param path_input_data: Path where the file containing the original Routledge dictionary in TXT format is located.
    :param logger: The logger.
    :return: The structured dictionary.
    """
    # LOAD RAW FREQUENCY LIST
    with open(path_input_data, mode="r", encoding="utf-8") as f:
        l_lines = [tuple(line.strip().split("\t")) for line in f.readlines()]

    # LOOP OVER LINES FROM RAW FILE AND STORE INFORMATION IN STRUCTURED DICTIONARY
    d_headwords_orig = {}
    l_headwords_orig = []  # list of headwords to be used for calculating number of duplicate entries (i.e. entries that occur in more than one list)
    d_sense_id_to_letter = {idx: letter for idx, letter in enumerate(string.ascii_lowercase, start=1)}

    for idx_line, tup_line in enumerate(l_lines):

        # Check if line is special line (i.e. line that does not contain an entry of the frequency list)
        if len(tup_line) != 11:
            logger.info(f"Line {idx_line + 1} does not contain 11 elements (only {len(tup_line)} found) and "
                        f"will be skipped: {tup_line}.")
            continue

        # Retrieve relevant information from line
        headword_orig = tup_line[1]  # original Routledge headword is situated at position 1 in tuple
        pos_raw = tup_line[2]  # Routledge POS is situated at position 2 (raw, unprocessed version)
        translation_en_raw = tup_line[3]  # English translation of the headword is situated at position 3
        tup_example_sents = tup_line[4:8]  # example sentences are situated at positions 4 to 7
        tup_mwes = tup_line[8:10]  # MWEs are situated at positions 8 to 9
        freq = float(tup_line[-1])  # frequency value is situated at the last position
        l_headwords_orig.append((headword_orig, pos_raw, freq))

        # Create empty dictionary in which data will be stored per POS
        d_data_per_pos = {}

        # Process entries that have more than one POS assigned to them (indicated in list by means of "1) ", "2) ",
        # etc.) for these entries, a separate entry per POS is created in the dictionary
        if "1) " in pos_raw:
            logger.info(f"\t\tEntry with more than one POS assigned to it: {tup_line}.")
            l_ids_pos = [f"POS_{item.strip(') ')}" for item in re.findall("[0-9]\\) ", pos_raw)]
            l_pos = [item.strip() for item in re.split("[0-9]\\) ", pos_raw) if item != ""]
            l_translations_en = [item.strip() for item in re.split("[0-9]\\) ", translation_en_raw) if item != ""]
            assert len(l_ids_pos) == len(l_pos) == len(l_translations_en)

            for id_pos, pos in zip(l_ids_pos, l_pos):

                if id_pos not in d_data_per_pos:
                    d_data_per_pos[id_pos] = {"POS": pos, "senses": {"sense_1": {}}}  # entries with multiple POSs never included multiple senses in the original Routledge dictionary (manually verified), so an empty sense entry can already be created

            for id_pos, translation_en in zip(l_ids_pos, l_translations_en):
                assert "a) " not in translation_en  # double check that entry does not have more than one sense
                d_data_per_pos[id_pos]["senses"]["sense_1"]["translation_EN"] = translation_en

        # Process entries that have one POS assigned to them
        else:

            #   - Retrieve different translations of the entry (indicated by means of "a) ", "b) ", etc.)
            l_ids_translations_en = [
                item.strip(") ") for item in re.findall("(^a\\) | [b-z]\\) )", translation_en_raw)
            ]
            l_translations_en = [
                item.strip() for item in re.split("^a\\) | [b-z]\\) ", translation_en_raw) if item != ""
            ]
            assert len(l_ids_translations_en) == 0 or len(l_ids_translations_en) >= 2, tup_line  # make sure no undesired matches are identified (either no matches are found or at least two)

            #   - Store translations to dictionary (as different senses)
            if len(l_ids_translations_en) == 0:  # only one translation = only one sense (length of `l_ids_translations_en` = 0 because no sense separators were found)
                d_data_per_pos["POS_1"] = {
                    "POS": pos_raw.strip(), "senses": {"sense_1": {"translation_EN": translation_en_raw.strip()}}
                }
            else:  # multiple translations = multiple senses (length of `l_ids_translations_en` = number of senses)
                d_data_per_pos["POS_1"] = {"POS": pos_raw.strip(), "senses": {}}

                for id_sense, translation_en in enumerate(l_translations_en):
                    entry_sense = f"sense_{id_sense + 1}"
                    assert entry_sense not in d_data_per_pos["POS_1"]["senses"]
                    d_data_per_pos["POS_1"]["senses"][entry_sense] = {"translation_EN": translation_en}

        if headword_orig not in d_headwords_orig:
            d_headwords_orig[headword_orig] = {"POS_tags": {}}

        for id_pos in d_data_per_pos:
            pos = d_data_per_pos[id_pos]["POS"]

            for id_sense in d_data_per_pos[id_pos]["senses"]:
                translation_en = d_data_per_pos[id_pos]["senses"][id_sense]["translation_EN"]

                if len(d_data_per_pos) > 1:
                    assert len(d_data_per_pos[id_pos]["senses"]) == 1  # double check that entries with multiple POSs have only one sense
                    l_example_sents = [
                        sent[3:] for sent in tup_example_sents if sent.startswith(f"{id_pos.split('_')[-1]}) ")
                    ]  # select example sentence corresponding to POS from tuple
                else:

                    if len(d_data_per_pos[id_pos]["senses"]) > 1:
                        l_example_sents = [
                            sent[3:] for sent in tup_example_sents
                            if sent.startswith(f"{d_sense_id_to_letter[int(id_sense.split('_')[-1])]}) ")
                        ]  # select example sentence corresponding to sense from tuple
                    else:
                        l_example_sents = [sent for sent in tup_example_sents if sent != ""]

                if pos not in d_headwords_orig[headword_orig]["POS_tags"]:
                    d_headwords_orig[headword_orig]["POS_tags"][pos] = {"senses": {}, "freq": freq}
                else:
                    assert d_headwords_orig[headword_orig]["POS_tags"][pos]["freq"] == freq, (headword_orig, pos, d_headwords_orig[headword_orig], freq)  # double check that words with multiple occurrences in word list have same frequency assigned to them

                if translation_en not in d_headwords_orig[headword_orig]["POS_tags"][pos]["senses"]:
                    d_headwords_orig[headword_orig]["POS_tags"][pos]["senses"][translation_en] = []

                for sent in l_example_sents:
                    d_headwords_orig[headword_orig]["POS_tags"][pos]["senses"][translation_en].append(sent)

    # DETERMINE NUMBER OF DUPLICATE ENTRIES
    n_duplicates = len(l_headwords_orig) - len(set(l_headwords_orig))
    logger.info(f"Total number of entries in original resource = {len(l_headwords_orig)} | "
                f"Number of unique entries = {len(set(l_headwords_orig))} | "
                f"Number of duplicate entries = {n_duplicates}")

    # DUMP DICTIONARY TO JSON FILE FOR FUTURE REFERENCE
    os.makedirs(config.DIREC_INP_DATA_PROCSD, exist_ok=True)
    dump_json(config.DIREC_INP_DATA_PROCSD, "d_headwords_orig.json", d_headwords_orig, indent=2)

    return d_headwords_orig


def convert_d_headwords_orig_into_unverified_spreadsheet(d_headwords_orig: dict) -> None:
    """Convert structured dictionary list into spreadsheet (CSV and XLSX) to be verified and completed by expert (some
    fields are automatically populated beforehand based on a set of predefined rules).

    :param d_headwords_orig: Structured dictionary.
    :return: `None`
    """
    # LOAD SUPPORT DATA
    d_headers = _load_d_headers_spreadsheet()
    d_mapping_pos = _load_d_mapping_pos_orig_to_new(d_headwords_orig)
    l_pos_new_sorted = sorted(list(set(list(d_mapping_pos.values()))))

    # CREATE EMPTY SPREADSHEET

    #   - XLSX
    os.makedirs(config.DIREC_INP_DATA_PROCSD, exist_ok=True)
    wb = Workbook()
    del wb["Sheet"]  # delete automatically created "Sheet" worksheet

    for pos_new in l_pos_new_sorted:
        ws = wb.create_sheet(pos_new)
        column = 1

        for header in d_headers[pos_new]:
            ws.cell(row=1, column=column, value=header)
            column += 1

    d_idx_row = {pos_new: 2 for pos_new in l_pos_new_sorted}

    for headword_orig in d_headwords_orig:

        for pos_word_full in d_headwords_orig[headword_orig]["POS_tags"]:
            pos_word_new = d_mapping_pos[pos_word_full]
            ws = wb[pos_word_new]
            row = d_idx_row[pos_word_new]
            ws[f"C{row}"] = d_headwords_orig[headword_orig]["POS_tags"][pos_word_full]["freq"]
            ws[f"D{row}"] = headword_orig
            ws[f"E{row}"] = pos_word_full

            if "," not in headword_orig:
                d_cells_to_vals = _apply_rules_xlsx(headword_orig, pos_word_full, pos_word_new, row)

                for cell in d_cells_to_vals:
                    ws[cell] = d_cells_to_vals[cell]

            d_idx_row[pos_word_new] += 1

    wb.save(os.path.join(config.DIREC_INP_DATA_PROCSD, "NLLP-5000_inflectedForms_unverified.xlsx"))

    #   - CSVs
    os.makedirs(config.DIREC_INP_DATA_PROCSD / "NLLP-5000_inflectedForms_unverified_CSVs", exist_ok=True)
    d_rows_per_pos = {pos_new: [] for pos_new in l_pos_new_sorted}
    l_col_letters_xlsx = [
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
        "W", "X", "Y", "Z"
    ]

    for headword_orig in d_headwords_orig:

        for pos_word_full in d_headwords_orig[headword_orig]["POS_tags"]:
            pos_word_new = d_mapping_pos[pos_word_full]
            d_data_xlsx_to_csv = {letter: "" for letter in l_col_letters_xlsx[:len(d_headers[pos_word_new])]}
            d_data_xlsx_to_csv["C"] = d_headwords_orig[headword_orig]["POS_tags"][pos_word_full]["freq"]
            d_data_xlsx_to_csv["D"] = headword_orig
            d_data_xlsx_to_csv["E"] = pos_word_full

            if "," not in headword_orig:
                d_cells_to_vals = _apply_rules_xlsx(headword_orig, pos_word_full, pos_word_new, 0)  # row is irrelevant, so put to 0

                for cell in d_cells_to_vals:
                    d_data_xlsx_to_csv[cell[0]] = d_cells_to_vals[cell]  # `cell[0]` corresponds to column letter XLSX

            if pos_word_new not in d_rows_per_pos:
                d_rows_per_pos[pos_word_new] = []

            d_rows_per_pos[pos_word_new].append(tuple([d_data_xlsx_to_csv[letter] for letter in l_col_letters_xlsx[:len(d_headers[pos_word_new])]]))

    for pos in l_pos_new_sorted:

        with open(
                config.DIREC_INP_DATA_PROCSD / "NLLP-5000_inflectedForms_unverified_CSVs" / f"NLLP-5000_inflectedForms_unverified_{pos}.csv",
                mode="w", newline="", encoding="utf-8"
        ) as f_csv:
            csv_writer = csv.writer(f_csv, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL)
            csv_writer.writerow(d_headers[pos])

            for tup in d_rows_per_pos[pos]:
                csv_writer.writerow(tup)
