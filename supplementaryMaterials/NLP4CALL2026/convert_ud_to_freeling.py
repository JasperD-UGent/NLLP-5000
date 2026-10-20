import argparse


def convert_ud_to_freeling_verbs(entry: str) -> str:
    """Convert entry in UD format to FreeLing tag (for verbs).

    :param entry: Raw entry containing UD information.
    :return: The FreeLing tag.
    """
    # CREATE PROVISIONAL DICTIONARY CONTAINING FREELING FEATURE-VALUE PAIRS
    d_vals_freeling = {
        # lexical ("full") verb by default; will be changed to "A" for regular auxiliary verbs or "M" for modal verbs if applicable
        "verb_type": "V",

        # 0 (i.e. value if mood is not applicable) by default; will be changed to "C" for conjunctive/subjunctive, "I" for indicative, "M" for imperative, "N" for infinitive, or "P" for participle if applicable
        "mood": 0,

        # 0 (i.e. value if tense is not applicable) by default; will be changed to "P" for present or "S" for past if applicable
        "tense": 0,

        # 0 (i.e. value if number is not applicable) by default; will be changed to "P" for plural or "S" for singular if applicable
        "number": 0,

        # 0 (i.e. value if person is not applicable) by default; will be changed to 1 for first person, 2 for second person, or 3 for third person if applicable
        "person": 0
    }

    # EXTRACT UD INFORMATION FROM ENTRY
    l_parts_entry = entry.split("\t")
    pos_ud = l_parts_entry[-2]
    feats_ud = l_parts_entry[-1]

    # APPLY RULES TO CHANGE VALUES IN PROVISIONAL DICTIONARY

    #   - Verb type
    if pos_ud == "AUX":

        if "VerbType=Mod" in feats_ud:
            d_vals_freeling["verb_type"] = "M"
        else:
            d_vals_freeling["verb_type"] = "A"

    #   - Mood
    if "Mood=Sub" in feats_ud:
        d_vals_freeling["mood"] = "C"

    if "Mood=Ind" in feats_ud:
        d_vals_freeling["mood"] = "I"

    if "Mood=Imp" in feats_ud:
        d_vals_freeling["mood"] = "M"

    if "VerbForm=Inf" in feats_ud:
        d_vals_freeling["mood"] = "N"

    if "VerbForm=Part" in feats_ud:
        d_vals_freeling["mood"] = "P"

    #   - Tense
    if "Tense=Pres" in feats_ud:
        d_vals_freeling["tense"] = "P"

    if "Tense=Past" in feats_ud:
        d_vals_freeling["tense"] = "S"

    #   - Number
    if "Number=Sing" in feats_ud:
        d_vals_freeling["number"] = "S"

    if "Number=Plur" in feats_ud:
        d_vals_freeling["number"] = "P"

    #   - Person
    if "Person=1" in feats_ud:
        d_vals_freeling["person"] = "1"

    if "Person=2" in feats_ud:
        d_vals_freeling["person"] = "2"

    if "Person=3" in feats_ud:
        d_vals_freeling["person"] = "3"

    freeling_tag = (f"V"
                    f"{d_vals_freeling['verb_type']}"
                    f"{d_vals_freeling['mood']}"
                    f"{d_vals_freeling['tense']}"
                    f"{d_vals_freeling['number']}"
                    f"{d_vals_freeling['person']}")

    return freeling_tag


def main(path_input_data: str) -> None:
    """Entry point for the script. Selects the appropriate function to process the input file and prints the FreeLing
    tags to the screen.

    :param path_input_data: Path to file containing the UD data to be converted into FreeLing tags.
    :return: `None`
    """

    with open(path_input_data, mode="r", encoding="utf-8") as f:
        l_entries = [line.strip() for line in f.readlines()]

    if path_input_data.endswith("_verbs.txt"):
        l_freeling_tags = [(entry, convert_ud_to_freeling_verbs(entry)) for entry in l_entries]
    else:
        raise ValueError(f"Invalid input file: {path_input_data}. File has to end with: '_verbs.txt'.")

    for idx, tup in enumerate(l_freeling_tags):
        print(f"=========== Entry {idx + 1} ===========")
        print(f"- Original entry: {tup[0]}")
        print(f"- FreeLing tag: {tup[1]}")
        print("\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument(
        "path_input_data", type=str,
        help="Path to file containing the UD data to be converted into FreeLing tags."
    )
    main(**vars(parser.parse_args()))
