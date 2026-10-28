import json
from pathlib import Path
import sys
from typing import Any


def dump_json(
        path_direc: str | Path, fn: str, variable: Any,
        *,
        indent: int | None = None
) -> None:
    """Dump a variable to a JSON file, creating the directory if needed.

    :param path_direc: Directory in which the JSON file is written (created if it does not exist).
    :param fn: Name of the JSON file.
    :param variable: JSON-serialisable variable to be written.
    :param indent: Indentation of the JSON file. Defaults to `None` (compact output).
    :return: `None`
    """
    path_direc = Path(path_direc)
    path_direc.mkdir(parents=True, exist_ok=True)
    path_file = path_direc / fn
    path_tmp = path_file.with_suffix(f"{path_file.suffix}.tmp")

    with path_tmp.open(mode="w", encoding="utf-8") as f:
        json.dump(variable, f, indent=indent, ensure_ascii=False)

    path_tmp.replace(path_file)  # atomic rename on the same file system


def load_json(
        path: str | Path,
        *,
        encoding: str = "utf-8"
) -> Any:
    """Load a JSON file.

    :param path: Path of the JSON file.
    :param encoding: Encoding of the JSON file. Defaults to `"utf-8"`.
    :return: The parsed JSON content.
    """
    with Path(path).open(encoding=encoding) as f:
        return json.load(f)
