from pathlib import Path

ALLOWED_EXTENSIONS = {'.csv', '.xls', '.xlsx'}

def is_valid_input_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in ALLOWED_EXTENSIONS


def collect_input_files(paths: list[str]) -> list[Path]:
    collected: list[Path] = []

    for p_str in paths:
        p = Path(p_str)

        if not p.exists():
            print(f"⚠️  Skipping '{p}': path does not exist.")
            continue

        if p.is_dir():
            for file in p.iterdir():
                if is_valid_input_file(file):
                    collected.append(file)
                elif file.is_file():
                    print(
                       f"⚠️  Skipping '{file.name}': "
                       f"'.{file.suffix.lstrip('.')}' is not an accepted file type. "
                       f"Accepted types are: {', '.join(ALLOWED_EXTENSIONS)}"
                    )
        else:
            if is_valid_input_file(p):
                collected.append(p)
            else:
                print(
                    f"⚠️  Skipping '{p.name}': "
                    f"'.{p.suffix.lstrip('.')}' is not an accepted file type. "
                    f"Accepted types are: {', '.join(ALLOWED_EXTENSIONS)}"
                )

    return collected