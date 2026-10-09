from pathlib import Path


def main() -> None:
    project_folder = Path(__file__).parent

    print("Invoice reconciliation project is running.")
    print(f"Project folder: {project_folder}")


if __name__ == "__main__":
    main()