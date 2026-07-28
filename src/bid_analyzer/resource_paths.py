from pathlib import Path


PACKAGE_DIRECTORY = Path(__file__).resolve().parent
RESOURCES_DIRECTORY = PACKAGE_DIRECTORY / "resources"


def resource_path(filename: str) -> Path:
    """
    Return the full path to a packaged application resource.

    Examples:
        resource_path("airports.csv")
        resource_path("app_icon.ico")
    """
    path = RESOURCES_DIRECTORY / filename

    if not path.exists():
        raise FileNotFoundError(
            f"Application resource was not found: {path}"
        )

    return path