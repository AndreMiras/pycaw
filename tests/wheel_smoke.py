import importlib.metadata
import pathlib
import sys

import tomllib

import pycaw


def main():
    source_root = pathlib.Path(sys.argv[1]).resolve()
    source_package = source_root / "pycaw"
    installed_package = pathlib.Path(pycaw.__file__).resolve().parent
    metadata = importlib.metadata.metadata("pycaw")
    with (source_root / "pyproject.toml").open("rb") as pyproject_file:
        expected_version = tomllib.load(pyproject_file)["project"]["version"]

    assert installed_package != source_package
    assert not installed_package.is_relative_to(source_package)

    source_modules = {
        path.relative_to(source_package) for path in source_package.rglob("*.py")
    }
    installed_modules = {
        path.relative_to(installed_package) for path in installed_package.rglob("*.py")
    }
    assert installed_modules == source_modules

    assert importlib.metadata.version("pycaw") == expected_version
    assert metadata["Requires-Python"] == ">=3.10"
    assert set(metadata.get_all("Requires-Dist")) == {
        "comtypes>=1.4.8",
        "psutil>=5.9.0",
    }


if __name__ == "__main__":
    main()
