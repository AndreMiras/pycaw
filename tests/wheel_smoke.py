import importlib.metadata
import pathlib
import sys

import pycaw


def main():
    source_package = pathlib.Path(sys.argv[1], "pycaw").resolve()
    installed_package = pathlib.Path(pycaw.__file__).resolve().parent
    metadata = importlib.metadata.metadata("pycaw")

    assert installed_package != source_package
    assert not installed_package.is_relative_to(source_package)

    source_modules = {
        path.relative_to(source_package) for path in source_package.rglob("*.py")
    }
    installed_modules = {
        path.relative_to(installed_package) for path in installed_package.rglob("*.py")
    }
    assert installed_modules == source_modules

    assert importlib.metadata.version("pycaw") == "20260921.dev0"
    assert metadata["Requires-Python"] == ">=3.10"
    assert set(metadata.get_all("Requires-Dist")) == {
        "comtypes>=1.4.8",
        "psutil>=5.9.0",
    }


if __name__ == "__main__":
    main()
