# How to release

This document describes the release process.


## Main branch & CHANGELOG.md

Releases use calendar versions in `YYYYMMDD` format. Start from an up-to-date
`main` branch:

```batch
set VERSION=YYYYMMDD
git switch main
git pull --ff-only origin main
```

Make sure [CHANGELOG.md](/CHANGELOG.md) is up to date and follows the
https://keepachangelog.com guidelines. Replace `YYYYMMDD` in `VERSION` with the
release date, rename the changelog's `[Unreleased]` section to that value, then
set `[project].version` in [pyproject.toml](/pyproject.toml) to the same value.

Install the locked contributor environment, then run the test and lint checks:

```batch
uv lock
uv sync --locked
uv run tox
uv run tox -e lint-check
```

Build and inspect the distributions with the locked release tools:

```batch
uv sync --locked --only-group release
uv run --no-sync python -m build
uv run --no-sync python -m twine check dist/*
tar -tvf dist\pycaw-*.tar.gz
```

Commit and push the release preparation to `main`:

```batch
git add CHANGELOG.md pyproject.toml uv.lock
git commit -m ":bookmark: v%VERSION%"
git push origin main
```

Wait for the `main` branch workflows to pass. Tag that verified commit with an
annotated tag, then push only the new tag:

```batch
git tag -a v%VERSION% -m "v%VERSION%"
git push origin v%VERSION%
```

## Publish to PyPI

Pushing an exact `vYYYYMMDD` tag triggers the
[PyPI release workflow](https://github.com/AndreMiras/pycaw/actions/workflows/pypi-release.yml).
The workflow rejects malformed tags and versions that do not match
`[project].version`. Its unprivileged Windows job builds and checks one wheel
and one source archive, then uploads them as the
`python-package-distributions` Actions artifact. A separate Linux job downloads
those unchanged files and publishes through PyPI Trusted Publishing. Only that
job receives an OIDC identity, and it does not check out or execute repository
code.

Approve the protected `pypi` GitHub environment if required. After publication:

1. Download the `python-package-distributions` artifact from the workflow run.
2. Compare its SHA-256 hashes with the files and hashes shown on PyPI.
3. Confirm the PyPI version, Python requirement, dependencies, classifiers, and
   project links.
4. Confirm PyPI displays attestations for both files.

Do not work around a Trusted Publishing failure with a long-lived token. Fix
the publisher identity, environment, workflow permission, tag, or version as
indicated by the failed job. If PyPI accepted any file, do not overwrite or
reuse that immutable version; prepare a new calendar version instead.

## GitHub

Go to GitHub [Release/Tags](https://github.com/AndreMiras/pycaw/tags) and click
"Add release notes" for the tag just created. Add the tag name in the "Release
title" field and the relevant CHANGELOG.md section in the "Describe this
release" field.

## Post release
Add a new `[Unreleased]` section to [CHANGELOG.md](/CHANGELOG.md), update the
`[project].version` in [pyproject.toml](/pyproject.toml) to `%VERSION%.dev0`,
then commit and push the next development version:

```batch
uv lock
git add CHANGELOG.md pyproject.toml uv.lock
git commit -m ":construction: Post release dev0"
git push origin main
```
