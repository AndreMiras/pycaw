# How to release

This document describes the release process.


## Main branch & CHANGELOG.md

Releases use calendar versions in `YYYYMMDD` format. Start from an up-to-date
`main` branch:

```batch
git switch main
git pull --ff-only origin main
```

Make sure [CHANGELOG.md](/CHANGELOG.md) is up to date and follows the
https://keepachangelog.com guidelines. Rename its `[Unreleased]` section to
`[YYYYMMDD]`, then set `[project].version` in
[pyproject.toml](/pyproject.toml) to `YYYYMMDD`.

Run the test, lint, and package checks before committing:

```batch
tox
tox -e lint-check
python -m build
python -m twine check dist/*
```

Commit and push the release preparation to `main`:

```batch
git add CHANGELOG.md pyproject.toml
git commit -m ":bookmark: vYYYYMMDD"
git push origin main
```

Wait for the `main` branch workflows to pass. Tag that verified commit with an
annotated tag, then push only the new tag:

```batch
git tag -a vYYYYMMDD -m "vYYYYMMDD"
git push origin vYYYYMMDD
```

## Publish to PyPI
Pushing the version tag triggers publication automatically through
[GitHub Actions](https://github.com/AndreMiras/pycaw/actions/workflows/pypi-release.yml).
Confirm that its build, Twine check, and upload steps pass.
If needed below are the instructions to perform it manually.
Build it:
```batch
python -m build
python -m twine check dist/*
```
Check archive content:
```batch
tar -tvf dist\pycaw-*.tar.gz
```
Upload:
```batch
python -m twine upload dist/*
```

## GitHub

Go to GitHub [Release/Tags](https://github.com/AndreMiras/pycaw/tags) and click
"Add release notes" for the tag just created. Add the tag name in the "Release
title" field and the relevant CHANGELOG.md section in the "Describe this
release" field.

## Post release
Add a new `[Unreleased]` section to [CHANGELOG.md](/CHANGELOG.md), update the
`[project].version` in [pyproject.toml](/pyproject.toml) to `YYYYMMDD.dev0`,
then commit and push the next development version:

```batch
git add CHANGELOG.md pyproject.toml
git commit -m ":construction: Post release dev0"
git push origin main
```
