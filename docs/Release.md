# How to release

This is documenting the release process.


## Git flow & CHANGELOG.md

Make sure the CHANGELOG.md is up to date and follows the http://keepachangelog.com guidelines.
Start the release with git flow:
```batch
git flow release start vYYYYMMDD
```
Now update the [CHANGELOG.md](/CHANGELOG.md) `[Unreleased]` section to match the new release version.
Also update the `version` string in the [setup.py](/setup.py) file. Then commit and finish release.
```batch
git commit -a -m ":bookmark: vYYYYMMDD"
git flow release finish
```
Push everything, make sure tags are also pushed:
```batch
git push
git push origin main:main
git push --tags
```

## Publish to PyPI
This process is handled automatically by [GitHub Actions](https://github.com/AndreMiras/pycaw/actions/workflows/pypi-release.yml).
If needed below are the instructions to perform it manually.
Build it:
```batch
python setup.py sdist bdist_wheel
```
Check archive content:
```batch
tar -tvf dist\pycaw-*.tar.gz
```
Upload:
```batch
twine upload dist\pycaw-*.tar.gz
```

## GitHub

Got to GitHub [Release/Tags](https://github.com/AndreMiras/pycaw/tags), click "Add release notes" for the tag just created.
Add the tag name in the "Release title" field and the relevant CHANGELOG.md section in the "Describe this release" textarea field.

## Post release
Update the [setup.py](/setup.py) `version string` with `YYYYMMDD.dev0`.
