#!/usr/bin/env python
import os

from setuptools import find_packages, setup


def read(fname):
    with open(os.path.join(os.path.dirname(__file__), fname)) as f:
        return f.read()


install_requires = ["comtypes", "psutil"]
setup(
    name="pycaw",
    version="20230407",
    description="Python Core Audio Windows Library",
    long_description=read("README.md"),
    long_description_content_type="text/markdown",
    author="Andre Miras",
    url="https://github.com/AndreMiras/pycaw",
    packages=find_packages(exclude=("tests", "examples")),
    install_requires=install_requires,
)
