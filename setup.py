#!/usr/bin/env python
import os

from setuptools import find_packages, setup


def read(fname):
    with open(os.path.join(os.path.dirname(__file__), fname), encoding="utf-8") as f:
        return f.read()


install_requires = ["comtypes>=1.4.8", "psutil>=5.9.0"]
setup(
    name="pycaw",
    version="20260921.dev0",
    description="Python Core Audio Windows Library",
    long_description=read("README.md"),
    long_description_content_type="text/markdown",
    author="Andre Miras",
    url="https://github.com/AndreMiras/pycaw",
    packages=find_packages(exclude=("tests", "examples")),
    python_requires=">=3.10",
    platforms=["Windows"],
    classifiers=[
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3 :: Only",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Programming Language :: Python :: 3.14",
    ],
    install_requires=install_requires,
)
