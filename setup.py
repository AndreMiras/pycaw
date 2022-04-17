#!/usr/bin/env python
import os

from setuptools import setup, find_packages


def read(fname):
    with open(os.path.join(os.path.dirname(__file__), fname)) as f:
        return f.read()


install_requires = [
    'comtypes', 'psutil']
setup(name='pycaw',
      version='20220416.dev0',
      description='Python Core Audio Windows Library',
      long_description=read('README.md'),
      long_description_content_type='text/markdown',
      author='Andre Miras',
      url='https://github.com/AndreMiras/pycaw',
      packages=find_packages(exclude=("tests", "examples")),
      install_requires=install_requires)
