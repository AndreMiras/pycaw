#!/usr/bin/env python
import os

from setuptools import setup


def read(fname):
    with open(os.path.join(os.path.dirname(__file__), fname)) as f:
        return f.read()


install_requires = [
    'comtypes', 'enum34;python_version<"3.4"', 'psutil', 'future']
setup(name='pycaw',
      version='20190904',
      description='Python Core Audio Windows Library',
      long_description=read('README.md'),
      long_description_content_type='text/markdown',
      author='Andre Miras',
      url='https://github.com/AndreMiras/pycaw',
      packages=['pycaw'],
      install_requires=install_requires)
