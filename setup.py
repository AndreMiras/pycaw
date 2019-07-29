#!/usr/bin/env python

from distutils.core import setup

install_requires = [
    'comtypes', 'enum34;python_version<"3.4"', 'psutil', 'future']
setup(name='pycaw',
      version='20181226',
      description='Python Core Audio Windows Library',
      author='Andre Miras',
      url='https://github.com/AndreMiras/pycaw',
      packages=['pycaw'],
      install_requires=install_requires)
