#!/usr/bin/env python

from distutils.core import setup

setup(name='pycaw',
      version='20190725',
      description='Python Core Audio Windows Library',
      author='Andre Miras',
      url='https://github.com/AndreMiras/pycaw',
      packages=['pycaw'],
      install_requires=['comtypes', 'enum-compat', 'psutil', 'future'])
