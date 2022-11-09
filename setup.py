# !/usr/bin/env python
# coding=utf-8
# GitHub: GWillS163
# User: 駿清清 
# Date: 09/11/2022 
# Time: 17:53

from setuptools import setup, find_packages

setup(
    name='pyCombiner',
    version="0.0.1",
    description=(
        'Combine that all your python files in your project sequential into one by the relationship of import statement. '
    ),
    long_description=open('README.rst').read(),
    author='JunQingQing',
    author_email='gwills@163.com',
    maintainer='JunQingQing',
    maintainer_email='gwills@163.com',
    license='BSD License',
    packages=find_packages(),
    platforms=["all"],
    url='https://gwills163.github.io/',
    classifiers=[
        'Development Status :: 4 - Beta',
        'Operating System :: OS Independent',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Programming Language :: Python :: Implementation',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Topic :: Software Development :: Libraries'
    ],
)