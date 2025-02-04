# !/usr/bin/env python
# coding=utf-8
# GitHub: GWillS163
# User: 駿清清 
# Date: 09/11/2022 
# Time: 17:53

from setuptools import setup, find_packages
from pyCombiner import version, description
setup(
    name='pyCombiner',
    version=version,
    description=(
        description
    ),
    include_dirs=[
        'tests/examples',
        'tests/examples_refer_result',
    ],
    include_package_data=True,
    long_description=open('readme.md').read(),
    long_description_content_type='text/markdown',
    author='KissesJun',
    author_email='realgwills@gmail.com',
    maintainer='KissesJun',
    maintainer_email='realgwills@gmail.com',
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

    # scripts=['bin/main'],
    entry_points={
        # 'console_scripts': ['pyCombiner=pyCombiner.command_line:main'],
        'console_scripts': ['pyCombiner=pyCombiner.__main__:run'],
    }
)
