# !/usr/bin/env python
# coding=utf-8
# GitHub: GWillS163
# User: 駿清清 
# Date: 09/11/2022 
# Time: 17:53

from setuptools import setup, find_packages
import tomli

def _get_version():
    with open("pyproject.toml", "rb") as f:
        return tomli.load(f)["version"]

setup(
    name='pycombiner',
    version=_get_version(),
    description=(
        "Combine that all your python files in your project sequential into one by the relationship of import satement."
    ),
    include_dirs=[
        'pycombiner/'
    ],
    package_data={
        'pycombiner': ['tests/examples/*', 'tests/examples_refer_result/*'],
    },
    packages=find_packages(),
    include_package_data=True,

    long_description=open('readme.md', encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    author='KissesJun',
    author_email='meng.junqing1022@gmail.com',
    maintainer='KissesJun',
    maintainer_email='meng.junqing1022@gmail.com',
    license='BSD License',
    platforms=["all"],
    url='https://kissesJun.github.io/',
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
        'console_scripts': ['pycombiner=pycombiner.__main__:run'],
    }
)
