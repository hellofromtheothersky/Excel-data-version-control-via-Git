from setuptools import setup, find_packages

with open('requirements.txt') as f:
    REQUIRED_LIBS = f.read().splitlines()

VERSION = "1.0"

setup(
    name="gitexcel",
    description="Python CLI tool with Git for MS Excel version control using text",
    # long_description_content_type="text/markdown",
    author="Hieu Nguyen Hung Trung",
    version=VERSION,
    license="Apache License, Version 2.0",
    packages=find_packages(),
    install_requires=REQUIRED_LIBS,
    setup_requires=["pytest-runner"],
    entry_points="""
        [console_scripts]
        gitexcel=gitexcel:cli
    """,
    url="https://github.com/hellofromtheothersky/Excel-data-version-control-via-Git",
)