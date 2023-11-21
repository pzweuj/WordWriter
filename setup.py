# coding=utf-8
# pzw
# 20231009

from setuptools import setup, find_packages

with open('README.md', 'r', encoding='utf-8') as readme_file:
    long_description = readme_file.read()

setup(name='WordWriter',
    version = '3.1.1',
    description = 'Docx file template replacing',
    long_description = long_description,
    long_description_content_type =' text/markdown',
    author = 'pzweuj',
    author_email = 'pzweuj@live.com',
    url = 'https://github.com/pzweuj/WordWriter',
    install_requires = ["python-docx", "pandas"],
    license = 'MIT License',
    packages = find_packages(),
    platforms = ["all"],
    classifiers = [
        'Intended Audience :: Developers',
        'Operating System :: OS Independent',
        'Natural Language :: Chinese (Simplified)',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Topic :: Software Development :: Libraries'
    ],
)
