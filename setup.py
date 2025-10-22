# coding=utf-8
# pzw
# 20251022

from setuptools import setup, find_packages
import os

# 获取 setup.py 所在目录的绝对路径
here = os.path.abspath(os.path.dirname(__file__))

# 读取 README.md，如果不存在则使用默认描述
try:
    with open(os.path.join(here, 'README.md'), 'r', encoding='utf-8') as readme_file:
        long_description = readme_file.read()
except FileNotFoundError:
    long_description = 'A Python library for Word document template processing with OOP API'

setup(
    name='WordWriter',
    version='4.0.3',
    description='A Python library for Word document template processing with OOP API',
    long_description=long_description,
    long_description_content_type='text/markdown',
    author='pzweuj',
    author_email='pzweuj@live.com',
    url='https://github.com/pzweuj/WordWriter',
    project_urls={
        'Bug Reports': 'https://github.com/pzweuj/WordWriter/issues',
        'Source': 'https://github.com/pzweuj/WordWriter',
    },
    install_requires=[
        "python-docx>=0.8.10",
        "pandas>=1.0.0"
    ],
    python_requires='>=3.6',
    license='MIT',
    packages=find_packages(),
    include_package_data=True,
    platforms=["all"],
    keywords=['docx', 'word', 'template', 'document', 'office'],
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
        'Natural Language :: Chinese (Simplified)',
        'Natural Language :: English',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
        'Topic :: Software Development :: Libraries',
        'Topic :: Office/Business',
        'Topic :: Text Processing',
    ],
)
