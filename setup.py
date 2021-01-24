from pathlib import Path
from setuptools import setup
from xlrd.info import __VERSION__

read_me = Path(__file__).parent / 'README.md'
long_description = read_me.read_text(encoding='utf-8')

setup(
    name='xlrd3',
    version=__VERSION__,
    author='John Machin',
    author_email='sjmachin@lexicon.net',
    maintainer='Nguyen Ba Duc Tin',
    maintainer_email='nguyenbaduc.tin@gmail.com',
    url='https://github.com/Dragon2fly/xlrd3',
    packages=['xlrd3'],
    package_dir={'xlrd3': 'xlrd'},
    scripts=[
        'scripts/runxlrd.py',
    ],
    description=(
        'Library for developers to extract data from '
        'Microsoft Excel (tm) spreadsheet files'
    ),
    long_description=long_description,
    long_description_content_type='text/markdown',
    platforms=["Any platform -- don't need Windows"],
    license='BSD',
    keywords=['xls', 'xlsx', 'excel', 'spreadsheet', 'workbook'],
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Operating System :: OS Independent',
        'Topic :: Database',
        'Topic :: Office/Business',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
    python_requires=">=3.6",
)
