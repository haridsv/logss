from setuptools import setup, find_packages
import sys, os

here = os.path.abspath(os.path.dirname(__file__))
README = open(os.path.join(here, 'README.rst')).read()
NEWS = open(os.path.join(here, 'NEWS.txt')).read()


version = '0.1'

install_requires = [
    'gdata>=2.0.0',
]


setup(name='logtogss',
    version=version,
    description="Log rows to a Google SpreadSheet",
    long_description=README + '\n\n' + NEWS,
    classifiers=[
      # Get strings from http://pypi.python.org/pypi?%3Aaction=list_classifiers
    ],
    keywords='gdata google spreadsheet csv import',
    author='Hari Dara',
    author_email='haridara@gmail.com',
    url='https://github.com/haridsv/logss',
    license='BSD 2-Clause',
    packages=find_packages('src'),
    package_dir = {'': 'src'},include_package_data=True,
    zip_safe=False,
    install_requires=install_requires,
    entry_points={
        'console_scripts':
            ['logtogss=logtogss:main']
    }
)
