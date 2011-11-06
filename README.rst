logtogss
========

Description
-----------

This is a small tool to log data to a Google Spreadsheet. Google Spreadsheets are very handy for manipulating data, but it's less easy to insert data programmatically. The tool depends on the the `python gdata client`_ package, which should be automatically installed as a reqirement when installing the ``logtogss`` from setup package.

Installation
------------

To install, download the setup package_ and install using pip::

    $ pip install logtogss-0.1.tar.gz

Alternatively, let pip download from pypi and install::

    $ pip install logtogss

Usage
-----

The basic Usage is like this::

    $ logtogss --name "Spreadsheet Name" --sheet="Sheet Name" col1:val1 col2:val2

You can insert multiple rows at once by piping on stdin, whilst specifying just the column names, like this::

    $ cat data
    val1 val2
    val3 val4
    $ logtogss --name 'Spreadsheet Name' --sheet="Sheet Name" col1 col2 < data

If you need to know what the valid column names are, leave off the data from the command line, like this::

    $ logtogss --name "Spreadsheet Name" --sheet="Sheet Name"
    col1
    col2

You can also use the ``--list`` option wth or with out the ``--name`` and ``--sheet`` options to see a list of the columns in the spreadsheet(s).

Note that ``logtogss`` uses OAuth, so the first time that you use it, you will be prompted with an URL to visit in order to allow ``logtogss`` access to your google docs.

Run ``logtogss`` with ``--help`` option to see the advanced options and their usage.

Credits
-------

- `Happygiraffe`_ creator of the initial version.
- `modern-package-template`_ used to bootstrap the project setup.
- `Buildout`_

.. _Happygiraffe: https://github.com/happygiraffe/logss/
.. _`modern-package-template`: http://pypi.python.org/pypi/modern-package-template
.. _`python gdata client`: http://code.google.com/p/gdata-python-client/
.. _Buildout: http://www.buildout.org/
.. _package: https://github.com/downloads/haridsv/logss/logtogss-0.1.tar.gz
