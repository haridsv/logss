#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Log a row to a Google Spreadsheet."""

__author__ = 'Dominic Mitchell <dom@happygiraffe.net>'


import getpass
import optparse
import sys

import gdata.spreadsheet.service


class Error(Exception):
  pass


def Authenticate(client, username):
  # TODO: OAuth.  We must be able to do this without a password.
  client.ClientLogin(username,
                     getpass.getpass('Password for %s: ' % username))


def ExtractKey(entry):
  # This is what spreadsheetExample seems to do…
  return entry.id.text.split('/')[-1]


def FindKeyOfSheet(client, name):
  spreadsheets = client.GetSpreadsheetsFeed()
  spreadsheet = [s for s in spreadsheets.entry if s.title.text == name]
  if not spreadsheet:
    raise Error('Can\'t find spreadsheet named %s', name)
  if len(spreadsheet) > 1:
    raise Error('More than one spreadsheet named %s', name)
  return ExtractKey(spreadsheet[0])


def DefineFlags():
  usage = u"""usage: %prog [options] spreadsheet_name [col1:va1 …]"""
  desc = """
Log data into a Google Spreadsheet.

With no further arguments, a list of column names will be printed to stdout.

Otherwise, remaining arguments should be of the form `columnname:value'.
One row will be added for each invocation of this program.
  """
  parser = optparse.OptionParser(usage=usage, description=desc)
  parser.add_option('--debug', dest='debug', action='store_true',
                    help='Enable debug output', default=False)
  parser.add_option('-u', '--username', dest='username',
                    help='Which username to log in as (default: %default)',
                    default='%s@gmail.com' % getpass.getuser())
  return parser


def main():
  parser = DefineFlags()
  (opts, args) = parser.parse_args()
  if not args:
    parser.error('You must specify a spreadsheet name.')
  spreadsheet_name = args[0]

  client = gdata.spreadsheet.service.SpreadsheetsService()
  client.debug = opts.debug
  Authenticate(client, opts.username)

  key = FindKeyOfSheet(client, spreadsheet_name)
  if len(args) > 1:
    args = dict(x.split(':', 1) for x in argv[1:])
    client.InsertRow(args, key)
  else:
    list_feed = client.GetListFeed(key)
    for col in sorted(list_feed.entry[0].custom.keys()):
      print col
  return 0


if __name__ == '__main__':
  sys.exit(main())
