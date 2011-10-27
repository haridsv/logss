#!/usr/bin/env python
# -*- coding: utf-8 -*-
# References:
# - Column tags: https://groups.google.com/d/msg/google-spreadsheets-api/dZvMNrfcfQU/OyXBWEZyKdwJ
# - GSX: http://code.google.com/apis/spreadsheets/data/3.0/reference.html#gsx_reference
# - List Feed: http://code.google.com/apis/spreadsheets/data/3.0/developers_guide.html#ListFeeds

"""Log a row to a Google Spreadsheet."""


__author__ = 'Dominic Mitchell <dom@happygiraffe.net>'


import pickle
import optparse
import os
import sys
import urllib

import gdata.gauth
import gdata.spreadsheets.client
import gdata.spreadsheets.data

import oneshot


# OAuth bits.  We use “anonymous” to behave as an unregistered application.
# http://code.google.com/apis/accounts/docs/OAuth_ref.html#SigningOAuth
CONSUMER_KEY = 'anonymous'
CONSUMER_SECRET = 'anonymous'
# The bits we actually need to access.
SCOPES = ['https://spreadsheets.google.com/feeds/']


class TokenStore(object):
  """Store and retreive OAuth access tokens."""

  def __init__(self, token_file=None):
    default = os.path.expanduser('~/.%s.tok' % os.path.basename(sys.argv[0]))
    self.token_file = token_file or default

  def ReadToken(self):
    """Read in the stored auth token object.

    Returns:
      The stored token object, or None.
    """
    try:
      with open(self.token_file, 'rb') as fh:
        return pickle.load(fh)
    except IOError, e:
      return None

  def WriteToken(self, tok):
    """Write the token object to a file."""
    with open(self.token_file, 'wb') as fh:
      os.fchmod(fh.fileno(), 0600)
      pickle.dump(tok, fh)


class ClientAuthorizer(object):
  """Add authorization to a client."""

  def __init__(self, consumer_key=CONSUMER_KEY,
               consumer_secret=CONSUMER_SECRET, scopes=None,
               token_store=None, logger=None):
    """Construct a new ClientAuthorizer."""
    self.consumer_key = consumer_key
    self.consumer_secret = consumer_secret
    self.scopes = scopes or list(SCOPES)
    self.token_store = token_store or TokenStore()
    self.logger = self.LogToStdout

  def LogToStdout(self, msg):
    print msg

  def FetchAccessToken(self, client):
    # http://code.google.com/apis/gdata/docs/auth/oauth.html#Examples
    httpd = oneshot.ParamsReceiverServer()

    # TODO Find a way to pass "xoauth_displayname" parameter.
    request_token = client.GetOAuthToken(
        self.scopes, httpd.my_url(), self.consumer_key, self.consumer_secret)
    url = request_token.generate_authorization_url()
    self.logger('Please visit this URL to authorize: %s' % url)
    httpd.serve_until_result()
    gdata.gauth.AuthorizeRequestToken(request_token, httpd.result)
    return client.GetAccessToken(request_token)

  def EnsureAuthToken(self, client):
    """Ensure client.auth_token is valid.

    If a stored token is available, it will be used.  Otherwise, this goes
    through the OAuth rituals described at:

    http://code.google.com/apis/gdata/docs/auth/oauth.html

    As a side effect, this also reads and stores the token in a file.
    """
    access_token = self.token_store.ReadToken()
    if not access_token:
      access_token = self.FetchAccessToken(client)
      self.token_store.WriteToken(access_token)
    client.auth_token = access_token


# The next three classes are overrides to add missing functionality in the
# python-gdata-client.

class MyListEntry(gdata.spreadsheets.data.ListEntry):

  def ColumnTags(self):
    """Return the names of all child elements in the GSX namespace."""
    ns = gdata.spreadsheets.data.GSX_NAMESPACE
    return [el.tag for el in self.get_elements(namespace=ns)]

  def ColumnValueToTagMap(self):
    """
    Return the values of all child elements keyed by their tags in the GSX namespace.
    Useful on a header row to map the column names to their tags.
    """
    ns = gdata.spreadsheets.data.GSX_NAMESPACE
    return dict([(el.text, el.tag) for el in self.get_elements(namespace=ns)])


class MyListsFeed(gdata.spreadsheets.data.ListsFeed):

  entry = [MyListEntry]

  def ColumnTags(self):
    if not self.entry:
      return []
    return self.entry[0].ColumnTags()

  def ColumnValueToTagMap(self):
    if not self.entry:
      return []
    return self.entry[0].ColumnValueToTagMap()


class MySpreadsheetsClient(gdata.spreadsheets.client.SpreadsheetsClient):
  """Add in support for List feeds."""

  LISTS_URL = 'https://spreadsheets.google.com/feeds/list/%s/%s/private/full'
  CELLS_URL = 'https://spreadsheets.google.com/feeds/cells/%s/%s/private/full'

  def get_list_feed(self, key, wksht_id='default', start_index=None, max_results=None, **kwargs):
    return self._get_feed(self.LISTS_URL, key, wksht_id,
                          desired_class=MyListsFeed,
                          start_index=start_index,
                          max_results=max_results,
                          kwargs=kwargs)

  GetListFeed = get_list_feed

  def get_cells_feed(self, key, wksht_id='default', start_index=None,
                     min_col=None, max_col=None, min_row=None, max_row=None,
                     max_results=None, **kwargs):
    return self._get_feed(self.CELLS_URL, key, wksht_id,
                          start_index=start_index,
                          min_col=min_col,
                          max_col=max_col,
                          min_row=min_row,
                          max_row=max_row,
                          max_results=max_results,
                          kwargs=kwargs)

  GetCellsFeed = get_cells_feed

  def _get_feed(self, baseuri, key, wksht_id='default', desired_class=None, **params):
    kwargs = params.pop('kwargs')

    uri = baseuri % (key, wksht_id)
    # Remove the params with None values, so they don't get into the query
    params = dict([(key.replace('_', '-'), value) for (key, value) in params.items() if value is not None])
    if params:
      uri += ('?' + urllib.urlencode(params.items()))
    if desired_class:
      kwargs['desired_class'] = desired_class
    return self.get_feed(uri, **kwargs)


class LogssAction(object):

  def __init__(self, debug=False):
    self.debug = debug
    self.client = MySpreadsheetsClient()
    self.client.debug = debug
    self.client.http_client.debug = debug
    self.client.source = os.path.basename(sys.argv[0])

  def Authenticate(self, logger=None):
    client_authz = ClientAuthorizer(logger=logger)
    client_authz.EnsureAuthToken(self.client)

  def GetSpreadsheets(self, ss=None, ss_is_id=False):
    """
    Return a generator of spreadsheet (name, id) pairs.
    Given a spreadsheet name or ID, entry for only that spreadsheet is
    generated.
    """
    # Get all spreadsheets.
    spreadsheets = self.client.GetSpreadsheets()
    for (ssname, ssid) in self._gen_name_id(spreadsheets.entry):
      if ss:
        if (ss_is_id and ss == ssid) or (ssname == ss):
          yield ssname, ssid
      else:
        yield ssname, ssid

  def GetWorksheets(self, ssid, ws=None, ws_is_id=False):
    """
    Return a generator of worksheet (nanme, id) pairs for the specified
    spreadsheet..
    Given a worksheet name or id, only the entry for that worksheet is
    generated.
    """
    worksheets = self.client.GetWorksheets(ssid)
    for (wsname, wsid) in self._gen_name_id(worksheets.entry):
      if ws:
        if (ws_is_id and ws == wsid) or (wsname == ws):
          yield wsname, wsid
      else:
        yield wsname, wsid

  def _gen_name_id(self, entries):
    for entry in entries:
      yield entry.title.text, entry.id.text.split('/')[-1]

class SpreadsheetInserter(LogssAction):
  """A utility to insert rows into a spreadsheet."""

  def __init__(self, debug=False):
    super(SpreadsheetInserter, self).__init__(debug)
    self.key = None
    self.wkey = None
    self.col_name_to_key = None

  def ColumnNamesHaveData(self, cols):
    """Are these just names, or do they have data (:)?"""
    return len([c for c in cols if ':' in c]) > 0

  def InsertRow(self, data):
    row_entry = gdata.spreadsheets.data.ListEntry()
    if self.col_name_to_key:
      data = dict([(self.col_name_to_key[name], value) for (name, value) in data.items()])
    row_entry.from_dict(data)
    self.client.add_list_entry(row_entry, self.key, self.wkey)

  def InsertFromColumns(self, cols):
    # Data is mixed into column names.
    data = dict(c.split(':', 1) for c in cols)
    self.InsertRow(data)

  def InsertFromFileHandle(self, cols, fh):
    for line in fh:
      vals = line.rstrip().split(None, len(cols) - 1)
      data = dict(zip(cols, vals))
      self.InsertRow(data)

  def ListColumns(self):
    """
    Return tuples containing the column name and the tags.
    """
    if self.col_name_to_key:
      cols = self.col_name_to_key.items()
    else:
      list_feed = self.client.GetListFeed(self.key, wksht_id=self.wkey, max_results=1)
      coltags = list_feed.ColumnTags()
      cols = zip(coltags, coltags)
    return sorted(cols)

  def SetColumnHeaderRowNum(self, headerRowNum):
    if headerRowNum == 1:
      list_feed = self.client.GetListFeed(self.key, wksht_id=self.wkey)
      coltags = list_feed.ColumnTags()
      colcells = self.client.GetCellsFeed(self.key, wksht_id=self.wkey, max_row=1)
      colnames = [e.content.text for e in colcells.entry]
      # This would not preserve duplicate column names.
      self.col_name_to_key = dict(zip(colnames, coltags))
    else:
      header_feed = self.client.GetListFeed(self.key, wksht_id=self.wkey, start_index=headerRowNum-1, max_results=1)
      self.col_name_to_key = header_feed.ColumnValueToTagMap()


def DefineFlags():
  usage = u"""usage: %prog [options] [col1:va1 …]"""
  desc = """
Log data into a Google Spreadsheet.

With no further arguments, a list of column tags will be printed to stdout.
These are typically column names in lower case with spaces and special
characters stripped.

Otherwise, remaining arguments should be of the form `columntag:value'.
One row will be added for each invocation of this program.

If you just specify column tags (without a value), then data will be read
from stdin in whitespace delimited form, and mapped to each column in order.
  """
  parser = optparse.OptionParser(usage=usage, description=desc)
  parser.add_option('--debug', dest='debug', action='store_true',
                    help='Enable debug output')
  parser.add_option('--key', '-k', dest='ssid',
                    help='The key of the spreadsheet to update')
  parser.add_option('--sheetid', '-i', dest='wsid',
                    help='The key of the worksheet to update')
  parser.add_option('--name', '-n', dest='ssname',
                    help='The name of the spreadsheet to update or list')
  parser.add_option('--sheet', '-w', dest='wsname',
                    help='The name of the worksheet to update (defaults to the first one)')
  parser.add_option('--list', '-l', dest='listkeys', action='store_true',
                    help='Lists the id of the specified sheet (by --name and --sheet) or all sheets in the specified spreadsheet (by --name) or all availale sheets')
  parser.add_option('--header-row-num', '-a', dest='headerRowNum', type="int",
                    help='Use an alternative row as the header, which allows specifying names in place of tags.')
  return parser


def main():
  parser = DefineFlags()
  (opts, args) = parser.parse_args()
  if ((opts.wsname or opts.wsid) and
      (not opts.ssname and not opts.ssid)):
    parser.error('You must first specify either --name or --key with the --sheet or --sheetid options')
  if (opts.ssname and opts.ssid):
    parser.error('You must specify only one of --name or --key options')
  if (opts.wsname and opts.wsid):
    parser.error('You must specify only one of --sheet or --sheetid options')
  if not opts.listkeys:
    if (not opts.ssname and not opts.ssid):
      parser.error('You must specify either --name or --key options')

  if opts.listkeys:
    lister = LogssAction(debug=opts.debug)
    lister.Authenticate()
    for (ssname, ssid) in lister.GetSpreadsheets(opts.ssid or opts.ssname,
                                                 not opts.ssname):
      print "%s: %s" % (ssname, ssid)
      for (wsname, wsid) in lister.GetWorksheets(ssid,
                                                 opts.wsid or opts.wsname,
                                                 not opts.wsname):
        print "\t%s: %s" % (wsname, wsid)
  else:
    inserter = SpreadsheetInserter(debug=opts.debug)
    inserter.Authenticate()
    ssid = opts.ssid or list(inserter.GetSpreadsheets(opts.ssname))[0][1]
    wsid = (opts.wsid or
            (opts.wsname and
             list(inserter.GetWorksheets(ssid, opts.wsname))[0][1]) or
            'default')
    inserter.key = ssid
    inserter.wkey = wsid

    if opts.headerRowNum:
      inserter.SetColumnHeaderRowNum(opts.headerRowNum)

    if len(args) > 1:
      cols = args
      if inserter.ColumnNamesHaveData(cols):
        inserter.InsertFromColumns(cols)
      else:
        # Read from stdin, pipe data to spreadsheet.
        inserter.InsertFromFileHandle(cols, sys.stdin)
    else:
      print('\n'.join("%s: %s" % (name, tag) for (name, tag) in inserter.ListColumns()))
  return 0


if __name__ == '__main__':
  sys.exit(main())
