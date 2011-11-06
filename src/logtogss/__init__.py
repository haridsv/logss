#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Copyright (c) 2011, "Dominic Mitchell" <dom@happygiraffe.net>
# Copyright (c) 2011, "Hari Dara" <haridara@gmail.com>
# All rights reserved.
# 
# Redistribution and use in source and binary forms, with or without modification, are permitted provided that the
# following conditions are met:
# 
#     Redistributions of source code must retain the above copyright notice, this list of conditions and the following
#       disclaimer.
#     Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following
#       disclaimer in the documentation and/or other materials provided with the distribution.
# 
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES,
# INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
# SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
# SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

# References:
# - Column tags: https://groups.google.com/d/msg/google-spreadsheets-api/dZvMNrfcfQU/OyXBWEZyKdwJ
# - GSX: http://code.google.com/apis/spreadsheets/data/3.0/reference.html#gsx_reference
# - List Feed: http://code.google.com/apis/spreadsheets/data/3.0/developers_guide.html#ListFeeds

"""Log a row to a Google Spreadsheet."""

__author__ = 'Dominic Mitchell <dom@happygiraffe.net>'
__authors__ = ['Hari Dara <haridara@gmail.com>', 'Dominic Mitchell <dom@happygiraffe.net>']
__license__ = 'BSD 2-Clause'
__version__ = '0.1'
__copyright__ = 'Copyright (c) 2011, "Dominic Mitchell" <dom@happygiraffe.net>\nCopyright (c) 2011, "Hari Dara" <haridara@gmail.com>'
__status__ = 'Development'
__maintainer__ = 'Hari Dara'
__email__ = 'haridara@gmail.com'

import pickle
import optparse
import os
import sys
import urllib
import textwrap
import csv

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
               token_store=None, auth_domain=None, logger=None):
    """Construct a new ClientAuthorizer."""
    self.consumer_key = consumer_key
    self.consumer_secret = consumer_secret
    self.scopes = scopes or list(SCOPES)
    self.token_store = token_store or TokenStore()
    self.auth_domain = auth_domain
    self.logger = self.LogToStdout

  def LogToStdout(self, msg):
    print msg

  def FetchAccessToken(self, client):
    # http://code.google.com/apis/gdata/docs/auth/oauth.html#Examples
    httpd = oneshot.ParamsReceiverServer()

    # TODO Find a way to pass "xoauth_displayname" parameter.
    request_token = client.GetOAuthToken(
        self.scopes, httpd.my_url(), self.consumer_key, self.consumer_secret)
    url = request_token.generate_authorization_url(self.auth_domain or 'default')
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
  GSX_NS = gdata.spreadsheets.data.GSX_NAMESPACE

  def ColumnTags(self):
    """Return the names of all child elements in the GSX namespace."""
    return [el.tag for el in self.get_elements(namespace=MyListEntry.GSX_NS)]

  def RowValues(self):
    """Return the values of all child elements in the GSX namespace."""
    return [el.text for el in self.get_elements(namespace=MyListEntry.GSX_NS)]

  def RowValueToTagMap(self):
    """
    Return the values of all child elements keyed by their tags in the GSX namespace.
    Useful on a header row to map the column names to their tags.
    """
    return dict([(el.text, el.tag) for el in self.get_elements(namespace=MyListEntry.GSX_NS)])

class MyListsFeed(gdata.spreadsheets.data.ListsFeed):

  entry = [MyListEntry]

  def ColumnTags(self):
    if not self.entry:
      return []
    return self.entry[0].ColumnTags()

  def ColumnValueToTagMap(self):
    if not self.entry:
      return []
    return self.entry[0].RowValueToTagMap()

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

  def __init__(self, debug=False, auth_domain=None):
    self.debug = debug
    self.auth_domain = auth_domain
    self.client = MySpreadsheetsClient()
    self.client.debug = debug
    self.client.http_client.debug = debug
    self.client.source = os.path.basename(sys.argv[0])

  def Authenticate(self, logger=None):
    client_authz = ClientAuthorizer(logger=logger, auth_domain=self.auth_domain)
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

def make_unique(name, name_map):
  """
  Make the name unique by appending a numeric prefix and return the new unique name.
  Also insert an entry into the name_map so that we can track it.
  """
  sffx = 1
  while True:
    tname = name + str(sffx)
    if tname not in name_map:
      break
    sffx += 1
  name_map[tname] = name
  return tname

def shorten(names, max_len=None):
  """
  Compute shortest unique prefixes and return in the same order.
  FIXME: Doesn't handle identical items.
  """
  name_map = {}
  if max_len:
    # Support for 
    names = [name[:max_len] for name in names]
  else:
    max_len = len(max(names, key=len))
  tbp2 = list(names)
  confl_names = {}
  prfx_len = 1
  while prfx_len <= max_len and tbp2:
    tbp1, tbp2 = tbp2, []
    confl_prfxes = set()
    for name in tbp1:
      prfx = name[:prfx_len]
      if prfx in name_map:
        pname = name_map.pop(prfx)
        if name == pname:
          cnt = confl_names.get(pname, 0)
          confl_names[pname] = cnt+1
        else:
          tbp2.append(pname)
          confl_prfxes.add(prfx)
      if name in confl_names:
        cnt = confl_names.get(name, 0)
        confl_names[name] = cnt+1
      elif prfx in confl_prfxes:
        tbp2.append(name)
      else:
        name_map[prfx] = name
    prfx_len += 1
  # Inverse the dictionary
  name_map = dict((name, prfx) for (prfx, name) in name_map.items())
  return [name in confl_names and make_unique(name, name_map) or name_map[name]
          for name
          in names]

class SpreadsheetInserter(LogssAction):
  """A utility to insert rows into a spreadsheet."""

  def __init__(self, debug=False, auth_domain=None):
    super(SpreadsheetInserter, self).__init__(debug, auth_domain)
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

  def InsertFromFileHandle(self, cols, fh, csvformat=False, verbose=False):
    if verbose:
      print >> sys.stderr, 'Columns selected: ' + str(cols)
    if csvformat:
        fh = csv.reader(fh)
    for line in fh:
      if csvformat:
          vals = line
      else:
          vals = line.rstrip().split(None, len(cols) - 1)
      data = dict(zip(cols, vals))
      if verbose:
        print >> sys.stderr, 'Inserting row: ' + str(vals)
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

  def expand_col_names(self, col_cells, shortenColumnNames=False, maxLen=None):
    def gen_col_names(col_names):
      last_seen_col = None
      for col_name in col_names:
        if not col_name and not last_seen_col:
          yield ''
        elif col_name:
          last_seen_col = col_name
          yield last_seen_col
        else:
          yield last_seen_col

    if isinstance(col_cells, list) and col_cells:
      # If list, we expect contigous cells, but some of them could be None's.
      known_col_names = shortenColumnNames and shorten(col_cells, maxLen) or col_cells
    elif col_cells.entry:
      col_names = [e.content.text for e in col_cells.entry]
      col_nums = [int(e.get_elements('cell')[0].get_attributes('col')[0].value)
                  for e in col_cells.entry]
      if shortenColumnNames:
        col_names = shorten(col_names, maxLen)
      col_num_map = dict(zip(col_nums, col_names))
      known_col_names = [num in col_num_map and col_num_map[num] or None
                         for num
                         in xrange(min(col_nums), max(col_nums)+1)]
    else:
      return []
    # This doesn't fill the voids after the last cell, since we don't know the
    # actual range. But this is not a problem as it would extended.
    return list(gen_col_names(known_col_names))

  def SetColumnHeaderRowNums(self, startHeaderRowNum, endHeaderRowNum=None, shortenColumnNames=False, maxLen=None):
    header_rows = []
    coltags = None
    # Special case for first row, we need to use cells feed to get to it.
    if startHeaderRowNum == 1:
      colcells = self.client.GetCellsFeed(self.key, wksht_id=self.wkey, max_row=1)
      names = self.expand_col_names(colcells, shortenColumnNames, maxLen)
      if not names:
        raise Exception("Header row 1 is empty")
      header_rows.append(names)
      startHeaderRowNum += 1
    if endHeaderRowNum and endHeaderRowNum >= startHeaderRowNum:
      header_row_feed = self.client.GetListFeed(self.key, wksht_id=self.wkey,
                                            start_index=startHeaderRowNum-1,
                                            max_results=(endHeaderRowNum - startHeaderRowNum + 1))
      for header_row in header_row_feed.entry:
        names = self.expand_col_names(header_row.RowValues())
        if not names:
          raise Exception("A header row between %s and %s is empty" %
                          (startHeaderRowNum, endHeaderRowNum))
        header_rows.append(shortenColumnNames and shorten(names, max_len=maxLen) or names)
        if not coltags:
          coltags = header_row.ColumnTags()
    if not coltags:
      list_feed = self.client.GetListFeed(self.key, wksht_id=self.wkey, max_results=1)
      coltags = list_feed.ColumnTags()
    if len(header_rows) > 1:
      # Make sure all the headers are of the same length.
      max_header_len = max([len(header_row) for header_row in header_rows])
      #for header_row in header_rows:
      #  if len(header_row) < max_header_len:
      #    header_row.extend(header_row[-1:] * (max_header_len - len(header_row)))
      header_rows = [header_row + header_row[-1:] * (max_header_len - len(header_row))
                     for header_row in header_rows]
    qual_col_names = ['.'.join([h for h in header_path if h]) for header_path in zip(*header_rows)]
    self.col_name_to_key = dict(zip(qual_col_names, coltags))

def alt_header_nums(option, opt_str, value, parser):
  num_range = value.strip().split('-')
  valid = False
  if num_range and len(num_range) < 3:
    try:
      setattr(parser.values, option.dest,
              sorted([int(num.strip()) for num in num_range]))
      valid = True
    except ValueError:
      pass
  if not valid:
    raise optparse.OptionValueError('%s only accepts a number or a single range of row numbers: e.g., 2 or 2-3' %
                                    option.get_opt_string())

class BetterDescOptionParser(optparse.OptionParser):
    """
    Format description without loosing the paragraphs.
    """

    def format_description(self, formatter):
        wrapper = textwrap.TextWrapper(width=self.formatter.width)
        return "\n".join(["\n".join(wrapper.wrap(p))
                          for p in self.expand_prog_name(
                              self.description).split("\n")])

def DefineFlags():
  usage = u"""usage: %prog [options] [col1:va1 …]"""
  desc = textwrap.dedent("""\
        Log data into a Google Spreadsheet.

        With no further arguments, a list of column tags will be printed to stdout.
        These are typically column names in lower case with spaces and special
        characters stripped.

        Otherwise, remaining arguments should be of the form `columntag:value'.
        One row will be added for each invocation of this program.

        If you just specify column tags (without a value), then data will be read
        from stdin in whitespace (or comma, if -c option is used) delimited form,
        and mapped to each column in order.
      """)
  parser = BetterDescOptionParser(usage=usage, description=desc)
  parser.add_option('--debug', dest='debug', action='store_true',
                    help='Enable debug output')
  parser.add_option('--verbose', dest='verbose', action='store_true',
                    help='Enable verbose output')
  parser.add_option('--domain', '-d', dest='domain',
                    help='Specify an apps domain for authentication')
  parser.add_option('--csvformat', '-c', dest='csvformat', action='store_true',
                    help='Specifies that the stdin is in CSV format')
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
  parser.add_option('--alt-header', '-a', dest='headerRowNums',
                    type='string',
                    action='callback',
                    callback=alt_header_nums,
                    help='Use an alternative row as the header, which allows specifying names in place of tags.')
  parser.add_option('--short-header-names', '-s', dest='shorten', action='store_true',
                    help='Shorten the names of the headers so that it is easier to type. This also has the benefit of making sure they are unique.'),
  parser.add_option('--max-header-len', '-m', dest='maxHeaderLen', type='int', default=3,
                    help='When using -s option, specify the max length or 0 to disable length restriction.'),
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
    lister = LogssAction(debug=opts.debug, auth_domain=opts.domain)
    lister.Authenticate()
    for (ssname, ssid) in lister.GetSpreadsheets(opts.ssid or opts.ssname,
                                                 not opts.ssname):
      print "%s: %s" % (ssname, ssid)
      for (wsname, wsid) in lister.GetWorksheets(ssid,
                                                 opts.wsid or opts.wsname,
                                                 not opts.wsname):
        print "\t%s: %s" % (wsname, wsid)
  else:
    inserter = SpreadsheetInserter(debug=opts.debug, auth_domain=opts.domain)
    inserter.Authenticate()
    ssid = opts.ssid or list(inserter.GetSpreadsheets(opts.ssname))[0][1]
    wsid = (opts.wsid or
            (opts.wsname and
             list(inserter.GetWorksheets(ssid, opts.wsname))[0][1]) or
            'default')
    inserter.key = ssid
    inserter.wkey = wsid

    if opts.headerRowNums:
      inserter.SetColumnHeaderRowNums(opts.headerRowNums[0],
                                      len(opts.headerRowNums) > 1 and opts.headerRowNums[1] or None,
                                      shortenColumnNames=opts.shorten,
                                      maxLen=opts.maxHeaderLen)

    if len(args) > 1:
      cols = args
      if inserter.ColumnNamesHaveData(cols):
        inserter.InsertFromColumns(cols)
      else:
        # Read from stdin, pipe data to spreadsheet.
        inserter.InsertFromFileHandle(cols, sys.stdin, csvformat=opts.csvformat, verbose=opts.verbose)
    else:
      print('\n'.join("%s: %s" % (name, tag) for (name, tag) in inserter.ListColumns()))
  return 0

if __name__ == '__main__':
  sys.exit(main())

# vim: sw=2 et
