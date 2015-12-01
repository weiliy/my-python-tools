#!/usr/bin/env python 

import os
import argparse
import xlrd
import re

parser = argparse.ArgumentParser()
parser.add_argument("ipschema", type=str,
		help="ipschema the filepath of the ip schema excel")
parser.add_argument("rdg", type=str,
		help="the filepath of the rdg file output")
parser.add_argument("-a","--address_col", type=int, default=0,
		help="The col number of the Server IP address in ip schema. Default value is 0")
parser.add_argument("-n", "--name_col", type=int, default=1,
		help="The col number of the Server Name in ip schema. Default value is 1")
parser.add_argument("-f", "--filter", default='.*',
		help="Filter to the Server Name, support Regular Expression")
parser.add_argument("-s", "--sheet", type=str, default='.*',
		help="The sheet filter, support Regular Expression")
parser.add_argument("-v","--verbosity", action="count", default=0,
		help="increase output verbosity")
args = parser.parse_args()

if args.verbosity:
	print "verbosity turned on"
	print "#######################################"
	print "# ipschema:", args.ipschema
	print "# address_col:", args.address_col
	print "# name_col:", args.name_col
	print "# filter:", args.filter
	print "# rdg", args.rdg
	print "# sheet", args.sheet
	print "#######################################"
	print


def read_server_from_sheet(sheet, servers, name_pat):
	for row_index in range(sheet.nrows):
		server = (unicode(sheet.cell(row_index, args.address_col).value),
				unicode(sheet.cell(row_index, args.name_col).value))
		if re.search(name_pat, server[1]):
			servers.append(server)

def output_rdg(servers, rdg): 
	rdg_template = r'<?xml version="1.0" encoding="utf-8"?><RDCMan programVersion="2.7" schemaVersion="3"><file><credentialsProfiles /><properties><expanded>False</expanded><name>%rdg_name%</name></properties>%servers%</file><connected /><favorites /><recentlyUsed /></RDCMan>'
	
	server_template = r'<server><properties><displayName>%display_name%</displayName><name>%connect_name%</name></properties></server>'
	
	rdg_servers = ""
	
	for address, name in servers:
		temp = server_template
		temp = temp.replace('%display_name%', name)
		temp = temp.replace('%connect_name%', address)
		rdg_servers += temp
	
	rdg_file = rdg_template.replace('%rdg_name%', rdg)
	rdg_file = rdg_template.replace('%servers%', rdg_servers)
	
	f = open(rdg, 'w')
	f.write(rdg_file)
	f.close()

def read_server_from_book(book, sheet_pat, name_pat):
	for sheet in book.sheets():
		print sheet.name, "-----",
		if re.search(sheet_pat, sheet.name):
			print 'read', 
			read_server_from_sheet(sheet, servers, name_pat)
		else:
			print 'not read'

# Main
books= []
servers = []
sheet_pat = re.compile(args.sheet)
name_pat = re.compile(args.filter)

if os.path.isfile(args.ipschema):
	books.append(xlrd.open_workbook(args.ipschema))
elif os.path.isdir(args.ipschema):
	for filename in os.listdir(args.ipschema):
		fullpath = os.path.join(args.ipschema, filename)
		print 'Loading', fullpath,
		try:
			books.append(xlrd.open_workbook(fullpath))
		except xlrd.biffh.XLRDError:
			print 'failed'
			continue
		print 'OK'
else:
	print 'Not file or directory:', args.ipschema
	exit()

for book in books:
	print
	print "------------------------------------------"
	read_server_from_book(book, sheet_pat, name_pat)
	print "------------------------------------------"

output_rdg(servers, args.rdg)
