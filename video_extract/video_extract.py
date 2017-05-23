from __future__ import print_function

from openpyxl import load_workbook

import datetime
import sys
import argparse

def main(argv):

	# Get parameters from command line
	parser = argparse.ArgumentParser('Interact with xls files')
	parser.add_argument('file', action='store', help='XLS file')
	parser.add_argument('--output', action='store', help='Output file, 1 sequence per line', required=True)

	args = parser.parse_args()

	conducteur = args.file
	file = args.output

	print('Analyzing "{}" XLS file...'.format(conducteur), end='\n')

	wb = load_workbook(conducteur, keep_vba=True, data_only=True)
	# TEST use this to display the sheets
	# print(wb.get_sheet_names())

	# selecting the relevant sheet
	ws = wb['Conducteur']
	# getting the number of lines
	max_row = ws.max_row

	# creating an output file
	output = open(file, 'w')
	

	# processing the file
	print('Processing {} lines from the file.'.format(max_row), end='')
	for i in range (3, max_row):
		a3 = ws['A' + str(i)].value
		a4 = ws['C' + str(i)].value
		print('.',end='')
		sys.stdout.flush()
		if a3 == None:
			pass
		else:
			a4bis = a4.strftime('%-H|%-M|%-S')
			#print(str(a3) + "|" + str(a4bis) + "|")
			output.write(str(a3) + "|" + str(a4bis) + "|\n")

	# closing file
	output.close()


if __name__ == '__main__':
	sys.exit(main(sys.argv))
