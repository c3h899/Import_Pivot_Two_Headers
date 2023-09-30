import pandas as pd
import pprint
import os
import openpyxl as opxl

def my_float(num):
	# Exception Free
	if isinstance(num, float):
		fl = num
	elif isinstance(num, str):
		num = num.strip(' ±%')
		num = num.replace(',', '')
		try:
			fl = float(num)
		except:
			fl = float('nan')
	else:
		# TODO: Support All the Things!
		print(type(fl))
		fl = num
	return fl

# https://docs.python.org/3/library/pprint.html
pp = pprint.PrettyPrinter(indent=4)

# List of Excel files (update this with your file paths)
import_dir = "import_dir"
# https://www.geeksforgeeks.org/python-list-files-in-a-directory/
# https://stackoverflow.com/questions/35510787/python-endswith-with-multiple-string
valid_excel = (".xlsx", ".xls")
valid_csv = (".csv")
Tables = [] # lists are cheap to add and delete, but expensive to search

# Import table(s) and convert them to high-quality data.
for x in os.listdir(import_dir):
	path_name = f"{import_dir}\\{x}"
		# https://www.w3schools.com/python/gloss_python_escape_characters.asp
		# https://realpython.com/python-f-strings/
	if x.endswith(valid_excel):
		df = pd.read_excel(path_name, sheet_name=1, header=[0,1])
		df.drop(df.columns[0], axis=1, inplace=True) # Dump the 'Label' column
		# Multi-Index Based on Cyclic Sub-Header
		col_head0 = []
		col_head0_set = set()
		col_head1 = []
		col_head1_set = set()
		for (head0, head1) in df.keys(): # Recover the unique labels, while preserving order
			if not head0 in col_head0_set:
				col_head0_set.add(head0)
				col_head0.append(head0) # Care is taken to preserve order
			if not head1 in col_head1_set:
				col_head1_set.add(head1)
				col_head1.append(head1) # Care is taken to preserve order
		# Ensure Entries are numeric.
		df2 = df.applymap(my_float)
		# Turn the Label Column into multiple indexes
		indexes = []
		level_0 = []
		heiarchy = ['Anarchy']
		prev_indent = 0
		max_indent = 0
		row = -2 # Compensated for 2 header rows
		for cell in opxl.load_workbook(path_name)['Data']['A']:
			# https://stackoverflow.com/questions/30746699/openpyxl-alignment-indent-vs-ident
			# https://stackoverflow.com/questions/34754077/openpyxl-how-to-read-only-one-column-from-excel-file-in-python
			indent_level = cell.alignment.indent
			label = cell.value
			if(indent_level == 0):
				level_0.append(row)
			if(indent_level > max_indent):
				max_indent = indent_level
			if(indent_level == prev_indent):
				heiarchy.pop()
				heiarchy.append(label)
			elif(indent_level < prev_indent):
				heiarchy.pop()
				heiarchy.pop()
				heiarchy.append(label)
			else: # (indent_level > prev_indent)
				heiarchy.append(label)
			prev_indent = indent_level
			indexes.append(list(heiarchy))
				# Not Redundent; Ensures a Unique copy of the data is made, and not a copy of the existing reference
			row = row + 1
		indexes.pop(0) # Remove 1st Header Row
		indexes.pop(0) # Remove 2nd Header Row
		headings = [f'H{x}' for x in range(int(max_indent + 1))]
		labels = pd.DataFrame(indexes, columns=headings)
		depth_labels = len(labels.columns)
		# Merge the Row-Labels & Data
		merge = pd.concat([labels, df2], axis=1)
		merge.drop(level_0, inplace=True)
		merge.set_index(headings, inplace=True) # Bydefault inplace=False
		# Multi-Index Columns
		merge.columns = pd.MultiIndex.from_product([col_head0, col_head1], names=['A', 'B'])
		# Export
		rec = dict(fname = x, data = merge, ftype = "Excel")
			# https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
			# As of Version 1.2 engine will be automatically specified based on extension
			# unless explicitly defined.
		# I have a seething hatred for the person whom formated the 1st column
		Tables.append(rec)
	elif x.endswith(valid_csv):
		df = pandas.read_csv(path_name)
		rec = dict(fname = x, data = df, ftype = "CSV")
		Tables.append(rec)
		# TODO: Complete this as needed

# Do Things with the data
for table in Tables:
	# File Names
#	print(table)
	fname = table['fname'].split('.')
	pre = fname[0][0:6] # Range is stupid: it's [0, 6) 0 is inclusive; 6 is excluded.
	post = fname[1][0:4]
	year = fname[0][-4:] # Just to demonstrate that you can start from the end and go backwards
	table['year'] = year
	ofname = f"{year}_{pre}-{post}.csv"
#	print(ofname)
	# https://stackoverflow.com/questions/53927460/select-rows-in-pandas-multiindex-dataframe
	# Select Columns
	Est = merge.xs('Estimate', level='B', axis=1)
	# Select Rows
	A = Est.xs('Total population', axis=0, level=1, drop_level=False)
	B = Est.xs('Total housing units', axis=0, level=2, drop_level=False)
	# Stitch
	St = pd.concat([A, B])
#	print(St)
	St.to_csv(ofname)
