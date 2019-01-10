# $Author: lchu $
# $Revision: 1.2 $
# $Date: 2012/11/14 11:30:00 $

# This module generate sv code from reg info in an .xls file

import ComExcel
import os
import time
import re
import random

fileHead = '// Code auto-generated at %s\n'%time.ctime() + '// Author: lchu\n'

def findLatestVersion(head):
	''' Get the name of the latest file within a series of versions
	whose name all start with argument head. '''
	files = []
	for filename in os.listdir(os.getcwd()):
		if filename.startswith(head):
			files.append(filename)
	if len(files):
		return os.path.join(os.getcwd(), max(files))
	else:
		raise Exception, 'File that starts with %s not found'%head

def extractReg(xlsFile = findLatestVersion('db_register_table')):
	'''
	Extract all regs' info from the xls file and store it into a database.
	Put the database into a py file, also return the database.
	An example of a reg's info in dict format: {
		'name':'rid', 'func':0, 'byteAddrM4':0, 'bitAddr':[8,15],
		'comment':'Revision Id', 'default':0, 'attribs':['RW',] }
	'''
	print '\n', '*'*70, '\nStart extract info from %s'%xlsFile
	bregDb = []
	fileObj = ComExcel.ExcelComObj(filename=xlsFile)
	sheetNames = []
	for i in range(fileObj.xlBook.Sheets.Count):
		sheetNames.append(fileObj.xlBook.Sheets(i+1).Name)
	for sheetName in sheetNames:
		print '\n--- Converting %s ---'%sheetName
		funcNum = int(sheetName.strip('Ffunction'))
		sheetObj = ComExcel.ExcelComObj(sheetnum=sheetName, filename=xlsFile)
		startRow = 5; rangeRow = 16
		startCol = 2; rangeCol = 8
		for row in range(startRow,startRow+rangeRow):
			for col in range(startCol,startCol+rangeCol):
				regName = sheetObj.getCellText(row,col).lower()
				if regName in [u'', u'reserved']: continue
				import sys; sys.stdout.write('.')

				regAddr = 4*(row-startRow)
				regEndBitAddr = 31 - (col - startCol)
				if sheetObj.checkMerge(row,col) == False:
					regStartBitAddr = 31 - (col - startCol)
				else:
					width = 1
					while sheetObj.checkMerge(row,col+width)==True and \
							sheetObj.getCellText(row,col+width)==u'':
						width += 1
					regStartBitAddr = 31 - (col - startCol) - (width - 1)

				regComment = sheetObj.getComment(row,col)
				is_attrib_found = False
				default_val = 0 # it's 0 if no default found in comment
				for line in regComment.splitlines():
					if 'Attribute' in line:
						is_attrib_found = True
						attribList = line.split(':')[1].split('/')
						for i in range(len(attribList)): 
							attribList[i] = attribList[i].strip().upper()
					elif 'Default :' in line:
						valExpr = line.split(':')[1].strip().lower()
						mobj = re.match(r'\d*\Wb([0-1]+)\S*', valExpr)
						if mobj:
							default_val = int(mobj.group(1), 2)
						else:
							mobj = re.match(r'\d*\Wd([0-9]+)\S*', valExpr)
							if mobj:
								default_val = int(mobj.group(1), 10)
							else:
								mobj = re.match(r'\d*\Wh([0-9a-f]+)\S*', valExpr)
								if mobj:
									default_val = int(mobj.group(1), 16)
								else:
									default_val = int(valExpr)
				if not is_attrib_found:
						raise Exception, '%s attrib not found'%regName

				reg = { 'name':regName, 'func':funcNum, 'byteAddrM4':regAddr, \
						'bitAddr':[regStartBitAddr,regEndBitAddr], 'comment':regComment, \
						'default':default_val, 'attribs':attribList}
				bregDb.append(reg)
	#sheetObj.close()

	print '\nGenerating breg_db.py'
	with open('breg_db.py','w') as f:
		print>>f, fileHead.replace(u'//', u'#'),
		print>>f, '# Registrer info database'
		print>>f, 'bregDb =', bregDb
	return bregDb


def genCode(pattern):
	'''Generate repetitive code from the database'''
	code = ''
	import breg_db; reload(breg_db)
	for reg in breg_db.bregDb:
		code += pattern(reg) + '\n'
	return code

def genFile(outFile, code, noteTxt):
	print 'Generating %s'%outFile
	with open(outFile,'w') as f:
		f.write(fileHead)
		f.write(noteTxt+'\n\n')
		f.write(code.encode('ascii','ignore'))

def genRegName2Addr(outFile='breg_name2addr.sv'):
	def myPattern(reg):
		'''convert info to code'''
		codePiece = "parameter bit[0:3][7:0] %s = { 8'd%d, 8'h%x, 8'd%d, 8'd%d };\n" \
				%('fn%d_%s'%(reg['func'],reg['name']), reg['func'], reg['byteAddrM4'], \
				1+reg['bitAddr'][1]-reg['bitAddr'][0], reg['bitAddr'][0])
		for line in reg['comment'].splitlines(True):
			codePiece += '// ' + line
		return codePiece

	myNoteTxt = '''
// This file is a register name->address table.
// format:
// breg_name = ( function, byte_addr, bit_width, start_bit_addr )
'''
	genFile(outFile, genCode(myPattern), myNoteTxt)

def run():
	extractReg()
	genRegName2Addr()

if __name__ == '__main__':
	try:
		run()
	except:
		import traceback; traceback.print_exc();
	finally:
		raw_input("\n---Press ENTER to quit---")

