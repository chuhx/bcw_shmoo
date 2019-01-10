# $Author: lchu $
# $Revision: 1.42 $
# $Date: 2015/03/27 11:30:02 $

# This module generate sv code from bcw info in an .xls file

import ComExcel
import os
import time
import re
import random
import copy

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

def extractReg(xlsFile = findLatestVersion('bcw_info')):
	'''
	Extract all bcws' info from the xls file and store it into a database.
	Put the database into a py file (bcw_db.py), also return the database.
	'''
	print '\n', '*'*70, '\nStart extract info from %s'%xlsFile
	bcwDb = {}
	for sheetName in ['4-bit', '8-bit']:
		print '\n--- Converting %s ---'%sheetName
		sheetObj = ComExcel.ExcelComObj(sheetnum=sheetName, filename=xlsFile)
		startRow = 2; startCol = 1;
		row = startRow;
		savedBcwId = ''
		while True:
			import sys; sys.stdout.write('.')
			bcwId = sheetObj.getCellText(row, startCol).strip()
			if bcwId == 'EOF': break
			elif bcwId == '': bcwId = savedBcwId
			else: savedBcwId = bcwId
			if sheetObj.getCellText(row, startCol+1).strip() == '': row += 1; continue;

			mobj = re.match('F([0-7])BC([0-9A-F])([0-9A-FxX])', bcwId)
			if not mobj:
				raise Exception('%s is an invalid BCW'%bcwId)
			func = int(mobj.groups()[0])
			if mobj.groups()[2] in ['x', 'X']:
				addr = int(mobj.groups()[1], 16) << 4
			else:
				addr = int(mobj.groups()[1]+mobj.groups()[2], 16)
			if func not in bcwDb.keys(): bcwDb[func] = {}
			if addr not in bcwDb[func].keys(): bcwDb[func][addr] = []
			field = {
					'name':       sheetObj.getCellText(row, startCol+1).lower(),
					'default':    int(sheetObj.getCellText(row, startCol+2)),
					'attrib':     sheetObj.getCellText(row, startCol+3),
					'bit_start':  int(sheetObj.getCellText(row, startCol+4)),
					'bit_len':    int(sheetObj.getCellText(row, startCol+5)),
					'is_context': sheetObj.getCellText(row, startCol+6),
					'dut_probe':  sheetObj.getCellText(row, startCol+7),
					'comment':  sheetObj.getCellText(row, startCol+8),
					}
			bcwDb[func][addr].append(field)
			row += 1

	#for func in bcwDb.keys():
		#for addr in bcwDb[func].keys():
			#for field in bcwDb[func][addr]:
				#print func, '%02X'%addr, field
	#sheetObj.close()

	print '\nGenerating bcw_db.py'
	with open('bcw_db.py','w') as f:
		print>>f, fileHead.replace(u'//', u'#'),
		print>>f, '# Registrer info database'
		print>>f, 'bcwDb =', bcwDb
	return bcwDb


def genCode(pattern):
	'''Generate repetitive code from the database'''
	code = ''
	from bcw_db import bcwDb
	for func in bcwDb.keys():
		for addr in bcwDb[func].keys():
			for field in bcwDb[func][addr]:
				line = pattern(func, addr, field) + '\n'
				if line.strip() != '':
					code += line
	return code

def genFile(outFile, code, noteTxt):
	print 'Generating %s'%outFile
	with open(outFile,'w') as f:
		f.write(fileHead)
		f.write(noteTxt+'\n')
		f.write(code.encode('ascii','ignore'))


def genUpdateBcur(outFile='update_bcur.sv'):
	updBcwCode = 'function void update_bcur(int func, int addr, int val, bit is_effective=1);\n'
	updBcwCode += 'case (func)\n'
	from bcw_db import bcwDb
	for func in bcwDb.keys():
		updBcwCode += '\t%d: case (addr)\n'%func
		for addr in sorted(bcwDb[func].keys()):
			updBcwCode += "\t\t\t8'h%02x: begin\n"%addr
			filledBits = []
			for field in bcwDb[func][addr]:
				filledBits.append(copy.copy([field['bit_start'], field['bit_len']]))
				valExpr = "(val >> %d) & 'h%x"%(field['bit_start'], (1<<field['bit_len'])-1)
				#if field['attrib'] != 'RO':
				if field['attrib'] in ['RW', 'WO']:
					if field['is_context'] == 'yes':
						updBcwCode +=\
'''				if (is_effective) bcur_name = val_expr;
				if (bcur_context_4_op_tran==0) bcur_name_c0 = val_expr;
				else bcur_name_c1 = val_expr;
'''.replace('name', field['name']).replace('val_expr', valExpr)
					else:
						updBcwCode += "\t\t\t\tbcur_%s = %s;\n"%(field['name'], valExpr)
				elif field['attrib'] in ['RWC', ]:
					updBcwCode += '\t\t\t\tif((val_expr) == 1) bcur_name = default_val;\n'.replace('val_expr', valExpr).replace('name', field['name']).replace('default_val', str(field['default']))
			mask = 0
			for start,len in filledBits:
				mask |= (1<<(start+len)) - (1<<start)

			updBcwCode += "\t\t\tend\n"
		updBcwCode += '\t\tendcase\n'
	updBcwCode += 'endcase\n'
	updBcwCode += 'endfunction\n\n'


	getBcwCode = '// Patch bcur_xx fields into one entire bcw.\n'
	getBcwCode += 'function automatic bit[7:0] get_bcw(int func, int addr);\n'
	getBcwCode += "bit[7:0] bcw=8'h0;\n"
	getBcwCode += 'case (func)\n'
	for func in bcwDb.keys():
		getBcwCode += '\t%d: case (addr)\n'%func
		for addr in sorted(bcwDb[func].keys()):
			getBcwCode += "\t\t\t8'h%02x: begin\n"%addr
			for field in bcwDb[func][addr]:
				if re.match('func_sel_f\d', field['name']) or field['name']=='cmd_space_ctrl' or field['attrib']=='WO':
					continue
				else:
					getBcwCode += '\t\t\t\tbcw |= bcur_%s << %0d;\n'%(field['name'], field['bit_start'])
			getBcwCode += "\t\t\tend\n"
		getBcwCode += '\t\tendcase\n'
	getBcwCode += 'endcase\n'
	getBcwCode += 'return bcw;\n'
	getBcwCode += 'endfunction\n\n'
	#with open('o.sv','w') as f:
		#print>>f, getBcwCode


	def switchPattern(func, addr, field):
		if field['is_context'] == 'no':
			return '\n\t// %s is not context-relevant.'%field['name']
		codePiece = '''
	if (bcur_context_4_op_tran == 0) bcur_name = bcur_name_c0;
	else bcur_name = bcur_name_c1;'''
		return codePiece.replace('name', field['name'])

	switchCode = 'function void db_switch_context();\n'\
			+ genCode(switchPattern)\
			+ 'endfunction\n\n'

	def cmpPattern(func, addr, field):
		if field['attrib'] in ['WO', ] or re.match('func_sel_f\d', field['name']):
			return '\n\t// %s is not suitable to compare.'%field['name']
		global funcSelIsHere
		if field['name'] == 'func_sel':
			if funcSelIsHere: return ''
			else: funcSelIsHere = True
		codePiece = ''
		if field['attrib'] == 'RW': prefix = 'rf_'
		elif field['attrib'] in ['RO','RWC']: prefix = 'i_rf_'
		else: raise Exception, 'Unexpected attrib %s'%field['attrib']
		if not ('err_' in field['name']):
			# A1 metal change: mpr override mode is set to 1 internally if entering HIW
			borrowedByHiw = ['mpr_ov_hiw_en', 'mpr0']
			pfUsedInTrainMode = ['cal_comp_pass', 'b0_sts', 'b1_sts', 'b2_sts', 'b3_sts', 'b4_sts', 'b5_sts', 'b6_sts', 'b7_sts', ]
			if field['name'] in borrowedByHiw:
				codePiece += '\n\t\tif(bcur_tran_mode !== 8) begin'
			elif field['name'] in pfUsedInTrainMode:
				codePiece += '\n\t\t`ifndef BCOM_INJECT_ERR'
			for probe in field['dut_probe'].split(','):
				if field['name'] == 'cas_latency':
					codePiece += '''
		if(bcur_cas_latency !== {probe.rf_cas_latency, probe.rf_cal_a12})
			$display("%t BCW_CMP ERROR bcur_name %x probe/pr_name %x", $realtime, bcur_name, {probe.rf_cas_latency, probe.rf_cal_a12});'''.replace('probe', '`DB_RF%d'%int(probe)).replace('pr_', prefix).replace('name', field['name'])
				elif field['name'] == 'f3bc6x_7_':
					codePiece += '''
		if(bcur_name !== probe.pr_name)
			$display("%t BCW_CMP ERROR bcur_name %x probe/pr_name %x", $realtime, bcur_name, probe.pr_name);'''.replace('probe', '`DB_RF%d'%int(probe)).replace('pr_', prefix).replace('name', field['name']).replace('rf_f3bc6x_7_', 'rf_f3bc6x[7]')
				else:
					codePiece += '''
		if(bcur_name !== probe.pr_name)
			$display("%t BCW_CMP ERROR bcur_name %x probe/pr_name %x", $realtime, bcur_name, probe.pr_name);'''.replace('probe', '`DB_RF%d'%int(probe)).replace('pr_', prefix).replace('name', field['name'])
			if field['name'] in borrowedByHiw:
				codePiece += '\n\t\tend'
			elif field['name'] in pfUsedInTrainMode:
				codePiece += '\n\t\t`endif'
			if field['name'] == 'func_sel':
				codePiece = codePiece.replace('rf_func_sel', 'rf_bc7x')
		return codePiece

	global funcSelIsHere; funcSelIsHere = False
	cmpCode = '''
// Compare bcur_name to dut internal name, except error flags and logs
function void cmp_bcur_dut();
// `ifndef BCOM_INJECT_ERR
if(!during_calibration && !is_op_speed_ever_equal_to_wrong_val_in_pba && !are_err_chk_ens_ever_changed_in_err_status && !is_ever_bcwwr_in_pba && !is_cus_pattern_ever_matched && !is_f3bc6x_ever_written) begin
''' + genCode(cmpPattern) + '''
end
// `endif //BCOM_INJECT_ERR
endfunction\n
'''
	def cmpErrPattern(func, addr, field):
		codePiece = ''
		prefix = 'i_rf_'
		if 'err_' in field['name']:
			#print field['name']
			for probe in field['dut_probe'].split(','):
				codePiece += '''
	if(bcur_name !== probe.pr_name)
		$display("%t BCW_CMP ERROR bcur_name %x probe/pr_name %x", $realtime, bcur_name, probe.pr_name);'''.replace('probe', '`DB_RF%d'%int(probe)).replace('pr_', prefix).replace('name', field['name'])
		# else:
			# codePiece = '\t//%s is not err flags or logs'%(field['name'])
		return codePiece

	def cmpErrFlagPattern(func, addr, field):
		codePiece = ''
		prefix = 'i_rf_'
		if 'err_flag' in field['name'] or 'err_gt1' in field['name']:
			for probe in field['dut_probe'].split(','):
				codePiece += '''
	if(bcur_name !== probe.pr_name)
		$display("%t BCW_CMP ERROR bcur_name %x probe/pr_name %x", $realtime, bcur_name, probe.pr_name);'''.replace('probe', '`DB_RF%d'%int(probe)).replace('pr_', prefix).replace('name', field['name'])
		return codePiece

	def cmpErrLogPattern(func, addr, field):
		codePiece = ''
		prefix = 'i_rf_'
		if 'err_log' in field['name']:
			for probe in field['dut_probe'].split(','):
				codePiece += '''
	if(bcur_name !== probe.pr_name)
		$display("%t BCW_CMP ERROR bcur_name %x probe/pr_name %x", $realtime, bcur_name, probe.pr_name);'''.replace('probe', '`DB_RF%d'%int(probe)).replace('pr_', prefix).replace('name', field['name'])
		return codePiece

	cmpErrCode = '''
// Compare error flags and logs of bcur_xxx and dut_xxx
function void cmp_err_flags_and_logs();
if(!during_calibration && !is_op_speed_ever_equal_to_wrong_val_in_pba && !are_err_chk_ens_ever_changed_in_err_status && !is_ever_bcwwr_in_pba && !is_cus_pattern_ever_matched && !is_f3bc6x_ever_written) begin
''' + genCode(cmpErrPattern) + '''
end
endfunction\n

// Compare error flags of bcur_xxx and dut_xxx
function void cmp_err_flags();
if(!during_calibration && !is_op_speed_ever_equal_to_wrong_val_in_pba && !are_err_chk_ens_ever_changed_in_err_status && !is_ever_bcwwr_in_pba && !is_cus_pattern_ever_matched && !is_f3bc6x_ever_written) begin
''' + genCode(cmpErrFlagPattern) + '''
end
endfunction\n

// Compare error logs of bcur_xxx and dut_xxx
function void cmp_err_logs();
if(!during_calibration && !is_op_speed_ever_equal_to_wrong_val_in_pba && !are_err_chk_ens_ever_changed_in_err_status && !is_ever_bcwwr_in_pba && !is_cus_pattern_ever_matched && !is_f3bc6x_ever_written) begin
''' + genCode(cmpErrLogPattern) + '''
end
endfunction\n

'''

	noteTxt = '''
// DB monitor uses the update_*() funcs to update bcur_name.
// 4-bit addr is like: 'h00, 'h0a 
// 8-bit addr is like: 'h00, 'h10, 'hf0
'''
	code = updBcwCode + getBcwCode + switchCode + cmpCode + cmpErrCode
	genFile(outFile, code, noteTxt)


global funcSelIsHere

def genBcwName2Addr(outFile='bcw_name2addr.sv'):
	def myPattern(func, addr, field):
		global funcSelIsHere
		if field['name'] == 'func_sel':
			if funcSelIsHere: return ''
			else: funcSelIsHere = True
		if func == 0:
			if addr <= 0x0f: bit_width = 4
			else: bit_width = 8
		else: bit_width = 8
		codePiece = "parameter bit[0:4][7:0] ba_%s = { 8'd%d, 8'h%02x, 8'd%d, 8'd%d, 8'd%d };\n"\
				%(field['name'], func, addr, bit_width, field['bit_start'], field['bit_len'])
		codePiece += '// Name: ' + field['name'].upper() + '\n' +\
						 '// Attribute: ' + field['attrib'] + '\n' +\
						 '// Default: ' + str(field['default']) + '\n' +\
						 '// Comment: ' + '\n'
		for line in field['comment'].splitlines(True):
			codePiece += '// \t' + line
		codePiece += '\n'
		return codePiece

	myNoteTxt = '''
// This file is a bcw name->address table.
// format: ba_name = ( func, addr, bcw_width, bit_start, bit_len )
// ba_name means the bcw_addr of name.
// addr is 8'h00~8'h0f or 8'hn0 with n being 1~f.
// bcw_width is 4 or 8. bit_len and bit_start is the attributes of the field.
'''
	global funcSelIsHere; funcSelIsHere = False
	genFile(outFile, genCode(myPattern), myNoteTxt)


def genReadBcwDefault(outFile='read_bcw_default.sv'):
	def myPattern(func, addr, field):
		if field['attrib'] in ['WO',]:
			return '\n// %s is not valid for read-default'%(field['name'])
		elif field['name'] == 'func_sel':
			return '\n// %s is ignored cause it is duplicated'%(field['name'])
		codePiece = '''
begin
int unsigned rd_val;
bcw_rd(ba_name, rd_val);
if(rd_val !== default_val)
	$display("%t BCW_SHMOO ERROR name default value err, exp 'h%0x actual 'h%0x",$realtime,default_val,rd_val);
end'''.replace('name', field['name']).replace('default_val',str(field['default']))
		return codePiece

	myNoteTxt = '''
// Read the default value of all bcws.
// ba_name means the bcw_addr of name.'''
	genFile(outFile, genCode(myPattern), myNoteTxt)


def genBcwWriteRead(outFile='bcw_write_read.sv'):
	def myPattern(func, addr, field):
		if field['attrib'] in ['RO',]:
			return '\n// %s is not valid for write-then-read'%(field['name'])
		elif field['name'] == 'func_sel':
			return '\n// %s is ignored cause it is duplicated'%(field['name'])
		codePiece = '''
begin
int unsigned wr_val;
wr_val = $urandom_range(max_val, 0);
write_read_then_cmp(ba_name, wr_val);
wr_val = max_val - wr_val;
write_read_then_cmp(ba_name, wr_val);
end'''.replace('name', '%s'%field['name'] ).replace('max_val', '\'h%x'%(2**field['bit_len']-1) )
		return codePiece

	myNoteTxtAndFunc = '''
// Write a random value to a reg, then read back and compare.
// Write the inverse value to the reg, then read back and compare again.
'''
#function void write_read_then_cmp(bit[0:4][7:0] ba_name, int wr_val);
	#int rd_val;
	#bcw_wr(ba_name, wr_val);
	#bcw_rd(ba_name, rd_val);
	#if(rd_val !== wr_val)
		#$display("%t BCW_SHMOO ERROR name write follow by read err, exp 'h%0x actual 'h%0x",$realtime,wr_val,rd_val);
#endfunction

	genFile(outFile, genCode(myPattern), myNoteTxtAndFunc)


def genBcurBcw(outFile='bcur_bcw_init.sv'):
	def declarePattern(func, addr, field):
		global funcSelIsHere
		if field['name'] == 'func_sel':
			if funcSelIsHere: return ''
			else: funcSelIsHere = True
		if field['is_context'] == 'yes':
			codePiece = ('bit[%d:0] bcw_name, bcur_name, bcur_name_c0, bcur_name_c1;'%(field['bit_len']-1)).replace('name', field['name'])
		else:
			codePiece = ('bit[%d:0] bcw_name, bcur_name;'%(field['bit_len']-1)).replace('name', field['name'])
		return codePiece

	global funcSelIsHere; funcSelIsHere = False
	declareCode = '// Declare vars of bcw_name, cur_name, etc.\n'\
			+ genCode(declarePattern)

	def setDefaultPattern(func, addr, field):
		if field['is_context'] == 'yes':
			codePiece = '\tbcur_name=dv; bcur_name_c0=dv; bcur_name_c1=dv; bcur_name_c0=dv; bcur_name_c1=dv;'.replace('name', field['name']).replace('dv', str(field['default']))
		else:
			codePiece = '\tbcur_name=dv;'.replace('name', field['name']).replace('dv', str(field['default']))
		return codePiece

	setDefaultCode = '''
// Set bcur_name to default value
function void set_bcur_default();
''' + genCode(setDefaultPattern) +\
'endfunction\n'

	def randPattern(func, addr, field):
		global funcSelIsHere
		if field['name'] == 'func_sel':
			if funcSelIsHere: return ''
			else: funcSelIsHere = True
		codePiece = '''
	`ifdef BC_NAME assert(std::randomize(bcw_name) with {`BC_NAME});
	`else assert(std::randomize(bcw_name)); `endif'''\
	.replace('NAME', field['name'].upper()).replace('name', field['name'])
		return codePiece

	def macroPattern(func, addr, field):
		global funcSelIsHere
		if field['name'] == 'func_sel':
			if funcSelIsHere: return ''
			else: funcSelIsHere = True
		codePiece = '\t`ifdef B_NAME bcw_name=`B_NAME; `endif'\
				.replace('NAME', field['name'].upper()).replace('name', field['name'])
		return codePiece

	funcSelIsHere = False
	randCode = '''
function void rand_bcw();
// Randomize bcw_name vars ''' + genCode(randPattern)
	funcSelIsHere = False
	macroCode = '\n// Define macros for command-line options\n' +\
			genCode(macroPattern) + '\nendfunction\n'

	code = declareCode + setDefaultCode + randCode + macroCode
	genFile(outFile, code, '')


def genBcwWw(outFile='bcw_ww.sv'):
	code = ''
	cnt = 0
	from bcw_db import bcwDb
	for func in bcwDb.keys():
		for addr in bcwDb[func].keys():
			for field in bcwDb[func][addr]:
				if func == 0:
					code += 'aa[index]=ba_name; dd[index]=bcw_name; cc+=1;\n'\
							.replace('index', str(cnt)).replace('name', field['name'])
				else:
					code += "// %s is not func0's bcw;\n"%(field['name'])
				cnt += 1

	genFile(outFile, code, '')


def run():
	#extractReg()
	genUpdateBcur()
	genBcwName2Addr()
	genReadBcwDefault()
	genBcwWriteRead()
	genBcurBcw()
	genBcwWw()

if __name__ == '__main__':
	try:
		run()
	except:
		import traceback; traceback.print_exc();
	finally:
		raw_input("\n---Press ENTER to quit---")

