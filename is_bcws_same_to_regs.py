# Generate names of those regs which are better not touched during csr shmoo

from breg_db import bregDb
from bcw_db import bcwDb
def isBcwsSameToRegs():
	regNames = []
	for reg in bregDb:
		regNames.append(reg['name'])
	bcwFieldNames = []
	for func in bcwDb.keys():
		for addr in bcwDb[func].keys():
			for bcwField in bcwDb[func][addr]:
				bcwFieldNames.append(bcwField['name'])
	print "Below are bcws which are found not in design's reg-file:"
	for field in bcwFieldNames:
		if field not in regNames:
			print '\t' + field.upper()
	print "Below are regs which are found not in bcw_info.xls:"
	for reg in regNames:
		if reg not in bcwFieldNames:
			print '\t' + reg.upper()

def run():
	isBcwsSameToRegs()

if __name__ == '__main__':
	try:
		run()
	except:
		import traceback; traceback.print_exc();
	finally:
		raw_input("\n---Press ENTER to quit---")

