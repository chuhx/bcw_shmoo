import os
def run(bigFile='db_auto.sv'):
	print '\nGenerating %s'%bigFile
	print '-'*40
	all = open(bigFile,'w')
	splitLine = '// --- split line ---\n'
	for fname in os.listdir('.'):
		# if fname.endswith('.sv') :
		if fname in ['bcur_bcw_init.sv', 'bcw_name2addr.sv', 'bcw_write_read.sv', 'bcw_ww.sv', 'breg_name2addr.sv', 'read_bcw_default.sv', 'update_bcur.sv', ]:
			if fname == bigFile: continue
			print 'Copying %s'%fname
			with open(fname,'r') as f:
				all.write(f.read())
			all.write('\n// Above is %s\n'%fname)
			all.write(splitLine)
	all.close()

if __name__ == '__main__':
	try: 
		run()
	except: 
		import traceback; traceback.print_exc(); 
	#finally:
		#raw_input("\n---Press ENTER to quit---")


