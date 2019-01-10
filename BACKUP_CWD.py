# Copy current working directory to a backup folder.
import shutil
import time
import os

def run(src='.', backupDir='D:\\Backup\\crater'):
	ct = time.ctime().split()
	suffix = '_' + ct[4] + ct[1] + ct[2] + '_' + ct[3].replace(':','-')
	dirpath = os.getcwd().replace(os.sep,'.').replace(':','')
	dst = backupDir + os.sep + dirpath + suffix
	print 'current working dir is copied to %s'%dst
	shutil.copytree(src, dst)

if __name__ == '__main__':
	try: 
		run()
	except: 
		import traceback; traceback.print_exc(); 
	finally:
		pass
		#raw_input("\n---Press ENTER to quit---")

