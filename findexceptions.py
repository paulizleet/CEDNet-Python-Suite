#findexceptions.py

import sys
import os

keywords=["save(", "write(", "open(", "load_workbook(", ".__len__()"]

def main():



	path = "C:\\Users\\pgallagherjr\\Desktop\\finalpython\\Pythons\\Definitely Useful\\"
	f = open("C:\\Users\\pgallagherjr\\Desktop\\exceptionlog.txt", 'w')
	for each in os.listdir("C:\\Users\\pgallagherjr\\Desktop\\finalpython\\Pythons\\Definitely Useful"):
		log = get_exceptions(path + each, each)
		
		for line in log:
			f.write(line)
			
	
	f.close()
	
	
def get_exceptions(filepath, filename):
	returns = ""
	nspaces = "                    "
	lspaces = "    "
	try:
		f = open(filepath, 'r')
		
		for i, line in enumerate(f):
			for each in keywords:
				if each in line:
					returns+="{f} {l} {t}\n\n".format(f=filename+nspaces[:-len(filename)],
					l=str(i) + lspaces[:-len(str(i))],
					t=line.strip())
					
					
	
			
		
	except PermissionError:
		pass
		
	return returns

main()
		
