#findexceptions.py

import sys
import os

keywords=["save(", "write(", "open(", "load_workbook(",

def main():

	asdf="1.2,3'4;5"
	
	print(asdf.split('.',',','\'',';'))
	
	quit()



	path = "C:\\Users\\pgallagherjr\\Desktop\\finalpython\\Pythons\\Definitely Useful\\"
	f = open("C:\\Users\\pgallagherjr\\Desktop\\exceptionlog.txt", 'w')
	for each in os.listdir("C:\\Users\\pgallagherjr\\Desktop\\finalpython\\Pythons\\Definitely Useful"):
		log = get_exceptions(path + each)
		
		for line in log:
			f.write(line)
			
	
	f.close()
	
	
def get_exceptions(filepath):
	try:
		f = open(filepath, 'r')
		
		for i, line in enumerate(f):
			for each in keywords:
			
			
			
			
		
	except PermissionError:
		pass
		
	return ""

main()
		
