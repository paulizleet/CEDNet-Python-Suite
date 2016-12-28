# -*- coding: utf-8 -*-
"""
Created on Tue Sep 13 10:31:55 2016

@author: Paul Gallagher

Setup script for my CEDNet Utilities
"""

import os
import pip
from shutil import copytree
import logging


    
from openpyxl import Workbook

import cycle
import nobins


path = "C:\\"

def make_dir(dir):
    print(path + dir)
    try:
        os.mkdir(path + dir)
		print("Created directory {d}".format(d=path+dir)
    except FileExistsError:
        print("Directory {d} already exists.".format(d=dir))

        
def _logpath(path, names):
    logging.info('Working in %s' % path)
    return []   # nothing will be ignored       



if __name__ == "__main__":
    
    if os.getcwd()[0] != 'C':
        print("I noticed you are running this from a flash drive. \n\n"\
        "This program is fully portable, but will run very slowly on a flash drive."\
        "\nConsider installing this utility to your hard disk to make it faster.")
        while True:
            inp=input("Would you like this program to copy itself to your hard drive? Y/N\n")
            if inp.upper() == 'Y':
            
                print("This is going to take a while.  Just don't close this program before it's finished.")
                copytree(os.getcwd()[:3], "C:\\Python", ignore=_logpath)
                print("\nFinished Copying files.")
                break
            elif inp.upper()=='N':
                print("Be prepared to wait a while.\n\n")
                break
                
                
                
    #Installing the required libraries
    try:
        pip.main(['install', 'openpyxl'])
    except:
        print("pip failed to install openpyxl.  Halting.")
        quit() 
    try:
        pip.main(['install', 'pypiwin32'])
    except:
        print("pip failed to install pywin32.  Halting.")
        quit()  
	try:
		pip.main(['install', 'requests'])
	except:
		print("pip failed to install pywin32. Halting.")
		quit()
    
    #Making all of the needed directories
    make_dir("PaulScripts")
    make_dir("PaulScripts\\Inventory Checking")
    make_dir("PaulScripts\\Pricing Matrices")
    make_dir("PaulScripts\\Shelving Addresses")
    make_dir("PaulScripts\\Solar Reporting")
    make_dir("PaulScripts\\Solar Reporting\\Weekly")
    make_dir("PaulScripts\\Solar Reporting\\Quarterly")
    make_dir("PaulScripts\\Solar Reporting\\Monthly")
    make_dir("PaulScripts\\Solar Reporting\\Monthly\\SMA")
    make_dir("PaulScripts\\Solar Reporting\\Monthly\\LG")
    make_dir("PaulScripts\\Speaks Exports")
    make_dir("PaulScripts\\Wire Matrix")
    make_dir("PaulScripts\\Wire Matrix\\wirebooks")
	make_dir("PaulScripts\\sys\\configs")
	make_dir("PaulScripts\\sys\\logs")
	
	if os.path.isfile("C:\\PaulScripts\\configs\\customer_info.txt") == False:
	
		f=open("C:\\PaulScripts\\configs\\customer_info.txt",'w')
		
		f.write("2		Account Number\n")
		f.write("3		Customer Name\n")
		f.write("5		Address Line 1\n")
		f.write("6		Address Line 2\n")
		f.write("9		Customer city\n")
		f.write("8		Customer zip\n")
		f.write("7		Customer state\n")
		f.close()
	else:
		print("customer_info.txt already exists.  Will not overwrite.")
	if os.path.isfile("C:\\PaulScripts\\configs\\solar_product_info.txt") == False:	
		f=open("C:\\PaulScripts\\configs\\solar_product_info.txt",'w')
		
		f.write("-5		Order Date\n")
		f.write("2		Cat Number\n")
		f.write("-3		Order Quantity\n")
		f.write("4		Invoice Number\n")
		f.write("
		f.close()
	else:
		print("solar_product_info.txt already exists.  Will not overwrite.")
    wb = Workbook()
    
    
    #Creating required worksheets and initialize them if they don't already exist.
	
	if os.path.isfile(path + "PaulScripts\\Inventory Checking\\cycle.xlsx") == False:
		wb.save(path + "PaulScripts\\Inventory Checking\\cycle.xlsx")
		cycle.run(skip_prompt=True)
		print("cycle ran for the first time")
	else:
		print("cycle.xlsx already exists.  Will not overwrite.")	
		

	if os.path.isfile(path + "PaulScripts\\This Week's Stock Status.xlsx") == False:
		wb.save(path + "PaulScripts\\This Week's Stock Status.xlsx")
		nobins.run(skip_prompt=True)
		print("Nobins ran for the first time.")
	else:
		print("This Week's Stock Status.xlsx already exists.  Will not overwrite.")
		
		


    
    f = open("C:\\PaulScripts\\.init.txt", 'w')
    f.write("Scripts setup OK")
    f.close()