# -*- coding: utf-8 -*-
"""
Created on Tue Sep 13 11:06:17 2016

@author: pgallagherjr

PaulScripts main menu program


"""

import basematrix
import matrix
import solar
import ShelfAddresses
import cycle
import nobins

import gitpython

def PaulScriptsMenu():
	update_from_git()

	try:
		
		print("Welcome to Paul's CEDNet Utility Scripts main menu")

		while True:
			print("\n")
			print("1. Inventory Level Checking")
			print("2. Solar Reporting")
			print("3. Matrix operations")
			print("4. Bin Locations")
			print("5. Quit")
			choice = ""

			while True:
				choice = input("Please make a selection.")
				if choice in ["1", "2", "3", "4", "5"]:
					break
				else:
					print("Invalid choice")

			if choice == "1":
				print("Running today's inventory level checking script.")
				cycle.run()

			if choice == "2":
				print("Solar Reporting")
				solar.run()

			if choice == "3":
				print("Matrix Operations")
				chose_matrix()

			if choice == "4":
				print("Bin Locations")
				chose_bins()
				
			if choice == "5":
				return
				
	except:
		print("A problem has occurred that I haven't seen before.\n  I've left a log of the error in C:\\PaulScripts\\sys\\logs")
		print("Just kidding i haven't gotten to this part yet")

def chose_matrix():
        print("\n")
        print("1. Wire/Pipe Matrix")
        print("2. Base Matrix")
        print("3. Go Back")
        
        while True:
            choice = input("Please make a selection.")
            if choice in ["1", "2", '3']:
                break
            else:
                print("Invalid choice")
                
        if choice == "1":
            print("Wire Matrix...")
            matrix.run()
        if choice == "2":
            print("Base Matrix...")
            basematrix.run()
        if choice == "3":
            return
def chose_bins():
        print("\n")
        print("1. Get items without bins")
        print("2. Enter Bin locations")
        print("3. Go Back")
        
        while True:
            choice = input("Please make a selection.")
            if choice in ["1", "2", '3']:
                break
            else:
                print("Invalid choice")
                
        if choice == "1":
            print("Finding items without bin locations...")
            nobins.run()
        if choice == "2":
            print("Preparing to enter bin locations...")
            ShelfAddresses.run()
        if choice == "3":
            return
			
def update_from_git():
	pass

if __name__ == "__main__":
    PaulScriptsMenu()



