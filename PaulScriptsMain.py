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
import shutil
import os
import requests
from zipfile import ZipFile



def PaulScriptsMenu():


	update_scripts()

	try:
		
		print("Welcome to Paul's CEDNet Utility Scripts main menu")

		while True:
			os.system('cls')
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
			
def update_scripts():
	
	

	#Check if there was an update.  
	try: 
		r=requests.get("https://api.github.com/repos/paulizleet/CEDNet-Python-Suite/git/refs/heads/master")
		if r.status_code != requests.codes.ok:
			print("Error fetching the update.  Are you online?  Skipping update this time.")
			return
		
		print("Got OK request")



		newreq = None
		sp = r.text.split(",")
		

		try:
			f = open("C:\\PaulScripts\\configs\\last_update.txt", 'r')
			txt=f.read()
			f.close()
		except:
			pass
		f = open("C:\\PaulScripts\\configs\\last_update.txt", 'w')
		
		
		counter=0
		for each in sp:
			
			if each[:5].strip() == "\"url\"":
				if counter == 0:
					counter+=1
					continue
				#input(each[6:].replace("\"", ""))
				newreq=requests.get(each[6:].replace("\"", "").replace("}", ""))
		
		print("got second request")


		sp=newreq.text.split(",")
		
		for each in sp:
			
			if each[:6].strip() == "\"date\"":
				if each.strip() == txt.strip():
					print("No update required")
					return
				else:
					f.write(each)
					f.close()
					break
	except:
		print("Error fetching the update.  Are you online?  Skipping update this time.")
		return
	
	
	#There was a mismatch between the latest commit and the current code.
	#Download the latest code and update.
	r = requests.get("https://github.com/paulizleet/CEDNet-Python-Suite/archive/master.zip")
	f=open(os.getcwd()+"\\update.zip", "wb")
	f.write(r.content)
	f.close()
	
	z=ZipFile(os.getcwd()+"\\update.zip", "r")
	z.extractall(".\\update")
	z.close()
	for roots, dirs, files in os.walk(os.getcwd()+"\\update\\CEDNet-Python-Suite-master"):
		print(roots)
		print(dirs)
		for each in files:
			shutil.copy(roots+"\\"+each, os.getcwd()+"\\Pythons\\CEDNet-Python-Suite\\")
		break
		
	shutil.rmtree(os.getcwd()+"\\update")
	os.remove(os.getcwd()+"\\update.zip")
	
	print("Update completed.  This script will now quit.")
	quit()
	
	

if __name__ == "__main__":
	PaulScriptsMenu()



