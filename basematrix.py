from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import coordinate_to_tuple
from openpyxl.formula import Tokenizer
from decimal import *
import time


productList = []
writeList = []
target = "BASE "
path = "C:\\PaulScripts\\Pricing Matrices\\"

TOP_CUSTOMER = .9
REG_CUSTOMER = .85


def writeMatrix(fn, productList, margin, margins):
    """whut"""
    #print(productList[0])
    #print(writing)
    
    f = open(path + fn, 'w')
    i = 0
    j=0
    for line in productList:

            
            
        price = formatPrice(line[4], line[2], line[5], margins) 
        if price == 0:
            i+=1
            continue
            
        j+=1
        
        wline = ""

        wline = ""
        wline+=line[0]
        wline+=line[1]
        wline+=line[2]
        wline+=line[3]
        wline+=price
        wline+=line[5]
        wline+=line[6]
        wline+=line[7]
        wline+=line[8]
        wline+=line[9]
        wline+=line[10]
        wline+=line[11]
        #print(wline)
        wline+="\n"
        f.write(wline)      
        #quit()
    
    print(str(i) + " Items skipped")
    print(str(j) + " Items given prices")
    f.close()

def formatPrice(price, mfr, uom, margins):
    multiplier = 0
    margin = 0
    zeroes = "000000000000000"
    
    if price is None:
        return zeroes
    
    print(mfr)
    margin = get_margin(mfr, margins)
    
    
    if margin == 0:
        return 0
        
    
    price = zeroes[:-len(str(str(int(round((float(price) * get_multiplier(uom) ) / margin,2)*100))))] + str(str(int(round((float(price) * get_multiplier(uom) ) / margin,2)*100)))
    
    return price

def readLPF(fn):
    """whut"""
    
    
    while True:
        try:
        
            f = open(fn, 'r')

        except FileNotFoundError:
            print("CEDNet data files not found.  \nPlease run the Data Files export function in CEDNet")
            print("It can be found at Maintainance > Product > Data Files")
            choice=input("\nPress enter when ready, or type quit to return to the last screen").strip()
            
            if choice.lower() == 'quit':
                return
        break   
        
    pl = []
    i = 0
    for line in f:
        pl.append(parseLPFLine(line))
        i+=1

    f.close()
    return pl
    
def parseLPFLine(line):

        pl = []
        
        line=line.split("|")
        
        i = 0
        #for each in line:
        #   print(str(i) + ". " + each)
        #   i+=1
        if "PS" in line[1]:
            print(line[1:3])
    
        pl.append(target) #Target matrix
        pl.append("5")#Level
        pl.append(line[1] + "     "[:-len(line[1])])#Mfr
        pl.append(line[2] + "                      "[:-len(line[2])])#Name
        pl.append(line[10])#Price
        pl.append(line[12])#UOM
        pl.append("01012020")#Exp
        pl.append("   ")#Pad 3
        pl.append("00000000000"[:-len(line[3])]+line[3])    #UPC
        pl.append("      ") #pad 6
        pl.append("000000000000000")#Zeroes
        pl.append(line[1] + "     "[:-len(line[1])]) #matrix


        return pl

def compare(pl, cm):
    
    for cmeach in cm:
        popUs = []
        print(cmeach[1])
        for i in range(0, len(pl)):
            if cmeach[1] == pl[i][3]:
                popUs.append(i)
        
        
        for each in popUs:
            pl.pop(each)
    
        

    return pl

def get_margin(mfr, margins):
    
    for each in margins:
        if each[0] == mfr.strip():
        #   input(each)
            return float(each[1])
    
    for each in margins:
            if each[0] == "DEFAULT":
        #   input(each)
                return float(each[1])
    
        
    return .75

def get_multiplier(uom):
            
    if uom == "E":  #each
        return 1
    if uom == "C":  #per-hundred
        return 100
    if uom == "M": #per-thousand
        return 1000

def readMargins():

    try:
        f = open("C:\PaulScripts\Pricing Matrices\mfr_margins.txt")
        
    except FileNotFoundError:
        input("There is no file containing margins for things.  I've created a template file at \n "\
                "C:\PaulScripts\Pricing Matrices\mfr_margins.txt\n"\
                "It has examples in it to show you how to add your own rules.\n"\
                "For now I'm using a margin of 75% \nPress enter to continue...")
        
        f = open("C:\PaulScripts\Pricing Matrices\mfr_margins.txt",'w')
        
        
        f.write("#Change the #DEFAULT value to be whatever margin you'd like\n")
        f.write("#And add rules for any products you like\n#\n")
        f.write("#To exclude products, use a margin of 0\n")
        
        f.write("#MFR/CAT\tMargin\n")
        f.write("DEFAULT\t.75\n")
        f.write("FAKEM\t.75\n")
        f.write("FAKEM\t.75\n")
        f.write("FAKEC\t.9")
        f.write("SKIPM\t0")
        
        f.close
        return [['DEFAULT', .75]]
        
        
    lns = []
    
    for line in f:
        if line[0] == "#":
            continue
        
        if len(line.split()) != 2:
            continue
            
        lns.append(line.split())
        
    return lns
    
def run():
    #read all products in

    productList = readLPF("C:\Invsys\Algorithm\SPKPRDDT.lsq")       
            
    margins = readMargins()
    
    i=0 
        
    writeMatrix("Import Base.lsq", productList, REG_CUSTOMER, margins)
    
    print(str(len(productList)) + " Items in total.")
    
if __name__ == "__main__":

    run()
    
    print("Base matrix written to \"C:\\PaulScripts\\Pricing Matrices\\Import Base.lsq\"")
    
else:

    print("imported basematrix.py")