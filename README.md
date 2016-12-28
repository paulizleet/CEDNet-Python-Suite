# CEDNet-Python-Suite
Packaged version of my CEDNet utilities.  The scripts will automatically update by cloning this repository.
A Loose group of utilities designed to make working with my company's proprietary software a little less of a headache.

It's kind of a mess of commented code and is probably a pain to look at, but it works well for my purposes.

Dependencies: [OpenPyxl](https://pypi.python.org/pypi/openpyxl)

Here is an outline of what everything is used for and why I created it in the first place.

## CEDNetUtils.py
This is a sort of header file that includes a bunch of things that are common to the rest of the scripts, such as parsing
the flat files our software prints out, and writing lots of data to spreadsheets.

##nobins.py
All of the products are assigned a place on our shelves called a Bin Location.  This script parses our software's
product file and prints out a spreadsheet detailing every item on our shelves, its bin location, and creates a different 
list for every item which doesn't have a bin location.  

##ShelfAddresses.py
Before I started working here, none of our products had a bin location on file.  
I wasn't going to manually enter 3000 products into our system, so this script handles that using keyboard and mouse inputs.
Bin locations are manually typed into the spreadsheet that nobins.py produced, and the script handles the rest.

##matrix.py
Every week I had to update pricing for wire and pipe.  This script assumes we bought our material from a certain vendor who always uses
the same spreadsheet to give us our prices.  It produces a pricing matrix that can be quickly imported into CEDNet.

##panels.py
This is a simple script that reads our product file, and writes our current inventories of solar panels to a spreadsheet.

##speaks.py
An agonizing script to maintain. Instead of just outputting data to a text file, it creates a pdf, and converts that into txt instead. 
This script takes a pdf disguised as a a text file and outputs only the data I need.  It's so difficult to maintain because sometimes 
SPEAKS outputs data in a different way than the last time. Therefore I'm not even sure if this works anymore.

##solar.py
I was responsible for reporting on our sales of certain solar equipment manufacturers.  This takes a SPEAKS printout of the previous 
week's sales, and produces spreadsheets detailing who bought what and where it went.  

#cycle.py
Maintains a spreadsheet of products and their inventory levels, and takes 40 or so of them every day for me to go and verify.
It keeps track of which products I've already checked so it won't make me check them again.  After they've all been checked, it wipes everything and starts fresh.
