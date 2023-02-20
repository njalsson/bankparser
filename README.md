# Guild bank parser that formats data into a excel sheet

I was too lazy to manually update the data in the bank, i did it once and it was horrible. so i wrote a program to do it for me 


using the guild bank addon that i've modified, trckster and solski already have it so i'm not going to bother linking it.





## naming the sheet correctly

change the player variable near the top of main.py or refactor the code how you see fit to have the sheet named correctly.


## requirements

to install all dependancies run pip(3) install -r requirements.txt


## how to run it?

i assume you have a python3 interpreter. after you have installed the dependancies and have the data from the bank addon
put the raw text from that addon into a 'bank.txt' file in this folder. then run the main.py and it'll compile a excell sheet called 'todaysBank.xlsx'


##  How to import the todaysBank.xlsx into google sheets?

click file in the top menu -> import -> upload file -> insert new sheet(s) -> click import data
