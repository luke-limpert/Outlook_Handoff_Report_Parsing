# Outlook_Handoff_Report_Parsing
 Comb through outlook and aggregate handoff reports
 
 This program seeks to take reports from outlook emails
 and compile all of the data into one larger data set.
 Due to the reporting systems in the organization these reports 
 are the only accurate representation of completed components 
 sent out by production.
 
 Many of the problems I have had in this project so far come 
 from the way they choose to place the report into the email.
 
Some copy and paste from their excel spreadsheets and the 
testfile.txt shows an example of that output when it is saved
as a text file. One way of dealing with these emails is by 
substitution of the new line with "". This produces the 
"substituted.txt". It looks something like the example below.

RP1 GA600415151312    28Training New Hires/Staffing    RP2 GA600520151212
128Training New Hire/Staffing No RSA 35 min.RP MACHINING
RP1 SPN59554075122   18Op 60 shutter door fell off. 20 min. 
Lazer etch not reading 15 min. RP1 RBM595475120-1 Thermal Overload 
fault wouldn’t clear 40 min. Op 160 LVDT broke 35 min.RP1 RBG595537
No RBs 60 min.RP1 BN595520Usach Part lost in cycle 20 min. 
RP2 SPN5957457712     8Great Job!!!!RP2 RBM595465Op 130 code read 
issues. Tune camera 45 min. RP2 RBG595403No RBs 145 min.RP2 
BN595625Toyoda grinder has been flooding due to the skimmer 
getting clogged for a couple of weeks daily.              
Comments 

From there i can use regex to get the information out of it 
pertaining to the production target vs actual. 

(?P<Line>[A-Za-z0-9]*)  
(?:\s)  
(?P<Stage>[A-Za-z]*)  
(?P<Target>\d{3})  
(?P<Actual>\d{1,3})  

Tried to make it a bit VERBOSE for readability. 

The other challenge is when there is no copy pasted report in the 
email and instead they choose to include it as an excel spreadsheet.

This is the part i am currently working on, but i think i should just
be able to resave those documents as .csv files. 
