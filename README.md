# attendance-spreadsheet
**Synopsis**
Built originally for World Oceans Day, this lightweight Python script helps users upload event attendees into an Excel spreadsheet.

**How to Use**
1. Write the attendees (one on each line) into orgs.txt in plain text format.
2. Click RUN PROGRAM.
3. Answer the question "What year is it?" in Terminal when prompted.
4. Sit back and relax while attendance-spreadsheet fills the Excel template. It's that simple.

**The Problem** <br/>
Over the past few years, many organizations have registered to hold events recognizing World Oceans Day (WOD). WOD volunteers would manually update an Excel spreadsheet to keep track of these organizations; however, as WOD grew, they needed a tool that would automatically add organizations to the spreadsheet. This tool would have to recognize organization names that had been misspelled, rearranged, and/or incorrectly capitalized as previous attendees. 
 
**The Solution** <br/>
I built a Python script that populates an Excel spreadsheet when WOD inputs attendance lists for each year. I read and wrote from this spreadsheet using the openpyxl library. To track duplicates between old organizations and new entries, I made fuzzy string comparisons using fuzzywuzzy. This solution can easily be adapted to track any attendance over a period of time by simplying entering attendee names. 

**Libraries Used**
- openpyxl
- fuzzywuzzy
- titlecase
