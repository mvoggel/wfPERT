## PERT Analysis and testing 
- Will require a local download of an excel spreadsheet, for which the program reads based on its file path: 
- You will need to update the file path in your python file to wherever you've saved your WF excel pull. 



How I run it inside VS Code editor terminal:  
/usr/local/bin/python3 /Users/myusername/Desktop/wfdata/main.py

Future Enhancements: 
- Recurring updates to data in real-time (security)
  - Possibly look at [this example](https://github.com/Workfront/workfront-api-examples-python/tree/master) for WF API calls?  
- Task-level analysis
- Post-analysis: print-out to new sheet within the excel once its ran
- Outlier handling (if over 50 days (PI length), default to 50 bus. days for calculation)
- Introduce logic handling for PI planning fields
  - % within expected PI, iteration
  - how often it moves
  - probability it WILL be complete in selected time frame (based on repeated history, criteria)
- Somewhere to run this other than a CLI in VS Code...
