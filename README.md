# gsbtb_tool_parse-emails
Parsing existing GSBTB's GMail inbox into CSV for import into gsbtb_tool

First, download example data: see Gmail inbox, a message labeled "Your Google data archive is ready", get the ZIP file, extract into this folder.

Then, execute the the script:
   python run.py

You should find four CSV files being generated.
Before the next run, detele those files, or the script will append the new data to those files.
   rm *.csv && python run.py
