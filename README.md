# Description
This is a test task for the Junior Python Developer position.

**Challenge**\
Your challenge is to automate the process of extracting data from itdashboard.gov:
- The bot should get a list of agencies and the amount of spending from the main page
- Click "DIVE IN" on the homepage to reveal the spend amounts for each agency
- Write the amounts to an excel file and call the sheet "Agencies".
- Then the bot should select one of the agencies, for example, National Science Foundation (this should be configured in a file or on a Robocloud)
- Going to the agency page scrape a table with all "Individual Investments" and write it to a new sheet in excel.
- If the "UII" column contains a link, open it and download PDF with Business Case (button "Download Business Case PDF")
- Your solution should be submitted and tested on Robocloud.
- Store downloaded files and Excel sheet to the root of the output folder
- This task should take no more than 4 hours.
- If you reach 4 hours with tasks still remaining, please describe how in theory you would complete this challenge if more time was allowed.
- Please leverage pure Python (ex: here) without Robot Framework using the rpaframework for this exercise. While API's and Web Requests are possible the focus is on RPA skillsets so please do not use API's or Web Requests for this exercise.
  
**Bonus task:**
- Extract data from PDF. You need to get the data from Section A in each PDF. Then compare the value "Name of this Investment" with the column "Investment Title", and the value "Unique Investment Identifier (UII)" with the column "UII"
