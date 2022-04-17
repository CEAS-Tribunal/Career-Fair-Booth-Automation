# Career-Fair-Booth-Automation

**Prerequisites:**

    1. Python 3 installed on your machine
    
    2. A Registered Companies Excel sheet (see below for specifications)
    
    3. A Layout Excel sheet (see below for specifications)

**Registered Companies Excel Sheet:**

    1. The exact names of relevant columns should be the following:
          - "Employer"
          - "Sessions"
          - "Employer Industry"
          - "Requested Booth Options"
          - "Combined Majors"
          - "General Items - Access to Electric"
          - "Big Company" (optional, must be added manually)
          - If any of these column names no longer exist when the Excel file is populated from Handshake, 
             the name of the columns must be changed in the code to reflect the new column names. This
             change should be made in companies.py, lines 53-59.
             
    2. "Sessions" column
        - Each session name must be separated by a semi-colon (;)
        - The sessions listed must follow a consistent order throughout the sheet
              - For example, maybe Tech Day One always comes before Tech Day Two, both of which always come before Professional Day
              
    3. "Requested Booth Options" column
        - Possible values are "Standard Booth" or "Premium Booth"
        - Values must be separated by a comma (,)
        - Values must be provided in the order corresponding to the session
              - For example, if a company's "Sessions" are: "Tech Day One; Tech Day Two"
              - and the company's "Requested Booth Options" are: "Standard Booth, Premium Booth",
              - it is assumed that they want a Standard Booth for Tech Day One, and Premium Booth for Tech Day Two.
              
    4. "General Items - Access to Electric" column
        - Possible values are 0 or 1
        
    5. "Big Company" column (optional)
        - This is an optional column which must be created and filled manually by the Career Dev Team.
        - This column can be placed anywhere in the sheet, but must be named "Big Company"
        - Possible values are "1" or (blank) - the Career Dev Team should place a "1" next to any companies they deem to be "Big"
          
 **Layout Excel Sheet:**
 
    1. All booths must have a name with the first character being a Capital Letter, and the second character being a number (e.g. A1, B22, C3*)
        - All booth names within a layout must be unique
        
    2. All premium booths must have a name ending in an asterisk (*) (e.g. C3*)
        - All premium booths should have a blank cell reserved next to them to indicate that premium booths take up two spaces in the layout
        - All premium booths should be highlighted with the default Excel yellow (FFFFFF00) to indicate that this booth should have access to electricity
        
    3. All booths with access to electricity should be highlighted with the default Excel yellow (FFFFFF00)
    
    4. Any cells with values in the layout should contain lowercase letters only (e.g. "check-in", "exit", etc.)

**Instructions to Run the Automation:**

    1. Download Career_Fair_Booth_Automation.bat, main.py, layout.py, and companies.py. 
    
    2. Place these files in the same directory as the existing Registered Companies and Layout Excel Sheets. 
    
    3. Edit the constants in main.py, lines 7-9.
        - SESSIONNAME = the exact name of the session the layout is being generated for as displayed in the Registered Companies Excel Sheet
            - The automation must be run separately for each session, so whichever session it's currently being run for should be enetered here.
        - LAYOUTFILENAME = the exact name of the Layout Excel File, which is assumed to be stored in the same folder as the code is being run from
        - COMPANIESFILENAME = the exact name of the Registered Companies Excel File, which is assumed to be stored in the same folder as the code is being run from
        
    3. (Optional) Create a column anywhere in the Registered Companies Excel Sheet named "Big Company", and mark a "1" next to all companies deemed to be "Big".
    
    4. Run Career_Fair_Booth_Automation.bat
    
    5. Enter any industries to be excluded from the current automation.
        - This feature serves as a way to exclude certain industries so they can be placed in a separate location, if multiple locations are available for the session.
        - Enter as many industries as you would like to exclude. When finished, Enter nothing to continue.
        
    6. If Step 3 was completed, enter a "1" when the automation asks "Do you want to sort by big companies?". If Step 3 was skipped, enter a "0".
    
    7. If there are not enough booths in the layout, warnings will be shown in the window. Be sure to check if any warnings are shown before closing the window. 
    
    8. The completed layout will be placed in a "results" folder. The Career Dev Team should review the layout and make any manual changes they see fit. 
        
