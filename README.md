claude:

* installation:  $  uv add streamlit gspread google-auth  openpyxl
* to run:  $  uv run streamlit run  ./src/streamlit_tanker/app.py

* TESTING flag: comes from secrets.toml (bool or string). 
                Missing -> True (local Excel, the default.)

    So:
        TESTING = true      # use local Excel file (data/tanker.xlsx)
        TESTING = false   # use Google Sheet


background: radial-gradient(
    circle at left center,
    #d01a8b 0%,
    #e64caf 20%,
    #eb8fd0 40%,
    #e9c0e3 70%,
    #f0dff0 100%
);


-------- APp info -----------------------------------------------
a) dashboard:  dashboard-wtanker.netlify.app
b) tanker data entry:  