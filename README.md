claude:

* installation:  $  uv add streamlit gspread google-auth  openpyxl
* to run:  $  uv run streamlit run  ./src/streamlit_tanker/app.py



also, make the background colour of the "body" as light yet bright lemon yellow


a) can you make the option for choosing the block name, currently depicted as radio buttons, 
as normal buttons, with background colour as dark teal?
b) project is created using uv tool/ So, the project structure is: 

Project_root
�   .gitignore
�   .python-version
�   pyproject.toml
�   README.md
�   uv.lock
�   
+---.streamlit
�       secrets.toml
�  
+---data
�       tanker.xlsx
+---src/streamlit_tanker
�                      app.py   
+---.venv
�       

So, note the location of the "data/tanker.xlsx" file. adjust the code accordingly.
