
# run python files
import os, sys

# download required python library with pip
os.system('pip install gspread oauth2client slacker slackclient')

# run gsheetproj.py
os.system('py updateGSheet.py ' + sys.argv[1])