from tkinter import filedialog
import os
# User defined
from verticalCheck import main as verticalCheck
from roomCheck import main as roomCheck
from generateTeamDocs import main as generateTeamDocs

if __name__ == "__main__":
    file = filedialog.askopenfilename(title = "Select File",filetypes = (("xlsx files","*.xlsx"),("xls files","*.xls"),("all files","*.*")))
    
    if verticalCheck(file=file) and roomCheck(file=file):
        generateTeamDocs(file=file)