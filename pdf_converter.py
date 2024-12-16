import customtkinter
import pypandoc
import os
from tkinter import *
from tkinter import filedialog
from CTkMessagebox import CTkMessagebox


#--BACKEND--

#function for selecting file
def file_select():
  global _input
  global filename
  global labelSelect
  _input = filedialog.askopenfilename(initialdir="C:\Documents", title="Select a file", filetypes=(("docx files", "*.docx"),)) #store path to file as str
  # extract file base name and save only name without extention from path. 'doc.docx' will be split into ('doc', 'docx'). list, thus should select the first str [0]
  filename = os.path.splitext(os.path.basename(_input))[0]
  # update the label if already exists
  if 'labelSelect' in globals() and labelSelect.winfo_exists():
      labelSelect.configure(text="Input: "+_input)
  else:
    labelSelect = customtkinter.CTkLabel(root, text="Input: "+_input, fg_color="transparent", text_color="white", wraplength=390)
    labelSelect.place(relx=0.5, rely=0.27, anchor=CENTER)
  

# location selection
def select_location():
  global location
  global labelLocation
  location = filedialog.askdirectory(initialdir="C:\Documents", title="Select Folder")
  if 'labelLocation' in globals() and labelLocation.winfo_exists():
    labelLocation.configure(text="Output location: "+location)
  else:
    labelLocation = customtkinter.CTkLabel(root, text="Output location: "+location, fg_color="transparent", text_color="white", wraplength=390)
    labelLocation.place(relx=0.5, rely=0.55, anchor=CENTER)


# function for file convertion
def convert():
  global _input, location, filename
  # check whether file and output locations were specified 
  try:
    _input, location
  except NameError:
    CTkMessagebox(title="Error", message="Provide input file and location for output file", icon="cancel")
  # creating output file location and file name
  output = rf"{location}/{filename}.pdf" # rf (raw format) used so that \ isn't considered as escape char
  # check whether such file exists
  if os.path.isfile(output):
    msg = CTkMessagebox(title="Overrite file", message="File already exists. Do you want to overrite it?",
                        icon="warning", option_1="No", option_2="Yes", button_width=10)
    if msg.get() == "No" or msg.get() == None:
      return
  # try to convert
  try:
    pypandoc.convert_file(_input, 'pdf', outputfile=output)
    CTkMessagebox(title= "Success", message="File converted successfully", icon="check", option_1="Thanks")
  except RuntimeError:
    CTkMessagebox(title="Error", message="Convertion failed", icon="cancel")


# --INTERFACE-- 

# root widget creation
root = customtkinter.CTk()
root.title("PDF Converter")
root.resizable(False, False)


# center the root widget
width = 400
height = 500
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = (screen_width/2) - (width/2)
y = (screen_height/2) - (height/2)
root.geometry('%dx%d+%d+%d' % (width, height, x, y))


# buttons creation
open_btn = customtkinter.CTkButton(root, text="Open DOCX file", height=50, width=150, command=file_select)
location_btn = customtkinter.CTkButton(root, text="Output file location", height=50, width=150, command=select_location)
convert_btn = customtkinter.CTkButton(root, text="Convert to PDF", height=50, width=150, fg_color="green", command=convert)


# footer by label
copyright = u"\u00A9"
footer = customtkinter.CTkLabel(root, text=copyright +" Developed by Shyma")


# place buttons and footer
open_btn.place(relx=0.5, rely=0.15, anchor=CENTER)
location_btn.place(relx=0.5, rely=0.43, anchor=CENTER)
convert_btn.place(relx=0.5, rely=0.7, anchor=CENTER)
footer.place(relx=0.5, rely=0.9, anchor=CENTER)


root.mainloop()