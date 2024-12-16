# DOCX to PDF Converter
#### Video Demo:  [YouTube](https://youtu.be/p6vc8tzocoU)
#### Description:
The choice for my final project app was dictated by my memories from the past, when I didn't even dreamt of being able to write code. In my early days, it was quite problematic to convert one file format into another one, especially **DOC** to **PDF**, since it was a functionality of paid versions of tools. (At least the tools I was familiar with.) So, I decided to try to make my own converter, even though now it's not an issue to find one.

## General description
My project represents a simple desktop application for converting DOCX files into PDF ones.
I chose Python and Tkinter (custom tkinter) for implementation for several reasons:

1. We simply used Python in the course, so I was familiar with the syntax and used the language.
2. It was the fastest way to implement such an app, since Python has a majority of open source libraries for convertion to PDF and manipulation of PDF files.
3. Tkinter turned out to be a pretty simple but functional tool to create a front-end part.

## Implementation
> NB: I used ChatGPT while planning and implementing the app. First of all, it helped me to list all possible technologies I can use to implement such an app. I used its suggestions and searched on the Internet the best tools and libraries for the app. Secondly, when I understood that it took too much time for me to find something on the Internet, ChatGPT helped to narrow the search field. Still, all the code I wrote was written by me.

My app uses **pypandoc** wrapper of **pandoc** converter. The other options I considered were **docx2pdf** and **unoserver**, but **docx2pdf** works only if Microsoft Office is intalled (which I don't have) and **unoserver** looked to complicated for me.

> ***MiKTeX installation is required for app usage!***

### Front-End part
As mentioned earlier, I used **Tkinter** and **Custom Tkinter** for GUI. Tkinter alone looked not so modern, so I found out that some parts of Custom Tkinter are more to my liking. So, Tkinter was still used for functionality and Custom Tkinter for design.

First of all, I create the main window (widget) with a title and dimensions. The window cannot be resized deliberately. The window is centered by finding out its coordinates on x and y axis (`x = (screen_width/2) - (width/2)` and `y = (screen_height/2) - (height/2)`) and then setting the geometry of the window `root.geometry('%dx%d+%d+%d' % (width, height, x, y))`.

There are 3 functional buttons:

- Open File (uses `command=file_select` for opening File Dialog and selecting 'DOCX' file)
- Choose the location for converted file (uses `command=select_location` for opening File Dialog and selecting save folder)
- and Convert file (uses `command=convert` for convertion initiation)

The main window, buttons and labels are from Custom Tkinter.

### Back-End part
The back-end part consists of 3 functions:

- `file_select()`
- `select_location()`
- `convert()`

**file_select()** includes 2 global variables with string type. They are used for storing the input file location on disk (`_input`) and the file name without extention (`filename`). Both variables are required for `convert()` function, that's why they are global. First, I let user select a file by opening file dialog in C:\Documents by default `filedialog.askopenfilename(initialdir="C:\Documents", title="Select a file", filetypes=(("docx files", "*.docx"),))`. User can select only 'DOCX' files. Path for file is stored in `global _input`. Second, I extract the name of the file without the extention from the path `filename = os.path.splitext(os.path.basename(_input))[0]`. And finally, place a label with the path to the file.

`def select_location()` promts user to select a location and stores it in `global location`. The label is placed as well for showing the save path for user.

`def convert()` performs checks and convertion itself. It has all the variables initialized by previous functions `global _input, location, filename`. Then it checks check whether file and output locations were specified. Othewise, user will receive a pop-up window with an error. After that it has to generate the path to converted file with .PDF extension. It's required since **pypandoc** function should receive both path to input and output files. Once the output file path was generated, function checks whether such output file already exists. If file exists, function promts the user to choose whether to overwrite the existing file or no. `pypandoc.convert_file` automatically creates new output file or overwrites the existing one. If no, the app waits for user to choose another location or rename the file or just close the app. If yes, we can proceed to the convertion itself. Function at first tries to convert a file. If successfull - "File converted successfully" is received. In case there was "RuntimeError", error message "Convertion failed" is received.

I didn't find information whether `pypandoc.convert_file` may though other errors exept for "RuntimeError". During implementation and testing I received only "RuntimeError".
