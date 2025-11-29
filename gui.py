import tkinter as tk
import tkinter.font as tkFont
from tkinter import Button, filedialog, messagebox
from find import process_find, open_sesame
from search import look_up

# initializing the canvas
root = tk.Tk()

# setting the Rutgers Logo
rutgersR = tk.PhotoImage(file="R.png")
root.iconphoto(False, rutgersR)

# informational text
text = 'This program offers two different functions. \n\nThe Find Function will take one excel spreadsheet with two sheets. ' \
'The first sheet will contain the ISBNs you would like to check and the second sheet will contain the ISBNs you would like to be checked. ' \
'The function will highlight all of the ISBNs in the first sheet that are also present in the second sheet. \n\nBE CAREFUL!\n\n' \
'The program will not run correctly if you do not follow the correct formatting. The first sheet should only contain ISBNs in the column A. ' \
'Any other ISBNs will not be checked. The second sheet supports four columns of ISBNs, but no more. \n' \
'\nThe Search Function will take one excel spreadsheet with one sheet. This sheet will contain the ISBNs you would like to check across different platforms.' \
'Before doing so, the program will ask you for a website name. The only valid options include "amazon", "barnes", and "rutgers". ' \
'Anything else will not work (capitalization doesn"t matter). The program will then search all of the ISBNs on the preferred ' \
'website and write either "Yes" or "No" depending on whether the chosen website is listed or not. \n\nBE CAREFUL!\n\nThe program will only read ' \
'ISBNs in Column A and will overwrite anything in Column B with either the "Yes" or "No". \n\n' \
'IMPORTANT!\n\n Make sure the excel sheet that you are selecting is CLOSED! The program will not function if the Excel sheet is open.'


# Window Setup
root.title("Rutgers Press File Explorer")
root.geometry("1200x900+300+50")
root.configure(bg="#fffdd0")

# Welcome line
font = tkFont.Font(family="Terminal", size = 18, weight="bold")
welcome = tk.Label(root, text="Welcome!", font=font, bg="#fffdd0")
tk.Label(root, text="", font = font, bg="#fffdd0").pack(padx=0, pady=10)
welcome.pack(padx=10, pady=20)
tk.Label(root, text="", font = font, bg="#fffdd0").pack(padx=10, pady=0)


# Canvas for informational text
canvas = tk.Canvas(root, width=800, height=450, bg="#fffdd0")
canvas.create_text(400, 200, text=text, width=700, anchor="center", font=("System", 14))
canvas.pack()
tk.Label(root, text="", font = font, bg="#fffdd0").pack(padx=50, pady=20)



#allowing input from the textbox to know which website to search
website_name = tk.StringVar()

# find function to search for specific values in two different sheets (works on button click)
def open_find():
    file_path = filedialog.askopenfilename(initialdir ="/Downloads", title="Select a File", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
    while not(file_path.lower().endswith(".xlsx")):
        root.filename = filedialog.askopenfilename(initialdir="/Downloads", title="Select a File", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
        file_path = root.filename
    process_find(file_path)
    open_sesame(file_path)

# checks if the website entered is correct
def check():
    website = website_name.get()
    web_searches = ["rutgers", "barnes", "amazon"]
    if (website.lower().strip() not in web_searches):
        messagebox.showwarning("Try Again.", f"{website} is not supported. Try again.")
    else:
        messagebox.showinfo("Success!", f"Starting the search process on {website}!")
        gather()

# gathers website name to scrape and the excel sheet to use
def gather():
    website = website_name.get()
    file_path = filedialog.askopenfilename(initialdir="/downloads", title="Select a File", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))

    while not(file_path.lower().endswith(".xlsx")):
        root.filename = filedialog.askopenfilename(initialdir="/downloads", title="Select a File", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
        file_path = root.filename

    look_up(file_path, website)
    open_sesame(file_path)
    
# search function that initializes and gathers input (on button click)
def open_search():
    entryFont = tkFont.Font(family="System", size = 10, weight="bold")
    direction = tk.Label(root, text="Please enter a website to search!", font=entryFont, bg="#fffdd0")
    direction.pack(padx=0, pady=10)
    entry = tk.Entry(root, textvariable=website_name, font=('calibre', 10, 'normal'))
    entry.pack(pady=0)
    entry.focus_set()
    submit = Button(root, text="Submit", command=check)
    submit.pack()


# all the buttons being used
find_button = Button(root, text="Find Function", command=open_find).pack()
search_button = Button(root, text="Search Function", command=open_search).pack()

root.mainloop()