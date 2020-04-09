import os
import sys
import tkinter as tk
import win32com.client
from shutil import copyfile
from tkinter import filedialog


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# initialize initial and converted file path
initial_file = None
converted_file = None

# window
master = tk.Tk()
master.title("WORD/PDF Converter")
master.wm_iconbitmap(resource_path('arrow.ico'))
master.geometry("450x400")

# frames
top_frame = tk.Frame(master)
top_frame.pack(side='top')
center_frame = tk.Frame(master)
center_frame.pack(side='top')
bottom_frame = tk.Frame(master)
bottom_frame.pack(side='top')


def open_file(open_message, open_label, convert_message, save_message):
    global initial_file
    global converted_file
    initial_file = None
    converted_file = None

    file_types = [("WORD", "*.docx"), ("PDF", "*.pdf"), ('All files', '*')]
    accepted_extensions = {'pdf', 'docx'}
    file = filedialog.askopenfilename(parent=top_frame, initialdir='/', title="Select file", filetypes=file_types)
    extension = file.split('/')[-1].split('.')[-1]

    if not file:
        open_label.config(fg='red')
        open_message.set("No file selected!")
    elif not os.path.isfile(file):
        open_label.config(fg='red')
        open_message.set("The selected file has been moved or cancelled!")
    elif extension not in accepted_extensions:
        open_label.config(fg='red')
        open_message.set("Invalid file selected!")
    else:
        open_label.config(fg='green')
        open_message.set("File loaded: {}".format(file.split('/')[-1]))
        initial_file = file
    convert_message.set("")
    save_message.set("")


def convert_file(message, label, format, type):
    global initial_file
    global converted_file

    if initial_file is None:
        label.config(fg='red')
        message.set("No valid file selected for the conversion! Please open a valid file")
        return
    elif converted_file is not None:
        label.config(fg='red')
        message.set("File has already been converted! Please open another valid file")
        return

    extension = initial_file.split('/')[-1].split('.')[-1]
    if not os.path.isfile(initial_file):
        label.config(fg='red')
        message.set("The file you want to convert has been moved or cancelled!")
    elif extension == type:
        label.config(fg='red')
        message.set("You cannot convert the file in the same extension it has!")
    else:
        try:
            output_file = initial_file.split('.')[0] + ".{}".format(type)
            word = win32com.client.Dispatch("Word.Application")
            document = word.Documents.Open(FileName=os.path.abspath(initial_file), Visible=False)
            document.SaveAs2(FileName=os.path.abspath(output_file), FileFormat=format)
            document.Close()
            word.Quit()
        except Exception:
            label.config(fg='red')
            message.set("Something went wrong in the conversion!")
            return
        label.config(fg='green')
        message.set("Conversion to {} succesfully completed!".format(type))
        converted_file = output_file


def save_file(message, label):
    global initial_file
    global converted_file

    if initial_file is None:
        label.config(fg='red')
        message.set("No valid file selected for the conversion! Please open a valid file")
        return
    elif converted_file is None:
        label.config(fg='red')
        message.set("No file converted to be saved! Please convert the file selected")
        return

    extension = converted_file.split('/')[-1].split('.')[-1]
    initial_file = converted_file.split('/')[-1].split('.')[0]
    file_types = [('All files', '*')]
    output_file = filedialog.asksaveasfilename(parent=bottom_frame, confirmoverwrite=True, title='Save As',
                                               initialdir="/",defaultextension='.'+extension, initialfile=initial_file,
                                               filetypes=file_types)

    output_extension = output_file.split('/')[-1].split('.')[-1]
    if not output_file:
        label.config(fg='green')
        message.set("File saved in the same folder of the initial one!")
    elif output_extension != extension:
        label.config(fg='red')
        message.set("Invalid file extension, it cannot be saved!")
        return
    else:
        if converted_file != output_file:
            try:
                copyfile(converted_file, output_file)
                os.remove(converted_file)
            except Exception:
                label.config(fg='red')
                message.set("Something went wrong in the saving phase!")
                return
            label.config(fg='green')
            message.set("File {} saved in the folder selected!".format(output_file.split('/')[-1]))
        else:
            label.config(fg='green')
            message.set("File {} saved in the same folder of the initial one".format(output_file.split('/')[-1]))
    initial_file = None
    converted_file = None


# bottom label
save_message = tk.StringVar()
save_label = tk.Label(bottom_frame, textvariable=save_message, font='Times', width=20, wraplength=200)
save_label.pack(side='right')

# save file button
save_image = tk.PhotoImage(file=resource_path("save.png"))
save_button = tk.Button(bottom_frame, image=save_image, text='Save file', bg='white', fg='black', compound='right',
                        command=lambda: save_file(save_message, save_label), activeforeground='black',
                        width=195)
save_button.pack(side="left", pady=10)

# center label
convert_message = tk.StringVar()
convert_label = tk.Label(center_frame, textvariable=convert_message, font='Times', width=20, wraplength=180)
convert_label.pack(side='right')

# convert to pdf button
pdf_button = tk.Button(center_frame, text='Convert to PDF', bg='red', fg='snow', height=2,
                       command=lambda: convert_file(convert_message, convert_label, 17, "pdf"))
pdf_button.pack(side="left", padx=4, pady=25)

# convert to word button
word_button = tk.Button(center_frame, text='Convert to WORD', bg='dodger blue', fg='snow', height=2,
                        command=lambda: convert_file(convert_message, convert_label, 16, "docx"))
word_button.pack(side="left", pady=25)

# top label
open_message = tk.StringVar()
open_label = tk.Label(top_frame, textvariable=open_message, font='Times', width=20, wraplength=200)
open_label.pack(side='right')

# open file button
folder_image = tk.PhotoImage(file=resource_path("folder.png"))
open_button = tk.Button(top_frame, image=folder_image, text='Open file', bg='white', fg='black', compound='right',
                        command=lambda: open_file(open_message, open_label, convert_message, save_message),
                        activeforeground='black', width=195)
open_button.pack(side='left', pady=10)

master.mainloop()
