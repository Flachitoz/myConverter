import tkinter as tk
from utils.utilities import resource_path
from handler.fileHandler import FileHandler


# initialize initial and converted file path
fileHandler = FileHandler()

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


# bottom label
save_message = tk.StringVar()
save_label = tk.Label(bottom_frame, textvariable=save_message, font='Times', width=20, wraplength=200)
save_label.pack(side='right')

# save file button
save_image = tk.PhotoImage(file=resource_path("save.png"))
save_button = tk.Button(bottom_frame, image=save_image, text='Save file', bg='white', fg='black', compound='right',
                        command=lambda: fileHandler.save_file(save_message, save_label), activeforeground='black',
                        width=195)
save_button.pack(side="left", pady=10)

# center label
convert_message = tk.StringVar()
convert_label = tk.Label(center_frame, textvariable=convert_message, font='Times', width=20, wraplength=180)
convert_label.pack(side='right')

# convert to pdf button
pdf_button = tk.Button(center_frame, text='Convert to PDF', bg='red', fg='snow', height=2,
                       command=lambda: fileHandler.convert_file(convert_message, convert_label, 17, "pdf"))
pdf_button.pack(side="left", padx=4, pady=25)

# convert to word button
word_button = tk.Button(center_frame, text='Convert to WORD', bg='dodger blue', fg='snow', height=2,
                        command=lambda: fileHandler.convert_file(convert_message, convert_label, 16, "docx"))
word_button.pack(side="left", pady=25)

# top label
open_message = tk.StringVar()
open_label = tk.Label(top_frame, textvariable=open_message, font='Times', width=20, wraplength=200)
open_label.pack(side='right')

# open file button
folder_image = tk.PhotoImage(file=resource_path("folder.png"))
open_button = tk.Button(top_frame, image=folder_image, text='Open file', bg='white', fg='black', compound='right',
                        command=lambda: fileHandler.open_file(open_message, open_label, convert_message, save_message),
                        activeforeground='black', width=195)
open_button.pack(side='left', pady=10)

master.mainloop()
