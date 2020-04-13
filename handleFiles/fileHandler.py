import os
import win32com.client
from shutil import copyfile
from tkinter import filedialog, StringVar, Label


class FileHandler:

    def __init__(self):
        self.input_file = None
        self.output_file = None

    def open_file(self, open_message: StringVar, open_label: Label, convert_message: StringVar, save_message: Label) -> None:
        self.input_file = None
        self.output_file = None

        file_types = [("WORD", "*.docx"), ("PDF", "*.pdf"), ('All files', '*')]
        accepted_extensions = {'pdf', 'docx'}
        file = filedialog.askopenfilename(initialdir='/', title="Select file", filetypes=file_types)
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
            self.input_file = file
        convert_message.set("")
        save_message.set("")

    def convert_file(self, message: StringVar, label: Label, file_format: int, file_type: str) -> None:
        if self.input_file is None:
            label.config(fg='red')
            message.set("No valid file selected for the conversion! Please open a valid file")
            return
        elif self.output_file is not None:
            label.config(fg='red')
            message.set("File has already been converted! Please open another valid file")
            return

        extension = self.input_file.split('/')[-1].split('.')[-1]
        if not os.path.isfile(self.input_file):
            label.config(fg='red')
            message.set("The file you want to convert has been moved or cancelled!")
        elif extension == file_type:
            label.config(fg='red')
            message.set("You cannot convert the file in the same extension it has!")
        else:
            try:
                output_file = self.input_file.split('.')[0] + ".{}".format(file_type)
                word = win32com.client.Dispatch("Word.Application")
                document = word.Documents.Open(FileName=os.path.abspath(self.input_file), Visible=False)
                document.SaveAs2(FileName=os.path.abspath(output_file), FileFormat=file_format)
                document.Close()
                word.Quit()
            except Exception:
                label.config(fg='red')
                message.set("Something went wrong in the conversion!")
                return
            label.config(fg='green')
            message.set("Conversion to {} succesfully completed!".format(file_type))
            self.output_file = output_file

    def save_file(self, message: StringVar, label: Label) -> None:
        if self.input_file is None:
            label.config(fg='red')
            message.set("No valid file selected for the conversion! Please open a valid file")
            return
        elif self.output_file is None:
            label.config(fg='red')
            message.set("No file converted to be saved! Please convert the file selected")
            return

        extension = self.output_file.split('/')[-1].split('.')[-1]
        initial_file = self.output_file.split('/')[-1].split('.')[0]
        file_types = [('All files', '*')]
        saved_file = filedialog.asksaveasfilename(confirmoverwrite=True, title='Save As', initialdir="/",
                                                  defaultextension='.'+extension, initialfile=initial_file,
                                                  filetypes=file_types)

        output_extension = saved_file.split('/')[-1].split('.')[-1]
        if not saved_file:
            label.config(fg='green')
            message.set("File saved in the same folder of the initial one!")
        elif output_extension != extension:
            label.config(fg='red')
            message.set("Invalid file extension, it cannot be saved!")
            return
        else:
            if self.output_file != saved_file:
                try:
                    copyfile(self.output_file, saved_file)
                    os.remove(self.output_file)
                except Exception:
                    label.config(fg='red')
                    message.set("Something went wrong in the saving phase!")
                    return
                label.config(fg='green')
                message.set("File {} saved in the folder selected!".format(saved_file.split('/')[-1]))
            else:
                label.config(fg='green')
                message.set("File {} saved in the same folder of the initial one".format(saved_file.split('/')[-1]))

        self.input_file = None
        self.output_file = None
