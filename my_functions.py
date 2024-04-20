import os
import json
import tkinter as tk
from tkinter import messagebox


def show_alert(title, message):
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Show a message box
    messagebox.showinfo(title, message)


def get_used_file(usage_path):

    todo_folder_path = os.path.join(usage_path)
    todo_files = [f for f in os.listdir(
        todo_folder_path) if f.endswith(".xlsx") or f.endswith(".xlsm")]
    if not todo_files:
        error_message = f"W Folderze {todo_folder_path} nie ma zadnego pliku do przetworzenia!"
        show_alert("Brak Pliku", error_message)
        print(error_message)
    else:
        todo_file_name = todo_files[0]
        todo_file_path = os.path.join(todo_folder_path, todo_file_name)
        return todo_file_path, todo_file_name


def read_config(file_path):
    with open(file_path, 'r') as file:
        config_data = json.load(file)
    
    return config_data


def remove_folder_contents(folder_path):
    print(os.listdir(folder_path))
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isfile(item_path):
            os.remove(item_path)