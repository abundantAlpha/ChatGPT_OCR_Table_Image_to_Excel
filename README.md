# Overview
Retreives most recent item from user's clipboard, reads the contents of the item using the native OCR capabilities of ChatGPT (gpt-4o), converts the content to a CSV file, and opens the file in Excel. The intended use is for the user to copy an image of a table using the Snipping tool, and convert it to numerical data in Excel using ChatGPT's OCR abilities. 

# The Interface
There is no interface by design. The user should use AutoHotkey to call the main.py function and execute the code based on a user-defined hotkey. 
