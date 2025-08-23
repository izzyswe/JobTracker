import pyperclip as pc

# Get the text from the clipboard
clipboard_content = pc.paste()

# Print the retrieved text
print(f"The text from the clipboard is: '{clipboard_content}'")
