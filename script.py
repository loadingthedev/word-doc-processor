import os
import win32com.client as win32

# Get the directory where the script is located
script_directory = os.path.dirname(os.path.abspath(__file__))
destination_folder = os.path.join(script_directory, 'processed_doc')

# Create the destination folder if it doesn't exist
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

word = win32.Dispatch('Word.Application')
word.Visible = False

for filename in os.listdir(script_directory):
    if filename.endswith('.docx'):
        doc_path = os.path.join(script_directory, filename)
        new_doc_path = os.path.join(destination_folder, filename)

        doc = word.Documents.Open(doc_path)
        doc.SaveAs2(new_doc_path)
        doc.Close()

word.Quit()
