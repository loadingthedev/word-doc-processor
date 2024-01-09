# Word Document Processor

This Python script automates the process of opening `.docx` files in the current directory, processing them, and saving the processed versions in a subfolder named `processed_doc`.

## Prerequisites

Before running the script, ensure the following prerequisites are met:

1. **Python Installation**: The script requires Python. If not already installed, download and install Python from [python.org](https://www.python.org/downloads/).

2. **pywin32 Library**: The script uses `pywin32` for COM automation with Microsoft Word. Install it using pip:

   ```bash
   pip install pywin32


1. Microsoft Word: Microsoft Word must be installed on your system, as the script uses Word's application interface.

## Installation
1. Download the Script: Download the process_docs.py script to your desired directory.

2. Prepare the Environment: Make sure Python and pywin32 are installed as mentioned in the prerequisites.

## Usage
1. Prepare Your Documents: Place the .docx files you wish to process in the same directory as the process_docs.py script.

2. Run the Script: Open a command prompt or terminal in the script's directory. Execute the script with Python:
   
      ```bash
      python script.py
      ```
3. Access Processed Files: Processed documents will be available in the processed_doc subfolder in the script's directory.

## Script Behavior

- The script processes all .docx files in its directory.
  
- It runs Word in the background (invisible mode). Close other Word instances for smooth operation.

- Processed files are saved in a subfolder named processed_doc


## Safety Precautions

- Only files intended for processing should be in the folder with the script.
- Always back up your original files to prevent accidental data loss.
