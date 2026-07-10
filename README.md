# mcq-generator
Shuffles mcqs along with its options


Here is a step-by-step guide to saving this project on GitHub so you can easily run it on any machine in the future.

Step 1: Create a GitHub Repository
Go to GitHub and log in (create a free account if you don't have one).

Click the + icon in the top-right corner and select New repository.

Fill in the following details:

Repository name: mcq-generator

Public/Private: Choose either (Private is best if you want to keep your exam questions hidden).

Initialize this repository with: Check the box that says Add a README file.

Click Create repository.

Step 2: Upload Your Script and CSV File (Directly on GitHub)
You don't need a terminal yet; you can upload your files directly through the website:

Inside your new repository page, click the Add file button and choose Upload files.

Create a file on your local computer named generate_mcqs.py using standard Notepad or VS Code, and paste the code below into it.

Drag and drop generate_mcqs.py and your data file Questions.xlsx - Sheet1.csv into the upload box on GitHub.

Click Commit changes at the bottom of the page to save them.

The Python Script to Use
Make sure your generate_mcqs.py contains this exact universal version, which handles both Windows/Linux path differences and Excel encoding problems out of the box:

Python
import pandas as pd
import random
import os
from docx import Document

# 1. Dynamically find the folder where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, 'Questions.xlsx - Sheet1.csv')

# 2. Safely read the file, handling Excel's default encoding formats
try:
    df = pd.read_csv(file_path, encoding='latin-1')
except:
    df = pd.read_csv(file_path, encoding='cp1252')

def create_paper(data, version_num):
    # 3. Shuffle question rows
    shuffled_df = data.sample(frac=1).reset_index(drop=True)
    
    # 4. Initialize Word Document
    doc = Document()
    doc.add_heading(f'Exam Paper - Version {version_num}', 0)
    
    for i, row in shuffled_df.iterrows():
        # Add question with native Word numbering layout
        p = doc.add_paragraph(style='List Number')
        p.add_run(str(row['Question'])).bold = True
        
        # 5. Extract and shuffle options (A, B, C, D sequence)
        opts = [str(row['OptionA']), str(row['OptionB']), str(row['OptionC']), str(row['OptionD'])]
        random.shuffle(opts)
        
        # 6. Apply static labels to randomized options
        labels = ['A', 'B', 'C', 'D']
        for label, text in zip(labels, opts):
            doc.add_paragraph(f"{label}) {text}")
        
        doc.add_paragraph() # Spacer between questions

    # 7. Save document inside the same script directory
    output_path = os.path.join(script_dir, f'MCQ_Version_{version_num}.docx')
    doc.save(output_path)

# 8. Loop to generate 5 distinct variants
for i in range(1, 6):
    create_paper(df, i)
    print(f"Successfully generated Version {i}")


Step 3: How to Run This on Any System Later
When you switch computers or want to run this in the cloud for free, you have two great options:

Option A: Run it for Free in the Cloud (via GitHub Codespaces)
GitHub gives every free account 60 hours per month of free cloud-hosted VS Code access.

Open your repository on GitHub.

Click the green <> Code button.

Select the Codespaces tab and click Create codespace on main.

A browser-based version of VS Code will launch. Open the terminal at the bottom and type:

Bash
pip install pandas python-docx
python generate_mcqs.py

Your 5 Word files will appear right in the left sidebar. Right-click them and select Download.

Option B: Run it on a New Local Computer
Install Git and Python on your new machine.

Open your command prompt/terminal and clone your repository down:

Bash
git clone https://github.com/your-username/mcq-generator.git
Open that folder, open your terminal, install your modules, and run the file:

Bash
pip install pandas python-docx
python generate_mcqs.py
Line-by-Line Explanation of the Universal Upgrades
os.path.abspath(__file__): Finds the explicit, unique file system path of where generate_mcqs.py sits. This guarantees it works seamlessly whether executed inside a local Windows folder, a Linux server, or a GitHub cloud environment.

os.path.join(...): Prevents system crashes caused by directory slashes. Windows uses backslashes (\) while Mac, Linux, and Cloud environments use forward slashes (/). This function auto-corrects them based on whatever computer is executing the code.

try: ... except: ...: Bypasses character errors. Standard Python crashes if it hits symbols it doesn't recognize. This block forces Python to use alternative decoders (latin-1 or cp1252) ensuring your Excel-exported CSV loads smoothly anywhere.
