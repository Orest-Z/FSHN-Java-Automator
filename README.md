**FSHN-Java-Automator**

A Python-based automation tool designed for students at the Faculty of Natural Sciences (FSHN) to generate perfectly formatted .docx lab reports. This tool eliminates the tedious work of manually pasting Java code and resizing screenshots into Microsoft Word.


🛠️ Java Lab Report Builder v2.0

Automated Documentation for Java Informatics Students

This tool is designed to take the "busy work" out of university laboratory assignments. It automatically compiles your Java source code, organizes your screenshots, and builds a perfectly formatted Microsoft Word document (.docx) following the layout standards used in professional informatics reports.
✨ New in v2.0: The "Folder-Based" Workflow

The latest update focuses on flexibility. Instead of manually uploading images one by one, the tool now watches your folders:

    Automatic Image Alignment: Place screenshots in exercise-specific folders, and the app will tuck them into the document automatically.

    Side-by-Side Layout: Matches the professional PDF style where screenshots appear next to or above your code blocks.

    Multi-Screenshot Support: If one exercise has three different outputs, the app handles all of them.

    Java Launcher: Includes a built-in helper to compile and run your .java files directly from the UI.

🚀 Getting Started from Scratch
1. Prerequisites

Before running the app, ensure you have the following installed:

    Python 3.8+: https://www.python.org/downloads/ (Make sure to check "Add Python to PATH" during installation).

    Java JDK: Ensure javac and java are working in your terminal (required for the auto-run feature).

2. Installation

Open your terminal (PowerShell or Command Prompt) and run these commands in order:
PowerShell

# 1. Navigate to your project folder
cd "C:\Path\To\Your\Automation Folder"

# 2. Create a virtual environment (optional but recommended)
python -m venv .venv

# 3. Activate the environment
# On Windows:
.\.venv\Scripts\activate

# 4. Install the required libraries
pip install python-docx Pillow streamlit

📂 How to Organize Your Files

To use the v2.0 Folder Workflow, organize your project like this:
Plaintext

📁 Your_Lab_Project/
├── 📄 lab_report_builder.py
├── 📁 screenshots/
│   ├── 📁 Ushtrimi1/
│   │   └── 🖼️ output.png
│   ├── 📁 Ushtrimi2/
│   │   ├── 🖼️ step1.png
│   │   └── 🖼️ step2.png
└── 📁 java_files/
    ├── ☕ Ushtrimi1.java
    └── ☕ Ushtrimi2.java

🎮 Running the App

    Open your terminal in the project folder.

    Start the interface by typing:
    PowerShell

    streamlit run lab_report_builder.py

    A new tab will open in your web browser.

The 3-Step Workflow:

    Global Info: Enter your name (e.g., Orest Zogju), group, and the Lab Title.

    Add Exercises: * Paste your code or point to your .java file.

        The app will automatically look for images in the screenshots/UshtrimiX folder.

        Choose between "Side-by-Side" (compact) or "Stacked" (large images) layouts.

    Generate: Click "Build Document" to save your finished .docx file.

🛠️ Customization

To change the default fonts, sizes, or colors to match your professor's specific requirements, open lab_report_builder.py in a text editor and look for the TEMPLATE dictionary at the top:
Python

TEMPLATE = {
    "header_font": "Times New Roman",
    "header_size_pt": 14,
    "code_font": "Courier New",
    "code_size": 8,
    "image_width_inches": 5.5
}

📦 Distribution (Creating an .exe)

If you want to share this with friends so they can run it without installing Python, use PyInstaller:

    Install it: pip install pyinstaller

    Run the build: pyinstaller --onefile --windowed lab_report_builder.py

    The finished application will be in the /dist folder.

Note: Ensure your screenshots are high quality. The app will automatically scale them to fit the margins while maintaining the aspect ratio, preventing any stretching or distortion.