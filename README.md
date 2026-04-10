-----

# FSHN-Java-Automator 🚀

**FSHN-Java-Automator** is a specialized laboratory report generator for Informatics students. It eliminates the "busy work" of university assignments by automatically compiling Java source code, organizing screenshots, and building a perfectly formatted Microsoft Word document (`.docx`) that matches the strict standards of FSHN.

-----

## ✨ New in v2.0: The "Folder-Based" Workflow

The latest update focuses on maximum automation. Instead of manual uploads, the tool now watches your local directory:

  * **Automatic Image Alignment:** Drop screenshots into exercise-specific folders; the app tucks them into the document automatically.
  * **Professional Side-by-Side Layout:** Matches the "PDF style" where screenshots float next to or above your code blocks.
  * **Multi-Screenshot Support:** Handles multiple outputs per exercise without breaking the layout.
  * **Integrated Java Launcher:** Compile and run `.java` files directly from the UI to verify output before capturing.

-----

## 🚀 Getting Started

### 1\. Prerequisites

Before running the app, ensure you have the following installed:

  * [Python 3.8+](https://www.python.org/downloads/) (Check **"Add Python to PATH"** during installation).
  * **Java JDK**: Ensure `javac` and `java` are accessible in your terminal.

### 2\. Installation

Open your terminal (PowerShell or Command Prompt) and run:

```powershell
# 1. Navigate to your project folder
cd "C:\Path\To\Your\Automation Folder"

# 2. Create and activate a virtual environment
python -m venv .venv
.\.venv\Scripts\activate

# 3. Install dependencies
pip install python-docx Pillow streamlit
```

-----

## 📂 Project Structure

To use the **v2.0 Folder Workflow**, organize your files as follows:

```text
📁 Your_Lab_Project/
├── 📄 lab_report_builder.py
├── 📁 screenshots/
│   ├── 📁 Ushtrimi1/
│   │   └── 🖼️ output.png
│   └── 📁 Ushtrimi2/
│       ├── 🖼️ step1.png
│       └── 🖼️ step2.png
└── 📁 java_files/
    ├── ☕ Ushtrimi1.java
    └── ☕ Ushtrimi2.java
```

-----

## 🎮 Usage

1.  **Launch the App:**
    ```powershell
    streamlit run lab_report_builder.py
    ```
2.  **Fill Global Info:** Enter your name, group, and the assignment title.
3.  **Add Exercises:**
      * Paste your code or select a `.java` file.
      * The app automatically pulls images from `screenshots/UshtrimiX`.
      * Choose between **"Side-by-Side"** (compact) or **"Stacked"** (full-width) layouts.
4.  **Generate:** Click **"Build Document"** to export your finished `.docx`.

-----

## ⚙️ Customization

You can adjust the report's visual style directly in `lab_report_builder.py`. Locate the `TEMPLATE` dictionary at the top of the script:

```python
TEMPLATE = {
    "header_font": "Times New Roman",
    "header_size_pt": 14,
    "code_font": "Courier New",
    "code_size": 8,
    "image_width_inches": 5.5
}
```

-----

## 📦 Distribution

To share this tool with colleagues who don't have Python installed, you can compile it into a standalone `.exe` using **PyInstaller**:

1.  **Install PyInstaller:**
    ```powershell
    pip install pyinstaller
    ```
2.  **Build the Executable:**
    ```powershell
    pyinstaller --onefile --windowed lab_report_builder.py
    ```
3.  The finished app will be located in the `/dist` folder.

-----

> [\!TIP]
> **Image Quality Note:** For the best results, use high-resolution screenshots. The automator will scale them to fit margins perfectly while maintaining the original aspect ratio to avoid distortion.

-----
