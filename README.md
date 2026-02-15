# Resume Generator

A Python-based automation tool that generates professional resumes in **LaTeX (.tex)** format using structured data from Excel files.
This project modularizes each resume section and compiles them into a complete, ready-to-build resume template powered by **Awesome-CV**.

---

## Features

* Reads candidate data from Excel sheets
* Modular resume section generators
* Generates LaTeX `.tex` files automatically
* Structured folder output per candidate
* Batch processing support via shell script
* Uses **Awesome-CV** professional resume template
* DOB validation & formatting support
* Easy customization & extensibility

---

## Project Structure

```
Resume Generator/
│
├── main.py                # Entry point
├── execute.sh             # Batch execution script
│
├── profile.py             # Profile section generator
├── dob-profile.py         # Profile generator with DOB validation
├── summary.py             # Summary section
├── skills.py              # Skills section
├── competencies.py        # Key competencies section
├── experience.py          # Work experience section
├── education.py           # Education section
│
└── excel_files/           # Input Excel files directory
```

---

## Excel File Format

Each Excel file must contain the following sheets:

| Sheet Name     | Description         |
| -------------- | ------------------- |
| `profile`      | Personal details    |
| `summary`      | Career summary      |
| `skills`       | Skills categorized  |
| `competencies` | Key strengths       |
| `experience`   | Work history        |
| `education`    | Academic background |

---

### Profile Sheet Fields (Order)

1. First Name
2. Last Name
3. Designation
4. Mobile Number
5. Email
6. Date of Birth
7. GitHub Profile (Optional)

---

## ⚙️ Installation

### 1️. Clone the Repository

```bash
git clone https://github.com/your-username/resume-generator.git
cd resume-generator
```

### 2️. Install Dependencies

```bash
pip install openpyxl
```

---

## Usage

###  Run for Single Excel File

```bash
python main.py data.xlsx
```

---

### Batch Process Multiple Excel Files

Place all Excel files inside:

```
excel_files/
```

Then run:

```bash
bash execute.sh
```

This will:

* Process each Excel file
* Generate `.tex` files
* Create structured folders automatically

---

## Output Structure

For an input file `john_doe.xlsx`:

```
john_doe/
│
├── main_resume.tex
└── resume/
    ├── summary.tex
    ├── skills.tex
    ├── competencies.tex
    ├── experience.tex
    └── education.tex
```

---

## Compile to PDF

After generating `.tex` files, compile using LaTeX:

```bash
pdflatex main_resume.tex
```

> Requires LaTeX distribution (TeX Live / MiKTeX).

---

## DOB Formatting Support

`dob-profile.py` includes:

* Multiple date format parsing
* Automatic conversion to `DD-MM-YYYY`
* Invalid format detection

Supported formats include:

* `DD/MM/YYYY`
* `YYYY-MM-DD`
* `MM-DD-YYYY`
* `DD Month YYYY`

---

## Template Dependency

This project uses the **Awesome-CV** LaTeX template.

Install or download from:

```
https://github.com/posquit0/Awesome-CV
```

Ensure the template files are available in your LaTeX environment.

---

## Customization

You can easily:

* Add new resume sections
* Modify LaTeX styling
* Extend Excel schema
* Integrate PDF auto-compilation

---

## How It Works

1. Excel data is read using **openpyxl**
2. Each module extracts its section data
3. `.tex` files are generated per section
4. `main_resume.tex` imports all sections
5. Final resume is compiled via LaTeX

---

## Requirements

* Python 3.x
* openpyxl
* LaTeX Distribution (TeX Live / MiKTeX)
* Awesome-CV Template

---

## Error Handling

The system checks for:

* Missing Excel directory
* Invalid files
* Script execution failures

Batch execution stops if any file fails.

---

## Future Enhancements

* Direct PDF generation
* Web UI for resume creation
* JSON/CSV input support
* Cloud resume storage
* Theme customization
