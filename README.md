# Registration to Registry Script

This guide explains how to use the script even if you do not know programming.

The package should contain these files:

- `registration_to_registry.py`
- `requirements.txt`
- `run_windows.bat`
- `run_unix.sh`
- the two Excel files you download each time

---

## What this script does

The script reads the downloaded **Registration Form (Responses)** Excel file, filters people from the selected region, and writes them into the matching sheet in the **Registry Database** Excel file.

It appends new people below the existing rows.

It is currently set to use a specific region inside the Python file. If you want to use another region, you must change these two lines near the top of the Python file:

```python
REGION_NAME = "South West"
REGISTRY_SHEET_NAME = "South West"
```

Example for London:

```python
REGION_NAME = "London"
REGISTRY_SHEET_NAME = "London"
```

---

## Before you start

You need:

1. **Python installed** on your computer
2. This script package in one folder
3. The downloaded Excel files placed in the same folder

---

## Step 1: Check whether Python is installed

### On Windows

1. Press **Win + R**
2. Type `cmd`
3. Press **Enter**
4. In the black window, type:

```bash
python --version
```

If you see something like `Python 3.12.x` or `Python 3.13.x`, Python is installed.

If you get an error, install Python first from the official Python website, then try again.

### On macOS / Linux

Open Terminal and run:

```bash
python3 --version
```

If needed, your system may use `python` instead of `python3`.

---

## Step 2: Put everything in one folder

Create one folder and place all files inside it.

Example:

```text
registration-folder/
├── registration_to_registry.py
├── requirements.txt
├── run_windows.bat
├── run_unix.sh
├── Registry Database.xlsx
└── Lan Berambeh Anak Sarawak 2026 UK Edition_ Registration Form (Responses).xlsx
```

Important:
- Keep the Python file and both Excel files in the **same folder**
- The Excel filenames should still contain words like `Registry`, `Database`, `Registration`, or `Responses`

---

## Step 3: Change the region if needed

If you are not using South West, open `registration_to_registry.py` in Notepad, VS Code, or any text editor.

Find these lines near the top:

```python
REGION_NAME = "South West"
REGISTRY_SHEET_NAME = "South West"
```

Change both values to your own region.

Save the file.

---

## Step 4: Create and use a virtual environment manually

A virtual environment keeps the required Python library inside this folder so it does not affect the rest of the computer.

### On Windows

1. Open Command Prompt in the folder
   - Open the folder in File Explorer
   - Click the address bar
   - Type `cmd`
   - Press **Enter**

2. Create the virtual environment:

```bash
python -m venv .venv
```

3. Activate it:

```bash
.venv\Scripts\activate
```

After activation, you may see `(.venv)` at the start of the command line.

4. Install the required library:

```bash
python -m pip install -r requirements.txt
```

5. Run the script:

```bash
python registration_to_registry.py
```

6. When finished, you can leave the virtual environment by typing:

```bash
deactivate
```

### On macOS / Linux

1. Open Terminal in the folder
2. Create the virtual environment:

```bash
python3 -m venv .venv
```

3. Activate it:

```bash
source .venv/bin/activate
```

4. Install the required library:

```bash
python -m pip install -r requirements.txt
```

5. Run the script:

```bash
python registration_to_registry.py
```

6. When finished:

```bash
deactivate
```

---

## Easier method: use the launcher scripts

You do not need to activate the virtual environment manually if you use the launcher files below.

They will:

1. create the virtual environment if it does not already exist
2. install the required library
3. run the Python script

### Windows

Double-click:

- `run_windows.bat`

or open Command Prompt in the folder and run:

```bash
run_windows.bat
```

### macOS / Linux

Open Terminal in the folder and run:

```bash
chmod +x run_unix.sh
./run_unix.sh
```

---

## What the script fills

The updated script now writes:

- **Actual ID** = the row number from the response form
- **Name**
- **Email**
- **Contact Number (Whatsapp)**
- **Status**
- **Number of dependents**
- **City**
- **Which part of Sarawak**
- **University**

It does **not** fill:

- `Confirmation email sent?`

That column is left for the secretary to use.

---

## How duplicates are handled

The updated script skips duplicates using:

- **Actual ID**
- **Email**

This is safer than checking name.

So if two different people have the same name but different email and Actual ID, both can still be added.

---

## What happens after running

The script saves the updated downloaded registry workbook in the same folder.

After that:

1. Open the updated registry Excel file
2. Check the correct region sheet
3. Copy the new data if needed
4. Paste it into the shared Excel file

---

## Expected terminal output

You may see something like:

```text
Registry file   : Registry Database.xlsx
Form file       : Lan Berambeh Anak Sarawak 2026 UK Edition_ Registration Form (Responses).xlsx
Region checked  : South West
Rows found      : 12
Rows added      : 3
Skipped existing: 9
Done.
```

Meaning:

- `Rows found` = total people from that region in the form
- `Rows added` = new rows written into the registry
- `Skipped existing` = rows skipped because they already exist

---

## Troubleshooting

### 1. `python` is not recognized

Python is not installed, or not added to PATH.

Install Python first.

### 2. `No module named openpyxl`

Run:

```bash
python -m pip install -r requirements.txt
```

### 3. It cannot find the Excel files

Make sure both Excel files are in the same folder as the Python script.

### 4. It says the sheet does not exist

Check the worksheet name inside the registry workbook and make sure:

```python
REGISTRY_SHEET_NAME = "Your Region"
```

matches exactly.

### 5. Nothing new was added

Possible reasons:

- there are no people from that region in the form
- those people were already copied before
- the selected region name does not exactly match the form value

---

## Quick summary

### First time setup

1. Install Python
2. Put all files in one folder
3. Change the region in the Python file if needed
4. Run `run_windows.bat` on Windows, or `./run_unix.sh` on macOS/Linux

### Every later use

1. Download the latest 2 Excel files
2. Replace the old Excel files in the folder
3. Run the launcher again
4. Copy the updated registry data into the shared Excel file

