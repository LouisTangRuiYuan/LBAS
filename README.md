# Registration Form to Registry Database

This script automatically copies participant details from the downloaded **registration form Excel file** into the correct region sheet in the **registry database Excel file**.

It is currently configured for:

```python
REGION_NAME = "South West"
REGISTRY_SHEET_NAME = "South West"
```

You can change these two lines to use the script for another region.

---

## 1. Install Python

Make sure Python is installed on your laptop.

To check, open **Command Prompt** or **Terminal** and run:

```bash
python --version
```

If Python is not installed, install it first.

---

## 2. Install required library

This script uses **openpyxl**.

Run:

```bash
python -m pip install openpyxl
```

To check whether it installed correctly:

```bash
python -c "import openpyxl; print(openpyxl.__version__)"
```

---

## 3. Prepare the files

Download these 2 Excel files:

1. **Registry Database**
2. **Registration Form (Responses)**

Then place both Excel files in the **same folder** as the Python script.

Example folder:

```text
registration-folder/
├── registration_to_registry.py
├── Registry Database.xlsx
└── Lan Berambeh Anak Sarawak 2026 UK Edition_ Registration Form (Responses).xlsx
```

The script automatically searches for:
- the **registry file** using words like `registry` and `database`
- the **form file** using words like `registration` and `responses`

So do not remove those keywords from the filenames.

---

## 4. Change the region if needed

Open `registration_to_registry.py` and look for:

```python
REGION_NAME = "South West"
REGISTRY_SHEET_NAME = "South West"
```

### Meaning
- `REGION_NAME` = the region name as written in the **registration form**
- `REGISTRY_SHEET_NAME` = the sheet name inside the **registry database**

### Example
If you want to use it for London, change it to:

```python
REGION_NAME = "London"
REGISTRY_SHEET_NAME = "London"
```

If the region name in the form and the registry sheet name are different, set them accordingly.

Example:

```python
REGION_NAME = "South West"
REGISTRY_SHEET_NAME = "SW"
```

---

## 5. Open terminal in the folder

Go to the folder where the Python file and the 2 Excel files are stored.

### Option 1: Use File Explorer
Open the folder, click the address bar, type `cmd`, then press **Enter**.

### Option 2: Use terminal manually
Open Command Prompt and type:

```bash
cd "C:\path\to\your\folder"
```

Example:

```bash
cd "C:\Users\User\Desktop\lbas south west"
```

---

## 6. Run the script

Run:

```bash
python registration_to_registry.py
```

---

## 7. What the script does

The script will:

1. open the registration form Excel file
2. search through all responses
3. filter only people from the selected region
4. open the selected region sheet in the registry database
5. append new people below the existing rows
6. save the updated registry workbook

This means:
- existing rows in the registry stay at the top
- newly added people are written below them

---

## 8. Important note about duplicates

The current script avoids duplicates by checking whether the following already exist in the registry sheet:

- **Name**
- **Email**
- **Actual ID**

If **any one** of these matches, that row will be skipped.

### Example
If two different people have:
- the same full name
- but different email and Actual ID

the second person may still be skipped because the script treats same name as duplicate.

So after running the script, it is a good idea to do a quick manual check for people with exactly the same name.

---

## 9. After running the script

The script will show a summary like this:

```text
Registry file   : Registry Database.xlsx
Form file       : Lan Berambeh Anak Sarawak 2026 UK Edition_ Registration Form (Responses).xlsx
Region checked  : South West
Rows found      : 10
Rows added      : 3
Skipped existing: 7
Done.
```

### Meaning
- `Rows found` = total people from that region in the form
- `Rows added` = new people written into the registry
- `Skipped existing` = people not added because they were treated as duplicates

---

## 10. Final step

After the script updates the downloaded registry Excel file:

1. open the updated file
2. copy the newly updated region sheet data if needed
3. paste it into the **shared Excel file**

Make sure you paste into the correct region sheet in the shared registry workbook.

---

## 11. Notes

- Run the script again whenever new form responses come in.
- Old registry entries will stay on top.
- New entries will be added below.
- The script saves changes directly into the downloaded registry workbook in the same folder.

---

## 12. Troubleshooting

### Error: `ModuleNotFoundError: No module named 'openpyxl'`
Install the library:

```bash
python -m pip install openpyxl
```

### Script cannot find the Excel files
Make sure:
- both Excel files are in the same folder as the script
- the filenames still contain keywords like:
  - `registry`
  - `database`
  - `registration`
  - `responses`

### Sheet not found
Make sure `REGISTRY_SHEET_NAME` exactly matches the worksheet name in the registry workbook.

---

## 13. Summary

### To use:
1. install Python
2. install `openpyxl`
3. download the 2 Excel files
4. place them in the same folder as the script
5. change the region name if needed
6. open terminal in that folder
7. run the script
8. copy the updated data into the shared Excel file
