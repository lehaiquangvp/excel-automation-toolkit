# Excel Automation Toolkit

**Merge, split and remove duplicates in Excel files in one click. No macros, no formulas.**

> 💡 **Not a programmer?** Get the ready-to-use **Windows app (.exe)**: just double-click, nothing to install.
> 👉 **[Download the Windows app on Gumroad](GUMROAD_LINK_HERE)**

<!-- After uploading screenshots to the screenshots/ folder, uncomment:
![Merge files](screenshots/merge.png)
-->

## Features

| Tool | What it does |
|---|---|
| **Merge files** | Combine dozens of Excel/CSV files into one workbook, stacked into one sheet (columns matched by header) or kept as separate sheets. Optional "Source File" column. |
| **Split** | One file per sheet, **or** split by a column, e.g. one file per Region / Salesperson / Month. |
| **Remove duplicates** | By whole row or by chosen columns (e.g. Email). Can ignore case and extra spaces. Removed rows are kept in a separate sheet. |

✅ Your original files are never changed. Results are always saved as new files.
✅ Works with `.xlsx`, `.xlsm`, `.xls`, `.csv`.

## Run from source (Windows / Mac / Linux)

Requires **Python 3.10+**.

```bash
git clone https://github.com/lehaiquangvp/excel-automation-toolkit.git
cd excel-automation-toolkit
pip install -r requirements.txt
python excel_toolkit.py
```

On Windows you can also just double-click `run_windows.bat`.

Try it with the files in `sample files/`.

## Use it in your own scripts

```python
import excel_core as xl

xl.merge_files(["reports/"], "Merged.xlsx", mode="stack")
xl.split_by_column("Sales.xlsx", "Region", "by_region/")
xl.remove_duplicates("Contacts.xlsx", "Clean.xlsx", columns=["Email"])
```

## Build the .exe yourself

On Windows, double-click `build_exe.bat`. The app is created at `dist\ExcelToolkit.exe`.

## Support the project

If this saves you time, the **[Windows app on Gumroad](GUMROAD_LINK_HERE)** is the easiest way to use it and to support future updates. ⭐ Starring the repo also helps!

## License

MIT
