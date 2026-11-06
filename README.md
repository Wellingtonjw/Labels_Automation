# Labels Automation

Python scripts for automating label workflows with NiceLabel, Excel, PDF and image processing.

## Scripts

| File | Purpose |
|---|---|
| `geraetiquetaspdf.py` | Uses desktop automation to populate NiceLabel variables from an Excel workbook and save label PDFs. |
| `identifica.py` | Groups image files by perceptual similarity using pHash. |
| `identifica2.py` | Groups images and PDFs by visual similarity; also creates PNG previews for PDFs. |
| `Convert_Finish_Image_Nicelabel_in_PDF(1).py` | Matches `LINHA_XXXX` folders to Excel descriptions and moves PDFs or converts images into PDFs. |
| `Convert_PDF_in_PNG.py` | Converts the first page of each PDF in a folder to PNG. |
| `Create_Groups_In_Folder_image.PY` | Groups images by perceptual similarity and records file names and detected dimensions. |
| `Create_Image_of_NiceLabel(1).py` | Uses the NiceLabel SDK and Excel automation to render label graphics by spreadsheet row. |
| `Create_List_Lay_on_Excel.py` | Assigns layout groups to Excel rows using codes, dimensions and text similarity. |

## Important notes

- These scripts were uploaded as provided; their logic and configuration values were not changed.
- Several scripts contain machine-specific Windows paths. Update the configuration constants before running them on another computer.
- The NiceLabel automation scripts require a compatible Windows environment and the relevant NiceLabel installation/SDK.
- Install each script's required Python packages before execution. Common dependencies across this collection include `openpyxl`, `Pillow`, `ImageHash`, `PyMuPDF`, `pyautogui`, `pyperclip`, `pywinauto`, `pywin32`, and `pythonnet`.
- Review the repository visibility before adding any further work-related files. This repository is currently public.

## Running

Open a terminal in the folder containing the script and run:

```powershell
python .\script_name.py
```

Replace `script_name.py` with the file you want to execute.
