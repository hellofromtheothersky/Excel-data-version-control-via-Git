# Excel-data-version-control-via-Git
GitExcely
A tool for spreadsheet version control in Excel, created by Python and using with Git
# Getting Started
## Mechanism
<p align="center">
  <img src="https://github.com/user-attachments/assets/e16764dd-f8d0-44bb-8cd4-fc6ea526424f" />
</p>

## Quickstart
- Install:

  `Python` and `Git` installed, then:

  `pip install gitexcel`

- Create a new project/repository:

  `gitexcel init --path <PATH TO PROJECT NAME>`

- Or Clone an existing GitExcel repository:

  `gitexcel clone --path <PATH TO PROJECT NAME> --url <URL of REPO>`

- Here is the project structure

      . Project Name
      ├── EXCEL/
      │   ├── (Store all *.xlsx here, subfolder accepted)
      ├── EXCEL_TEXT/
      │   ├── (Excel data will be parsed into text here)
      ├── EXCEL_METADATA.json/
      │   (mannually setup header line and keys in Excel sheets)
      ├── CHANGES.log/
      │   Changes summary of Excel when before running git push (automatically removed before pushing)
      ├── DEBUGS.log/
      │   Show jobs running when run git commit (automatically removed before pushing)

# Why do I need GitExcel












