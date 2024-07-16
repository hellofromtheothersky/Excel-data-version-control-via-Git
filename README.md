# Excel-data-version-control-via-Git
GitExcely
A tool for spreadsheet version control in Excel, created by Python and using with Git
# Getting Started
## Mechanism
<p align="center">
  <img src="https://github.com/user-attachments/assets/e16764dd-f8d0-44bb-8cd4-fc6ea526424f" />
</p>

User edit the Excel `->` commit `->` Pre-commit hook is triggerd `->` Parse Excel into text `->` Check & manage diff via text

User edit the Excel text version (e.g. to solve conflict) `->` commit `->` Pre-commit hook is triggerd `->` Gen Excel from text

**GitExcely create a repo that every Excel file in these has its text version, they always sync together before being pushed**

## Quickstart
- Install:

  `Python` and `Git` installed, then:

  `pip install gitexcely`

- Create a new project/repository:

  `gitexcely init --path <PATH TO PROJECT NAME>`

- Or Clone an existing GitExcel repository:

  `gitexcely clone --path <PATH TO PROJECT NAME> --url <URL of REPO>`

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
# Demo
[GitExcely Demo Document](https://github.com/hellofromtheothersky/Excel-data-version-control-via-Git/blob/main/GitExcely%20Demo.pdf)
# Backgrounds
- I have worked on a large-scale data migration project in an organization. Ensuring the data mapping rules are correct is important, and the project has more than 200 tables to migrate, with the detailed specifications down to the column level involving more than 2,000 columns. Excel is used by Business Analysts to specify the requirements, and many other roles still review it for their purposes.
- These Excel files often have unexpected changes due to someone's mistakes or faults. These unexpected changes are sometimes harmless, but sometimes they can cause significant errors when the pipeline reads them and spoil downstream tasks.
- The current MS version history is not a good deal for complicated Excel file like that, and it lack of a real git mechanism.
- And this the starting point of GitExcely.
# Limitations (will improve soon)
- Styles supported (the other rest can not be tracked): font, border, fill, number format, protection, alignment, column width
- Performance: for larger Excel file, it take time to parse into text (due to large numnber of small files problem when let every row be parsed into an csv and json file)













