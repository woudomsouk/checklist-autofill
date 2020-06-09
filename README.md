# Checklist Autofill Background

A checklist needs to be completed for every drawing that is signed off on to ensure that all necessary gates are completed. However, the document is unnecessarily time consuming, and requires a lot of time to fill out redundant information.

The main issues are as follows:

    1. Fields that are marked as "N/A" based on the project stage.
    2. Multiple fields that constantly need to be filled in with information such as "Part number, part name, revision level, designer
        name, change folder." etc.
    3. Complicated naming scheme required for company sorting method.
    
This code solves those problems by taking information input that a user would log into Excel, filling out the required/known fields and generating a .PDF of file named appropriately, which can then be finished manually.

Ex. Information such as whether the project is at the quotation or production phase is indicated in the file naming convention, so these inputs would be used not only in the file naming, but also fill in relevant boxes as "N/A" based on project phase.

## Usage

Code relies on Microsoft Office formatting (Excel and Word).

Code is implemented in VBA (in Developer Tools), and fields to be filled use built in "Bookmark" feature in Microsoft Word.

## Additional Features

Code will also check for version of checklist on local network, so that if the checklist is ever updated it will prompt user to update the code to reflect new checklist.
