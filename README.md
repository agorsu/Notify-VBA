# Notify-VBA
Notify is an excel VBA macro designed to allow you to quickly create a single one-off word document from a JIRA data extract.

## Description
A simple mail merge alternative that can be shared on Onedrive, allowing users to choose and create a word document from a list of configured word templates. 
Word documents are generated and saved into the same folder as the spreadsheet using the key field data in the filename.

The Notify macro works by using find and replace method for any placeholders in the document that have the format <<columnName>> which matches the column name in the data tab.
It can support long text data and will retain the formatting of the placeholder in the template.

Two additional reserved fields have also been included:
<<TODAY>> 	  Current date when generated.
<<USERNAME>> 	Account name of the user who generated the document.

## How to setup
Simply download or clone the repo and run the Notify.xlsm document

1. Download by clicking Code buttong and Download zip
2. Extract Zip file to a new directory.
3. Notify.xlsm - (sample data included)


Give it a star :tada:
---------------------
Did you find this information useful, then give it a star 
