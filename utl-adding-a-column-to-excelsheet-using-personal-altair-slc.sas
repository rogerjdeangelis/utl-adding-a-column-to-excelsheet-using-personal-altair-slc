%let pgm=utl-adding-a-column-to-excelsheet-using-personal-altair-slc;

%stop_submission;

Adding a column to existing excel sheet using personal altair slc

github
https://tinyurl.com/ms3psd5s
https://github.com/rogerjdeangelis/utl-adding-a-column-to-excelsheet-using-personal-altair-slc

OPS Question

I have a worksheet with transactions neatly organized.
But the first four lines of the sheet have information which
is not deemed needed except for the account number which
is the last nine characters of cell A2.
How can I extract/append that number to the last column of the sheet?


Note this does not do inplace addition of a column, sas and the cannot do that.
In fact it is difficult with R and Python.

  Process
     1 load the sheet into gourlines.xlsx
     2 read fourlines.xlsx and create dataset extract
     3 delete fourlines.xlsx
     4 recreate fourlines.xlsx from extract

     FOR SOME UNKNOWN REASON THE FINAL EXCEL WORKBOOK IS NOT LISTED IN EXPLORER
     YOU HAVE TO SEARCH FOR IT, IT IS THERE

Note:
  SAS and tthe SLC can edite existing celles and columns, but not add new columns.
  Styles can be maintaines see below


community.altair
https://tinyurl.com/3xumrws2
https://community.altair.com/discussion/comment/195972?tab=all#Comment_195972?utm_source=community-search&utm_medium=organic-search&utm_term=sas

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/
s:/xls/fourlines.xlsx

 -----------------------+
 | A1| fx    |LINE      |
 ----------------------------------------+
 [_] |              A                    |
 ----------------------------------------|
  1  |  LINE                             |
  -- |-----------------------------------|
  2  |  This is the first line           |
  -- |-----------------------------------|
  3  |  This is the second line123456789 |
  -- |-----------------------------------|
  4  |  This is the third line           |
  -- |-----------------------------------|
  5  |  This is the fourth line          |
  -- |----------+---------+---------+----|


/*--- create xlsx workbook ---*/
&_init_;
data fourlines;
 input line & $44.;
cards4;
This is the first line
This is the second line123456789
This is the third line
This is the fourth line
;;;;
run;quit;

%utlfkil(d:/xls/fourlines.xlsx);

ods excel file="d:/xls/fourlines.xlsx"
  options(sheet_name="fourlines");
proc print data=fourlines;
run;quit;

ods excel close;

/*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

/*---- convert fourlines sheet to slc dataset ----*/

libname xls xlsx "d:/xls/fourlines.xlsx";
data xls.extract;
  set xls.fourlines;
  if _n_=2 then  a2=substr(line,length(line)-8,9);
  else a2=.;
run;quit;
libname xls clear;

/*---- recreate fourlines.xlsx ----*/
%utlfkil(d:/xls/fourlines.xlsx);

ods excel file="d:/xls/fourlines.xlsx"
  options(sheet_name="fourlines");
proc print data=extract;
run;quit;

ods excel close;

/*           _               _
  ___  _   _| |_ _ __  _   _| |_
 / _ \| | | | __| `_ \| | | | __|
| (_) | |_| | |_| |_) | |_| | |_
 \___/ \__,_|\__| .__/ \__,_|\__|
                |_|
*/

-----------------------+
| A1| fx    |LINE      |
---------------------------------------------------+
[_] |              A                     |    B    |
---------------------------------------------------|
 1  |  LINE                              |   A2    |
 -- |------------------------------------+---------|
 2  |  This is the first line            |         |
 -- |------------------------------------+---------|
 3  |  This is the second line123456789  |123456789|
 -- |------------------------------------+---------|
 4  |  This is the third line            |         |
 -- |------------------------------------+---------|
 5  |  This is the fourth line           |         |
 -- |----------+---------+---------+-----+---------|

/*
 _ __ ___ _ __   ___  ___
| `__/ _ \ `_ \ / _ \/ __|
| | |  __/ |_) | (_) \__ \
|_|  \___| .__/ \___/|___/
         |_|
*/

UPDATE EXCEL INPACE (but not add colums)
-----------------------------------------------------------------------------------------------------------------------------------------
https://github.com/rogerjdeangelis/utl-excel-database-schema-and-tables-as-workbooks-sheets-and-named-ranges-update-existing-workbooks
https://github.com/rogerjdeangelis/utl-in-palce-updates-to-an-existing-shared-excel-workbook
https://github.com/rogerjdeangelis/utl-ods-excel-update-excel-sheet-in-place-python
https://github.com/rogerjdeangelis/utl-r-open-closed-excel-workbook-an-update-master-sheet-using-transaction-sheet
https://github.com/rogerjdeangelis/utl-update-a-master-sheet-with-transaction-sheet-using-excel-and-r-openxls-package-and-sqldf
https://github.com/rogerjdeangelis/utl-update-an-excel-workbook-in-place
https://github.com/rogerjdeangelis/utl-update-an-existing-excel-named-range-R-python-sas
https://github.com/rogerjdeangelis/utl-update-existing-excel-sheet-in-place-using-r-dcom-client
https://github.com/rogerjdeangelis/utl-update-in-place-sheet2-by-adding-dinner-costs-from-sheet1-preserving-excel-formatting-r
https://github.com/rogerjdeangelis/utl_excel_create_sql_insert_and_value_statements_to_update_databases
https://github.com/rogerjdeangelis/utl_excel_update_inplace
https://github.com/rogerjdeangelis/utl_excel_update_rectangle
https://github.com/rogerjdeangelis/utl_excel_update_xlsm_workbook_using_SAS_dataset
Select update excel from d:/git/git_010_repos.sasbdat

SLC
-----------------------------------------------------------------------------------------------------------------------------------------
https://github.com/rogerjdeangelis/setup-personal-edition-altair-slc-eclipse-workspace-config-sasautos-sasuser-saswork-autoexec
https://github.com/rogerjdeangelis/utl-altair-slc-to-fill-gaps-in-proc-sql-select-third-place-in-the-daily-double-r-python-solutions
https://github.com/rogerjdeangelis/utl-calling-python-from-personal-altair-slc-and-integrating-python-with-sql
https://github.com/rogerjdeangelis/utl-calling-r-from-personal-altair-slc-and-integrating-r-with-sql
https://github.com/rogerjdeangelis/utl-dropping-down-to-powershell-from-personal-altair-slc
https://github.com/rogerjdeangelis/utl-how-to-create-a-sas-dataset-from-python-panda-dataframe-using-the-personal-altair-slc
https://github.com/rogerjdeangelis/utl-how-to-create-a-sas-dataset-from-r-dataframe-using-the-personal-altair-slc
/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
