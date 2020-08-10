ECHO OFF
rem <- this is a comment (rem for remarks)
:: <- this is also a comment 

rem below opens empty excel 
start  excel.exe 

rem below opens a specific file in default viewer (excel for .xlsx)
start "" "C:\enter path here\filename.xlsx"

rem Enter desired websites to open in default web browser 
start "" "http://bing.com"
start "" "http://google.com"  

rem below opens outlook
start outlook.exe


rem PAUSE