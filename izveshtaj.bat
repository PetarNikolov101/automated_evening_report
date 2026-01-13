cd /d C:\skripta_neotstraneti\skripta_neotstraneti

@echo off

echo Downloading File...
call python C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti\DownloadFile.py

echo Creating Report...
call python C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti\ExcelAutomation.py 

echo Sending Report...
call python C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti\SendMail.py

pause