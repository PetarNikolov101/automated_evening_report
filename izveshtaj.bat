@echo off
cd /d "C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti"

echo Downloading File...
call "C:\Python314\python.exe" DownloadFile.py >> log.txt 2>&1

echo Creating Report...
call "C:\Python314\python.exe" ExcelAutomation.py >> log.txt 2>&1

echo Sending Report...
call "C:\Python314\python.exe" SendMail.py >> log.txt 2>&1

echo Done.