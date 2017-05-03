@echo off
mkdir "C:\Sherby Interface"
echo.>"C:\Sherby Interface\bookmarkNameStorage.txt"
echo.>"C:\Sherby Interface\bookmarkUrlStorage.txt"
break>"C:\Sherby Interface\names.txt"
break>"C:\Sherby Interface\voicerecognition.txt"
echo.>"C:\Sherby Interface\toDoGoals.txt"
echo.>"C:\Sherby Interface\xmlWolframFile.txt"
echo.>"C:\Sherby Interface\weather.txt"
echo.>"C:\Sherby Interface\notesStorage.txt"

START CMD /C "ECHO The file has finished creating the Sherby Interface folder, and will now delete itself. (For your convenience) && PAUSE"
start /b "" cmd /c del "%~f0"&exit /b