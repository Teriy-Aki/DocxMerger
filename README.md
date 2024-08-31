# DocxMerger

This is a simple PowerShell script, which
1. Automatically merges two or multiple docx files into one
2. Allows you to merge files based on either filename, modification date, or creation date
3. Has a simple GUI to track the progress and report errors

All while preserving all inherent document formatting and inserting page breaks between documents.

The Installation should take ~ 5 Minutes.


# Prerequisites
This script requires Windows Interop. There are multiple ways to get it. The easiest is probably through Nuget, which you can get through the following steps:
1. Install Winget - https://apps.microsoft.com/detail/9nblggh4nns1?rtc=1&hl=de-de&gl=DE
2. Get Nuget - In a Powershell window, run: winget install Microsoft.NuGet
3. Get Windows Interop - In a Powershell window, run: install-Package Microsoft.Office.Interop.Word.


# Installation
This script is designed for the context menu (via send to).
1. Press Win + R and type "shell:sendto" (without the quotes). Press enter.
3. Drop the script in the newly opened folder
4. Open the script with Notepad and replace the path in line 39 with any path you want. The file is only temporarily stored there, so I recommend using the desktop.
5. Create a shortcut of the script in the same folder, then name it however you want. Open its properties and prepend "powershell.exe -ExecutionPolicy Bypass -File " (without the quotes) to the target path.

# Usage
Select all the docx files you want to merge in your file explorer. Right-click -> Send to -> DocxMerger

Press 1 to merge by filename, 2 to merge by creation date, and 3 to merge by modification date.

Done!

# Notes
You have to allow Windows to execute local scripts from any folder: In a Powershell window, run: set-executionpolicy remotesigned


This is my very first Powershell script. You could probably automate some installation requirements, but I still need to learn how to do that. Kindly help me improve this script by reporting bugs or issues.
