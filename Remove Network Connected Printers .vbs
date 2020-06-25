' Program: Remove Network Connected Printers 
' Description: VBScript script that removes all network connected printers on a local machine 
' Author: Austin Way
' Date: 6/25/2020

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colInstalledPrinters = objWMIService.ExecQuery _
    ("Select * From Win32_Printer Where Network = True")

For Each objPrinter in colInstalledPrinters
    objPrinter.Delete_
Next