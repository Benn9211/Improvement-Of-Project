Option Explicit

Dim IE
Dim shell
Set shell = WScript.CreateObject("WScript.Shell")
Set IE = WScript.CreateObject("InternetExplorer.Application")
IE.Visible = True
shell.AppActivate IE
IE.Navigate "https://kenwood.garmin.com/kenwood/site/viewHeadUnit?headUnitInfoId=12170&updateMediaTypeId=13&modelYear=2015"