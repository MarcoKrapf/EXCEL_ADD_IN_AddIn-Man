Attribute VB_Name = "modAddInMan"
Option Explicit
Option Private Module

'AddIn-Man
'---------
'Version 1.0 (30.06.2017) - Windows 2010-2016
'Autor: Marco Krapf - excel@marco-krapf.de - https://marco-krapf.de/excel
'
'Das Excel-Add-in (.xlam-Datei) fügt am rechten Rand der
'Registerkarte "Start" in das Excel-Menüband eine Gruppe "AddIn-Man"
'mit zwei Buttons ein, mit denen der Standardordner für Office Add-ins
'im Windows-Explorer geöffnet wird (zum Ablegen von neuen Add-ins)
'bzw. der Excel Add-in-Manager angezeigt wird (zum Aktivieren/Deaktivieren von Add-ins).

'Windows-Standardordner für Office Add-ins im Windows-Explorer öffnen
Sub OfficeAddInFolder_show(control As IRibbonControl)
    On Error Resume Next 'z.B. wenn anderes Betriebssystem als Windows
    Shell "explorer.exe " & Application.UserLibraryPath, vbNormalFocus
End Sub

'Excel Add-in-Manager anzeigen
Sub AddInManager_show(control As IRibbonControl)
    On Error Resume Next 'eventuellen Fehler ignorieren
    Application.Dialogs(xlDialogAddinManager).Show
End Sub
