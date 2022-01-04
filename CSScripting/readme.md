# Verwendung

## VB.Net und der Export von COM Klassen
.Net Assemblies exportieren von sich aus keine Klassen als COM Klassen. Dieses muss explizit 
über InteropServices codiert werden. Zu diesem Zweck enthält jede Klasse die exportiert werden soll drei
öffentliche Konstante (ClassId, InterfaceId, und EventId) die je eine eindeutige GUID enthalten die
zusammen die Klasse representieren und mit dem Namen der Klasse exportiert werden, um eine COM Klasse
zu erstellen.

```vbs
<ComClass(FolderBrowserDialog.ClassId, FolderBrowserDialog.InterfaceId, FolderBrowserDialog.EventId)>
Public Class FolderBrowserDialog
```

Die Visual Studio IDE erstellt die dazu notwendigen Einträge ich der Registry automatisch bei Ausführung.
Soll das auf der Zielmaschine funktionieren, muss ein Setup die Registrierung der COM Klassen übernehmen.
Für die Möglichkeit regsrv32.exe dafür zu nutzen müssten weitere Methoden exportiert werden, das soll aber 
nicht Teil dieses Beispiels sein und diese Methoden sind in den Klassen nicht enthalten.

## Beispiel
Als Beispiel nehme ich die FolderBrowserDialog Klasse. Die anderen Dialog sind equivalent zu nutzen. Die 
Beschreibung jeder Methode, Konstante, Variable und Eigenschaft sind im Code kommentiert.

```vbs
Dim FBDlg As Object
dim sPath As String

Set FBDlg = CreateObject("FolderBrowserDialog")

FBDlg.Description = "Bitte das Verzeichnis wählen."
FBDlg.RootFolder = 0
FBDlg.SelectedPath = "C:\"

If FBDlg.ShowDialog() = 1 Then
  sPath = FBDlg.SelectedPath
End If
```
