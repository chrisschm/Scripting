Imports System.Runtime.InteropServices

''' <summary>
''' Zeigt ein Standarddialogfeld an, das den Benutzer zum Speichern einer Datei auffordert.
''' </summary>
<ComClass(SaveFileDialog.ClassId, SaveFileDialog.InterfaceId, SaveFileDialog.EventId)>
Public Class SaveFileDialog

#Region "Konstante"
    ''' <summary>
    ''' Klassen-ID mit der die <see cref="SaveFileDialog"/>-Klasse eindeutig identifiziert wird.
    ''' </summary>
    Public Const ClassId As String = "2A7A3102-DF71-47A7-B3E1-9C0B60C69C17"
    ''' <summary>
    ''' Schnittstellen-ID
    ''' </summary>
    Public Const InterfaceId As String = "943BF0C3-1013-4C8D-BD17-F4A33A85031E"
    ''' <summary>
    ''' Ereignis-ID
    ''' </summary>
    Public Const EventId As String = "A3EC67D7-7AB4-451A-BA6A-3CFD8CE4D3E0"
#End Region

    Private SFD As Windows.Forms.SaveFileDialog

#Region "Auflistungen"
    ''' <summary>
    ''' Gibt als Bezeichner den Rückgabewert eines Dialogfelds an.
    ''' </summary>
    Public Enum DialogResult As Integer
        ''' <summary>
        ''' Das Dialogfeld gibt Nothing zurück.<br/>
        ''' Dies bedeutet, dass das modale Dialogfeld weiterhin ausgeführt wird.
        ''' </summary>
        None = 0
        ''' <summary>
        ''' Der Dialogfeld Rückgabewert ist Ok (in der Regel von der Schaltfläche Ok gesendet).
        ''' </summary>
        Ok = 1
        ''' <summary>
        ''' Der Dialogfeld Rückgabewert ist Cancel (in der Regel von der Schaltfläche Abbrechen gesendet).
        ''' </summary>
        Cancel = 2
        ''' <summary>
        ''' Der Dialogfeld Rückgabewert ist Abort (in der Regel von der Schaltfläche Abbrechen gesendet).
        ''' </summary>
        Abort = 3
        ''' <summary>
        ''' Der Dialogfeld Rückgabewert ist Retry (in der Regel von der Schaltfläche Wiederholen gesendet).
        ''' </summary>
        Retry = 4
        ''' <summary>
        ''' Der Dialogfeld Rückgabewert ist Ignore (in der Regel von der Schaltfläche Ignorieren gesendet).
        ''' </summary>
        Ignore = 5
        ''' <summary>
        ''' Der Dialogfeld Rückgabewert ist Yes (in der Regel von der Schaltfläche Ja gesendet).
        ''' </summary>
        Yes = 6
        ''' <summary>
        ''' Der Dialogfeld Rückgabewert ist No (in der Regel von der Schaltfläche Nein gesendet).
        ''' </summary>
        No = 7
    End Enum
#End Region

#Region "Methoden"
    ''' <summary>
    ''' Initialisiert eine Instanz der <see cref="SaveFileDialog"/>-Klasse.
    ''' </summary>
    Public Sub New()

        SFD = New Windows.Forms.SaveFileDialog
        'Standardwerte setzen, falls das Windows.Forms.OpenFileDialog sein Verhalten ändert
        SFD.AddExtension = True
        SFD.AutoUpgradeEnabled = True
        SFD.CheckFileExists = False
        SFD.CheckPathExists = True
        SFD.CreatePrompt = False
        SFD.DefaultExt = ""
        SFD.DereferenceLinks = True
        SFD.FileName = ""
        SFD.FilterIndex = 1
        SFD.InitialDirectory = ""
        SFD.OverwritePrompt = True
        SFD.RestoreDirectory = False
        SFD.ShowHelp = False
        SFD.SupportMultiDottedExtensions = False
        SFD.Title = ""
        SFD.ValidateNames = True

    End Sub

    ''' <summary>
    ''' Setzt alle Eigenschaften auf die Standardwerte zurück.
    ''' </summary>
    Public Sub Reset()

        SFD.Reset()

    End Sub

    ''' <summary>
    ''' Führt ein Standarddialogfeld mit einem Standardbesitzer aus.
    ''' </summary>
    Public Function ShowDialog() As DialogResult

        Return SFD.ShowDialog()

    End Function

    ''' <summary>
    ''' Führt ein Standarddialogfeld mit dem angegebenen Besitzer aus.
    ''' </summary>
    ''' <param name="owner">
    ''' Ein beliebiges Objekt, das <see cref="Windows.Forms.IWin32Window"/> implementiert, dass das Fenster der 
    ''' obersten Ebene und damit den Besitzer des modalen Dialogfelds darstellt.
    ''' </param>
    ''' <returns></returns>
    Public Function ShowDialog(owner As Windows.Forms.IWin32Window) As DialogResult

        Return SFD.ShowDialog(owner)

    End Function
#End Region

#Region "Eigenschaften"
    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob einem Dateinamen im Dialogfeld automatisch eine Erweiterung
    ''' hinzugefügt wird, wenn der Benutzer keine Erweiterung angibt, oder legt diesen Wert fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld einem Dateinamen eine Erweiterung hinzugefügt, wenn der Benutzer keine 
    ''' Erweiterung eingeben hat; andernfalls <c>False</c>. Der Standardwert ist <c>True</c>.
    ''' </returns>
    Public Property AddExtension() As Boolean
        Set(value As Boolean)
            SFD.AddExtension = value
        End Set
        Get
            Return SFD.AddExtension
        End Get
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob diese SaveFileDialog Instanz die Darstellung automatisch
    ''' aktualisieren soll, oder legt diesen Wert fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn diese SaveFileDialog Instanz automatisch die Darstellung aktualisieren soll,
    ''' andernfalls <c>False</c>. Der Standardwert ist <c>True</c>.
    ''' </returns>
    Public Property AutoUpgradeEnabled() As Boolean
        Set(value As Boolean)
            SFD.AutoUpgradeEnabled = value
        End Set
        Get
            Return SFD.AutoUpgradeEnabled
        End Get
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob im Dialogfeld eine Warnung angezeigt wird, wenn der 
    ''' Benutzer den Namen einer nicht vorhandenen Datei angibt, oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn im Dialogfeld bei Angabe eines nicht vorhandenen Dateinamens durch den Benutzer 
    ''' eine Warnung angezeigt wird, andernfalls <c>False</c>. Der Standardwert ist <c>True</c>.
    ''' </returns>
    Public Property CheckFileExists() As Boolean
        Set(value As Boolean)
            SFD.CheckFileExists = value
        End Set
        Get
            Return SFD.CheckFileExists
        End Get
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob im Dialogfeld eine Warnung angezeigt wird, wenn der 
    ''' Benutzer einen nicht vorhandenen Pfad angibt, oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn im Dialogfeld bei Angabe eines nicht vorhandenen Pfades durch den Benutzer 
    ''' eine Warnung angezeigt wird, andernfalls <c>False</c>. Der Standardwert ist <c>True</c>.
    ''' </returns>
    Public Property CheckPathExists() As Boolean
        Set(value As Boolean)
            SFD.CheckPathExists = value
        End Set
        Get
            Return SFD.CheckPathExists
        End Get
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob das Dialogfeld eine Bestätigung fordert, wenn der 
    ''' Benutzer eine Datei angibt, die nicht vorhanden ist, oder legt diesen Wert fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld eine Bestätigung fordert das eine Datei erstellt wird, wenn der 
    ''' Benutzer einen Dateinamen angibt, der nicht vorhanden ist, <c>False</c> wenn eine neue Datei automatisch 
    ''' erstellt wird, ohne den Benutzer zur Bestätigung aufzufordern. Der Standardwert ist <c>False</c>.
    ''' </returns>
    Public Property CreatePrompt() As Boolean
        Set(value As Boolean)
            SFD.CreatePrompt = value
        End Set
        Get
            Return SFD.CreatePrompt
        End Get
    End Property

    ''' <summary>
    ''' Ruft die Standarddateierweiterung ab oder legt diese fest.
    ''' </summary>
    ''' <returns>
    ''' Die Standarddateinamenerweiterung. Die zurückgegebene Zeichenfolge enthält keinen 
    ''' Punkt. Der Standardwert ist eine leere Zeichenfolge („“).
    ''' </returns>
    Public Property DefaultExt() As String
        Set(value As String)
            SFD.DefaultExt = value
        End Set
        Get
            Return SFD.DefaultExt
        End Get
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob das Dialogfeld den Speicherort der Datei auf die die 
    ''' Verknüpfung verweist, oder den Speicherort der Verknüpfung (.lnk) zurückgibt, oder legt diesen Wert fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld den Speicherort der Datei auf die Verknüpfung verweist zurückgibt.
    ''' andernfalls <c>False</c>. Der Standardwert ist <c>True</c>.
    ''' </returns>
    Public Property DereferenceLinks() As Boolean
        Set(value As Boolean)
            SFD.DereferenceLinks = value
        End Set
        Get
            Return SFD.DereferenceLinks
        End Get
    End Property

    ''' <summary>
    ''' Ruft eine Zeichenfolge ab die den im Dateidialogfeld ausgewählten Dateinamen enthält, oder legt diese fest.
    ''' </summary>
    ''' <returns>
    ''' Der Dateiname der im Dateidialogfeld ausgewählt wurde. Der Standardwert ist eine leere Zeichenfolge („“).
    ''' </returns>
    Public Property FileName() As String
        Set(value As String)
            SFD.FileName = value
        End Set
        Get
            Return SFD.FileName
        End Get
    End Property

    ''' <summary>
    ''' Ruft die Dateinamen aller im Dialogfeld ausgewählten Dateien ab.
    ''' </summary>
    ''' <returns>
    ''' Ein Array vom Typ <see cref="String"/>, das die Dateinamen aller ausgewählten Dateien im Dialogfeld enthält.
    ''' </returns>
    Public ReadOnly Property FileNames() As String()
        Get
            Return SFD.FileNames
        End Get
    End Property

    ''' <summary>
    ''' Ruft die aktuelle Filterzeichenfolge für Dateinamen ab, die die im Dialogfeld im Feld 
    ''' „Dateityp“ angezeigte Auswahl bestimmt, oder legt diese fest.
    ''' </summary>
    ''' <returns>
    ''' Die im Dialogfeld verfügbaren Optionen zum Filtern von Dateien.
    ''' </returns>
    ''' <exception cref="ArgumentException">Das Format ist ungültig.</exception>
    Public Property Filter() As String
        Set(value As String)
            SFD.Filter = value
        End Set
        Get
            Return SFD.Filter
        End Get
    End Property

    ''' <summary>
    ''' Ruft den Index des derzeit im Dateidialogfeld ausgewählten Filters ab oder legt diesen Index fest.
    ''' </summary>
    ''' <returns>
    ''' Ein Wert, der den Index des derzeit im Dateidialogfeld ausgewählten Filters enthält. Der Standardwert ist 1.
    ''' </returns>
    Public Property FilterIndex() As Integer
        Set(value As Integer)
            SFD.FilterIndex = value
        End Set
        Get
            Return SFD.FilterIndex
        End Get
    End Property

    ''' <summary>
    ''' Ruft das Ausgangsverzeichnis ab das im Dateidialogfeld angezeigt wird, oder legt dieses fest.
    ''' </summary>
    ''' <returns>
    ''' Das Ausgangsverzeichnis das im Dateidialogfeld angezeigt wird. Der Standardwert ist eine leere Zeichenfolge ("").
    ''' </returns>
    Public Property InitialDirectory() As String
        Set(value As String)
            SFD.InitialDirectory = value
        End Set
        Get
            Return SFD.InitialDirectory
        End Get
    End Property

    ''' <summary>
    ''' Ruft ab oder legt einen Wert fest der angibt, ob der SaveFileDialog eine Warnung angezeigt, 
    ''' wenn der Benutzer einen Dateinamen angibt, der bereits vorhanden ist.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld eine Warnung anzeigt bevor eine vorhandene Datei überschrieben wird, 
    ''' wenn der Benutzer einen Dateinamen angibt der bereits vorhanden ist, <c>False</c> wenn die vorhandene 
    ''' Datei automatisch überschrieben wird ohne den Benutzer zur Bestätigung aufzufordern. Der Standardwert ist <c>True</c>.
    ''' </returns>
    Public Property OverwritePrompt() As Boolean
        Set(value As Boolean)
            SFD.OverwritePrompt = value
        End Set
        Get
            Return SFD.OverwritePrompt
        End Get
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab oder legt diesen fest, der angibt ob das Dialogfeld das Verzeichnis 
    ''' im zuvor ausgewählten Verzeichnis vor dem Schließen wiederherstellt.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld das aktuelle Verzeichnis im zuvor ausgewählten Verzeichnis wiederherstellt, wenn der Benutzer 
    ''' bei der Suche nach Dateien das Verzeichnis gewechselt hat, andernfalls <c>False</c>. Der Standardwert ist <c>False</c>.
    ''' </returns>
    Public Property RestoreDirectory() As Boolean
        Set(value As Boolean)
            SFD.RestoreDirectory = value
        End Set
        Get
            Return SFD.RestoreDirectory
        End Get
    End Property

    ''' <summary>
    ''' Ruft ab oder legt fest, ob das Dialogfeld Anzeige und Speichern von Dateien mehrere Dateinamenerweiterungen unterstützt.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld mehrere Dateinamenerweiterungen unterstützt. andernfalls <c>False</c>. Die Standardeinstellung ist <c>False</c>.
    ''' </returns>
    Public Property SupportMultiDottedExtensions() As Boolean
        Set(value As Boolean)
            SFD.SupportMultiDottedExtensions = value
        End Set
        Get
            Return SFD.SupportMultiDottedExtensions
        End Get
    End Property

    ''' <summary>
    ''' Ruft den Titel des Dateidialogfelds ab oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' Der Titel des Dateidialogfelds. Der Standardwert ist eine leere Zeichenfolge („“).
    ''' </returns>
    Public Property Title() As String
        Set(value As String)
            SFD.Title = value
        End Set
        Get
            Return SFD.Title
        End Get
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob das Dialogfeld nur gültige Win32-Dateinamen akzeptiert, oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld nur gültige Win32-Dateinamen akzeptiert, andernfalls <c>False</c>. Der Standardwert ist <c>True</c>.
    ''' </returns>
    Public Property ValidateNames() As Boolean
        Set(value As Boolean)
            SFD.ValidateNames = value
        End Set
        Get
            Return SFD.ValidateNames
        End Get
    End Property
#End Region

End Class
