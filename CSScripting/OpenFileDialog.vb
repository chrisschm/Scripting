Imports System.Runtime.InteropServices

''' <summary>
''' Zeigt ein Standarddialogfeld an, das den Benutzer zum Öffnen einer Datei auffordert.
''' </summary>
<ComClass(OpenFileDialog.ClassId, OpenFileDialog.InterfaceId, OpenFileDialog.EventId)>
Public Class OpenFileDialog

#Region "Konstante"
    ''' <summary>
    ''' Klassen-ID mit der die <see cref="OpenFileDialog"/>-Klasse eindeutig identifiziert wird.
    ''' </summary>
    Public Const ClassId As String = "F3B24688-D0B2-4B5D-86B2-7F46E6B40117"
    ''' <summary>
    ''' Schnittstellen-ID
    ''' </summary>
    Public Const InterfaceId As String = "FAF38C68-E41F-414C-AF12-6594AD2421A1"
    ''' <summary>
    ''' Ereignis-ID
    ''' </summary>
    Public Const EventId As String = "2133DC5F-A3C3-4D77-B590-CA9FBD97DE57"
#End Region

    Private OFD As Windows.Forms.OpenFileDialog

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
    ''' Initialisiert eine Instanz der <see cref="OpenFileDialog"/>-Klasse.
    ''' </summary>
    Public Sub New()

        OFD = New Windows.Forms.OpenFileDialog
        'Standardwerte setzen, falls das Windows.Forms.OpenFileDialog sein Verhalten ändert
        OFD.AddExtension = True
        OFD.AutoUpgradeEnabled = True
        OFD.CheckFileExists = False
        OFD.CheckPathExists = True
        OFD.DefaultExt = ""
        OFD.DereferenceLinks = True
        OFD.FileName = ""
        OFD.FilterIndex = 1
        OFD.InitialDirectory = ""
        OFD.Multiselect = False
        OFD.ReadOnlyChecked = False
        OFD.RestoreDirectory = False
        OFD.ShowHelp = False
        OFD.ShowReadOnly = False
        OFD.SupportMultiDottedExtensions = False
        OFD.Title = ""
        OFD.ValidateNames = True

    End Sub

    ''' <summary>
    ''' Setzt alle Eigenschaften auf die Standardwerte zurück.
    ''' </summary>
    Public Sub Reset()

        OFD.Reset()

    End Sub

    ''' <summary>
    ''' Führt ein Standarddialogfeld mit einem Standardbesitzer aus.
    ''' </summary>
    ''' <returns>
    ''' <see cref="DialogResult.OK"/> Wenn der Benutzer im Dialogfeld auf OK klickt, andernfalls <see cref="DialogResult.Cancel"/>.
    ''' </returns>
    Public Function ShowDialog() As DialogResult

        Return OFD.ShowDialog()

    End Function

    ''' <summary>
    ''' Führt ein Standarddialogfeld mit dem angegebenen Besitzer aus.
    ''' </summary>
    ''' <param name="owner">
    ''' Ein beliebiges Objekt, das <see cref="Windows.Forms.IWin32Window"/> implementiert, dass das Fenster der 
    ''' obersten Ebene und damit den Besitzer des modalen Dialogfelds darstellt.
    ''' </param>
    ''' <returns>
    ''' <see cref="DialogResult.OK"/> Wenn der Benutzer im Dialogfeld auf OK klickt, andernfalls <see cref="DialogResult.Cancel"/>.
    ''' </returns>
    Public Function ShowDialog(owner As Windows.Forms.IWin32Window) As DialogResult

        Return OFD.ShowDialog(owner)

    End Function
#End Region

#Region "Eigenschaften"
    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob einem Dateinamen im Dialogfeld automatisch eine Erweiterung 
    ''' hinzugefügt wird wenn der Benutzer keine Erweiterung angibt, oder legt diesen Wert fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld einem Dateinamen eine Erweiterung hinzugefügt wenn der Benutzer 
    ''' keine Erweiterung angibt, andernfalls <c>False</c>. Der Standardwert ist <c>True</c>.</returns>
    Public Property AddExtension() As Boolean
        Get
            Return OFD.AddExtension
        End Get
        Set(value As Boolean)
            OFD.AddExtension = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab oder legt ihn fest, der angibt ob diese <see cref="OpenFileDialog"/> Instanz 
    ''' die Darstellung und das Verhalten automatisch aktualisieren soll bei der Ausführung unter Windows Vista.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn diese <see cref="OpenFileDialog"/> Instanz die Darstellung und das Verhalten automatisch 
    ''' aktualisieren soll bei der Ausführung unter Windows Vista, andernfalls <c>False</c>. Die Standardeinstellung ist <c>True</c>.
    ''' </returns>
    Public Property AutoUpgradeEnabled() As Boolean
        Get
            Return OFD.AutoUpgradeEnabled
        End Get
        Set(value As Boolean)
            OFD.AutoUpgradeEnabled = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob im Dialogfeld eine Warnung angezeigt wird, wenn der Benutzer den 
    ''' Namen einer nicht vorhandenen Datei angibt, oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn im Dialogfeld eine Warnung angezeigt wird wenn der Benutzer einen Dateinamen 
    ''' angibt der nicht vorhanden ist, andernfalls <c>False</c>. Der Standardwert ist <c>False</c>.
    ''' </returns>
    Public Property CheckFileExists() As Boolean
        Get
            Return OFD.CheckFileExists
        End Get
        Set(value As Boolean)
            OFD.CheckFileExists = value
        End Set
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
        Get
            Return OFD.CheckPathExists
        End Get
        Set(value As Boolean)
            OFD.CheckPathExists = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft die Standarddateierweiterung ab oder legt diese fest.
    ''' </summary>
    ''' <returns>
    ''' Die Standarddateinamenerweiterung. Die zurückgegebene Zeichenfolge enthält keinen 
    ''' Punkt. Der Standardwert ist eine leere Zeichenfolge („“).
    ''' </returns>
    Public Property DefaultExt() As String
        Get
            Return OFD.DefaultExt
        End Get
        Set(value As String)
            OFD.DefaultExt = value
        End Set
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
        Get
            Return OFD.DereferenceLinks
        End Get
        Set(value As Boolean)
            OFD.DereferenceLinks = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft eine Zeichenfolge ab die den im Dateidialogfeld ausgewählten Dateinamen enthält, oder legt diese fest.
    ''' </summary>
    ''' <returns>
    ''' Der Dateiname der im Dateidialogfeld ausgewählt wurde. Der Standardwert ist eine leere Zeichenfolge („“).
    ''' </returns>
    Public Property FileName() As String
        Get
            Return OFD.FileName
        End Get
        Set(value As String)
            OFD.FileName = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft die Dateinamen aller im Dialogfeld ausgewählten Dateien ab.
    ''' </summary>
    ''' <returns>
    ''' Ein Array vom Typ <see cref="String"/>, das die Dateinamen aller ausgewählten Dateien im Dialogfeld enthält.
    ''' </returns>
    Public ReadOnly Property FileNames() As String()
        Get
            Return OFD.FileNames
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
        Get
            Return OFD.Filter
        End Get
        Set(value As String)
            OFD.Filter = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft den Index des derzeit im Dateidialogfeld ausgewählten Filters ab oder legt diesen Index fest.
    ''' </summary>
    ''' <returns>
    ''' Ein Wert, der den Index des derzeit im Dateidialogfeld ausgewählten Filters enthält. Der Standardwert ist 1.
    ''' </returns>
    Public Property FilterIndex() As Integer
        Get
            Return OFD.FilterIndex
        End Get
        Set(value As Integer)
            OFD.FilterIndex = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft das Ausgangsverzeichnis ab das im Dateidialogfeld angezeigt wird, oder legt dieses fest.
    ''' </summary>
    ''' <returns>
    ''' Das Ausgangsverzeichnis das im Dateidialogfeld angezeigt wird. Der Standardwert ist eine leere Zeichenfolge ("").
    ''' </returns>
    Public Property InitialDirectory() As String
        Get
            Return OFD.InitialDirectory
        End Get
        Set(value As String)
            OFD.InitialDirectory = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob im Dialogfeld mehrere Dateien ausgewählt werden können, oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn im Dialogfeld mehrere Dateien zusammen oder gleichzeitig ausgewählt werden können, 
    ''' andernfalls <c>False</c>. Der Standardwert ist <c>False</c>.
    ''' </returns>
    Public Property Multiselect() As Boolean
        Get
            Return OFD.Multiselect
        End Get
        Set(value As Boolean)
            OFD.Multiselect = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob das Kontrollkästchen für den Schreibschutz aktiviert ist, oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Kontrollkästchen für den Schreibschutz aktiviert ist, andernfalls <c>False</c>. Der Standardwert ist <c>False</c>.
    ''' </returns>
    Public Property ReadOnlyChecked() As Boolean
        Get
            Return OFD.ReadOnlyChecked
        End Get
        Set(value As Boolean)
            OFD.ReadOnlyChecked = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab oder legt diesen fest, der angibt ob das Dialogfeld das Verzeichnis 
    ''' im zuvor ausgewählten Verzeichnis vor dem Schließen wiederherstellt.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld das aktuelle Verzeichnis im zuvor ausgewählten Verzeichnis wiederherstellt, 
    ''' wenn der Benutzer bei der Suche nach Dateien das Verzeichnis gewechselt hat, andernfalls <c>False</c>. Der Standardwert ist <c>False</c>.
    ''' </returns>
    Public Property RestoreDirectory() As Boolean
        Get
            Return OFD.RestoreDirectory
        End Get
        Set(value As Boolean)
            OFD.RestoreDirectory = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft den Dateinamen und die Erweiterung für die im Dialogfeld ausgewählte Datei ab. Der Dateiname enthält keine Pfadangabe.
    ''' </summary>
    ''' <returns>
    ''' Der Dateiname und die Erweiterung für die im Dialogfeld ausgewählte Datei. Der Dateiname enthält keine Pfadangabe. 
    ''' Der Standardwert ist eine leere Zeichenfolge („“).
    ''' </returns>
    Public ReadOnly Property SafeFileName() As String
        Get
            Return OFD.SafeFileName
        End Get
    End Property

    ''' <summary>
    ''' Ruft ein Array von Dateinamen und Erweiterungen für alle ausgewählten Dateien im Dialogfeld ab. 
    ''' Die Dateinamen enthalten keine Pfadangaben.
    ''' </summary>
    ''' <returns>
    ''' Ein Array von Dateinamen und Erweiterungen für alle ausgewählten Dateien im Dialogfeld. Die Dateinamen enthalten 
    ''' keine Pfadangaben. Wenn keine Dateien ausgewählt sind, wird ein leeres Array zurückgegeben.
    ''' </returns>
    Public ReadOnly Property SafeFileNames() As String()
        Get
            Return OFD.SafeFileNames
        End Get
    End Property

    ''' <summary>
    ''' Ruft ab oder legt einen Wert fest der angibt ob die Hilfe-Schaltfläche im Dateidialogfeld angezeigt wird.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld die Hilfeschaltfläche enthält, andernfalls <c>False</c>. Der Standardwert ist <c>False</c>.
    ''' </returns>
    Public Property ShowReadOnly() As Boolean
        Get
            Return OFD.ShowReadOnly
        End Get
        Set(value As Boolean)
            OFD.ShowReadOnly = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft ab oder legt fest, ob das Dialogfeld Anzeige und Öffnen von Dateien mehrere Dateinamenerweiterungen unterstützt.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld mehrere Dateinamenerweiterungen unterstützt. andernfalls <c>False</c>. Die Standardeinstellung ist <c>False</c>.
    ''' </returns>
    Public Property SupportMultiDottedExtensions() As Boolean
        Get
            Return OFD.SupportMultiDottedExtensions
        End Get
        Set(value As Boolean)
            OFD.SupportMultiDottedExtensions = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft den Titel des Dateidialogfelds ab oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' Der Titel des Dateidialogfelds. Der Standardwert ist eine leere Zeichenfolge („“).
    ''' </returns>
    Public Property Title() As String
        Get
            Return OFD.Title
        End Get
        Set(value As String)
            OFD.Title = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob das Dialogfeld nur gültige Win32-Dateinamen akzeptiert, oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld nur gültige Win32-Dateinamen akzeptiert, andernfalls <c>False</c>. Der Standardwert ist <c>True</c>.
    ''' </returns>
    Public Property ValidateNames() As Boolean
        Get
            Return OFD.ValidateNames
        End Get
        Set(value As Boolean)
            OFD.ValidateNames = value
        End Set
    End Property
#End Region

End Class
