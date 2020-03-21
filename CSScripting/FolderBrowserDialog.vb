Imports System.Runtime.InteropServices

''' <summary>
''' Zeigt ein FolderBrowserDialogStandarddialogfeld an, das den Benutzer zur Auswahl eines Ordners auffordert.
''' </summary>
<ComClass(FolderBrowserDialog.ClassId, FolderBrowserDialog.InterfaceId, FolderBrowserDialog.EventId)>
Public Class FolderBrowserDialog

#Region "Konstante"
    ''' <summary>
    ''' Klassen-ID mit der die <see cref="OpenFileDialog"/>-Klasse eindeutig identifiziert wird.
    ''' </summary>
    Public Const ClassId As String = "14D45182-D309-4E44-8AFE-CCE720795D9F"
    ''' <summary>
    ''' Schnittstellen-ID
    ''' </summary>
    Public Const InterfaceId As String = "DF4A7C1D-C579-4AD6-A337-958E629056D1"
    ''' <summary>
    ''' Ereignis-ID
    ''' </summary>
    Public Const EventId As String = "7D82B617-6D32-4F7C-92F1-527F4EF5EA76"
#End Region

    Private FBD As Windows.Forms.FolderBrowserDialog

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

    ''' <summary>
    ''' Gibt Enumerationskonstanten an, mit denen Verzeichnispfade für besondere Systemordner abgerufen werden.
    ''' </summary>
    Public Enum SpecialFolder As Integer
        ''' <summary>
        ''' Der logische Desktop und nicht der physische Speicherort im Dateisystem.
        ''' </summary>
        Desktop = 0
        ''' <summary>
        ''' Das Verzeichnis, das die Programmgruppen des Benutzers enthält.
        ''' </summary>
        Programs = 2
        ''' <summary>
        ''' Das Verzeichnis, das als allgemeines Repository für Dokumente verwendet wird.
        ''' </summary>
        Personal = 5
        ''' <summary>
        ''' Der Ordner Eigene Dateien.
        ''' </summary>
        MyDocuments = 5
        ''' <summary>
        ''' Das Verzeichnis, das als allgemeines Repository für die Favoriten des Benutzers verwendet wird.
        ''' </summary>
        Favorites = 6
        ''' <summary>
        ''' Das Verzeichnis, das der Programmgruppe "Autostart" des Benutzers entspricht.
        ''' </summary>
        Startup = 7
        ''' <summary>
        ''' Das Verzeichnis, das die vom Benutzer zuletzt verwendeten Dokumente enthält.
        ''' </summary>
        Recent = 8
        ''' <summary>
        ''' Das Verzeichnis, das die Elemente für das Menü "Senden an" enthält.
        ''' </summary>
        SendTo = 9
        ''' <summary>
        ''' Das Verzeichnis, das die Elemente für das Menü "Start" enthält.
        ''' </summary>
        StartMenu = 11
        ''' <summary>
        ''' Der Ordner Eigene Musik.
        ''' </summary>
        MyMusic = 13
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das als Repository für Videos dient, die zu einem Benutzer gehören.
        ''' </summary>
        MyVideos = 14
        ''' <summary>
        ''' Das Verzeichnis, das für das physische Speichern von Dateiobjekten auf dem Desktop verwendet wird.
        ''' </summary>
        DesktopDirectory = 16
        ''' <summary>
        ''' Der Ordner Arbeitsplatz.
        ''' </summary>
        MyComputer = 17
        ''' <summary>
        ''' Ein Dateisystemverzeichnis, das die Linkobjekte enthält, die im virtuellen Ordner Netzwerkumgebung vorhanden sein können.
        ''' </summary>
        NetworkShortcuts = 19
        ''' <summary>
        ''' Ein virtueller Ordner, der Schriftarten enthält.
        ''' </summary>
        Fonts = 20
        ''' <summary>
        ''' Das Verzeichnis, das als allgemeines Repository für Dokumentvorlagen verwendet wird.
        ''' </summary>
        Templates = 21
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das die Programme und Ordner enthält, die im Menü Start für alle Benutzer angezeigt werden. Dieser besondere Ordner ist nur für Windows NT-Systeme gültig.
        ''' </summary>
        CommonStartMenu = 22
        ''' <summary>
        ''' Ein Ordner für Komponenten, die von mehreren Anwendungen gemeinsam verwendet werden. Dieser besondere Ordner nur für Windows NT-, Windows 2000- und Windows XP-Systeme gültig. 
        ''' </summary>
        CommonPrograms = 23
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das die Programme enthält, die im Ordner Startup für alle Benutzer angezeigt werden. Dieser besondere Ordner ist nur für Windows NT-Systeme gültig.
        ''' </summary>
        CommonStartup = 24
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das Dateien und Ordner enthält, die auf dem Desktop für alle Benutzer angezeigt werden. Dieser besondere Ordner ist nur für Windows NT-Systeme gültig.
        ''' </summary>
        CommonDesktopDirectory = 25
        ''' <summary>
        ''' Das Verzeichnis, das als allgemeines Repository für programmspezifische Daten des aktuellen Roamingbenutzers verwendet wird.
        ''' </summary>
        ApplicationData = 26
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das die Linkobjekte enthält, die im virtuellen Ordner Drucker vorhanden sein können.
        ''' </summary>
        PrinterShortcuts = 27
        ''' <summary>
        ''' Das Verzeichnis, das als allgemeines Repository für programmspezifische Daten verwendet wird, die von einem aktuellen Benutzer verwendet werden, der kein Roamingbenutzer ist.
        ''' </summary>
        LocalApplicationData = 28
        ''' <summary>
        ''' Das Verzeichnis, das als allgemeines Repository für temporäre Internetdateien verwendet wird.
        ''' </summary>
        InternetCache = 32
        ''' <summary>
        ''' Das Verzeichnis, das als allgemeines Repository für Internetcookies verwendet wird.
        ''' </summary>
        Cookies = 33
        ''' <summary>
        ''' Das Verzeichnis, das als allgemeines Repository für die Internetverlaufselemente verwendet wird.
        ''' </summary>
        History = 34
        ''' <summary>
        ''' Das Verzeichnis, das als allgemeines Repository für programmspezifische Daten verwendet wird, die von allen Benutzern verwendet werden.
        ''' </summary>
        CommonApplicationData = 35
        ''' <summary>
        ''' Das Windows-Verzeichnis oder SYSROOT. Dies entspricht den Umgebungsvariablen %windir% oder %SYSTEMROOT%.
        ''' </summary>
        Windows = 36
        ''' <summary>
        ''' Das Verzeichnis "System".
        ''' </summary>
        System = 37
        ''' <summary>
        ''' Das Verzeichnis für Programmdateien. Auf einem nicht x86-System gibt die Übergabe von ProgramFiles an die GetFolderPath(SpecialFolder)-Methode 
        ''' den Pfad für nicht x86-Programme zurück. Um das Dateiverzeichnis für x86 Programme auf einem nicht x86-System abzurufen, verwenden Sie den SpecialFolder.ProgramFilesX86-Member.
        ''' </summary>
        ProgramFiles = 38
        ''' <summary>
        ''' Der Ordner Eigene Bilder.
        ''' </summary>
        MyPictures = 39
        ''' <summary>
        ''' Der Profilordner des Benutzers. 
        ''' </summary>
        UserProfile = 40
        ''' <summary>
        ''' Der Windows-Ordner System.
        ''' </summary>
        SystemX86 = 41
        ''' <summary>
        ''' Der x86-Ordner Programme.
        ''' </summary>
        ProgramFilesX86 = 42
        ''' <summary>
        ''' Das Verzeichnis für Komponenten, die von mehreren Anwendungen gemeinsam genutzt werden. 
        ''' </summary>
        CommonProgramFiles = 43
        ''' <summary>
        ''' Der Ordner Programme.
        ''' </summary>
        CommonProgramFilesX86 = 44
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das die für alle Benutzer verfügbaren Vorlagen enthält. Dieser besondere Ordner ist nur für Windows NT-Systeme gültig.
        ''' </summary>
        CommonTemplates = 45
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das Dokumente enthält, die von allen Benutzern gemeinsam genutzt werden. Dieser besondere Ordner ist für Windows NT-Systeme gültig.
        ''' </summary>
        CommonDocuments = 46
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das Verwaltungstools für alle Benutzer des Computers enthält. 
        ''' </summary>
        CommonAdminTools = 47
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das zum Speichern von Verwaltungstools für einen einzelnen Benutzer verwendet wird. Die Microsoft Management Console (MMC) 
        ''' speichert angepasste Konsolen in diesem Verzeichnis, das für den Benutzern von überall aus zugänglich ist.
        ''' </summary>
        AdminTools = 48
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das als Repository für Musikdateien dient, die von allen Benutzern gemeinsam genutzt werden. 
        ''' </summary>
        CommonMusic = 53
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das als Repository für Bilddateien dient, die von allen Benutzern gemeinsam genutzt werden. 
        ''' </summary>
        CommonPictures = 54
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das als Repository für Musikdateien dient, die von allen Benutzern gemeinsam genutzt werden. 
        ''' </summary>
        CommonVideos = 55
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das Ressourcendaten enthält. 
        ''' </summary>
        Resources = 56
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das lokalisierte Ressourcendaten enthält.
        ''' </summary>
        LocalizedResources = 57
        ''' <summary>
        ''' Dieser Wert wird in Windows Vista für die Abwärtskompatibilität erkannt, aber der besondere Ordner selbst wird nicht mehr verwendet.
        ''' </summary>
        CommonOemLinks = 58
        ''' <summary>
        ''' Das Dateisystemverzeichnis, das als Stagingbereich für Dateien fungiert, die auf eine CD geschrieben werden sollen. 
        ''' </summary>
        CDBurning = 59
    End Enum
#End Region

#Region "Methoden"
    ''' <summary>
    ''' Initialisiert eine Instanz der <see cref="FolderBrowserDialog"/>-Klasse.
    ''' </summary>
    Public Sub New()

        FBD = New Windows.Forms.FolderBrowserDialog

    End Sub

    ''' <summary>
    ''' Setzt alle Eigenschaften auf die Standardwerte zurück.
    ''' </summary>
    Public Sub Reset()

        FBD.Reset()

    End Sub

    ''' <summary>
    ''' Führt ein Standarddialogfeld mit einem Standardbesitzer aus.
    ''' </summary>
    ''' <returns>
    ''' <see cref="DialogResult.OK"/> Wenn der Benutzer im Dialogfeld auf OK klickt, andernfalls <see cref="DialogResult.Cancel"/>.
    ''' </returns>
    Public Function ShowDialog() As DialogResult

        Return FBD.ShowDialog()

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

        Return FBD.ShowDialog(owner)

    End Function
#End Region

#Region "Eigenschaften"

    ''' <summary>
    ''' Ruft den beschreibenden Text ab, der über dem Strukturansicht-Steuerelement im Dialogfeld angezeigt wird, oder legt ihn fest.
    ''' </summary>
    ''' <returns>
    ''' Die Beschreibung die angezeigt wird. Der Standardwert ist eine leere Zeichenfolge ("").
    ''' </returns>
    Public Property Description() As String
        Get
            Return FBD.Description
        End Get
        Set(value As String)
            FBD.Description = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft den Stammordner ab in dem das Durchsuchen gestartet wird.
    ''' </summary>
    ''' <returns>
    ''' Einer der <see cref="SpecialFolder"/>-Werte. Die Standardeinstellung ist Desktop.
    ''' </returns>
    ''' <exception cref="ComponentModel.InvalidEnumArgumentException">
    ''' Der zugewiesene Wert ist keiner der <see cref="SpecialFolder"/> Werte.
    ''' </exception>
    Public Property RootFolder() As SpecialFolder
        Get
            Return FBD.RootFolder
        End Get
        Set(value As SpecialFolder)
            FBD.RootFolder = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft ab oder legt den vom Benutzer ausgewählten Pfad fest.
    ''' </summary>
    ''' <returns>
    ''' Der Pfad des ersten im Dialogfeld ausgewählten Ordners oder der letzte Ordner, der vom Benutzer ausgewählt wurde. Der Standardwert ist eine leere Zeichenfolge ("").
    ''' </returns>
    Public Property SelectedPath() As String
        Get
            Return FBD.SelectedPath
        End Get
        Set(value As String)
            FBD.SelectedPath = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob die neuen Ordner Schaltfläche im Dialogfeld für den Browser angezeigt wird, oder legt ihn fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn die neuen Ordner Schaltfläche im Dialogfeld angezeigt wird, andernfalls <c>False</c>. Die Standardeinstellung ist <c>True</c>.
    ''' </returns>
    Public Property ShowNewFolderButton() As Boolean
        Get
            Return FBD.ShowNewFolderButton
        End Get
        Set(value As Boolean)
            FBD.ShowNewFolderButton = value
        End Set
    End Property
#End Region

End Class
