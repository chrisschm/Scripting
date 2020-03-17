Imports System.Runtime.InteropServices

''' <summary>
''' Stellt ein Standarddialogfeld dar in dem die verfügbaren Farben angezeigt werden, sowie 
''' Steuerelemente mit denen Benutzer benutzerdefinierte Farben definieren können.
''' </summary>
<ComClass(ColorDialog.ClassId, ColorDialog.InterfaceId, ColorDialog.EventId)>
Public Class ColorDialog

#Region "Konstante"
    ''' <summary>
    ''' Klassen-ID mit der die <see cref="ColorDialog"/>-Klasse eindeutig identifiziert wird.
    ''' </summary>
    Public Const ClassId As String = "B8B5466F-EA41-4E45-8E20-E2238282EA48"
    ''' <summary>
    ''' Schnittstellen-ID
    ''' </summary>
    Public Const InterfaceId As String = "1E60E717-62BF-4960-BBEF-CA01C101BCA7"
    ''' <summary>
    ''' Ereignis-ID
    ''' </summary>
    Public Const EventId As String = "F68CB41A-D02E-411C-B0B8-5AB50BF61DDC"
#End Region

    Private CD As Windows.Forms.ColorDialog

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
    ''' Initialisiert eine Instanz der <see cref="ColorDialog"/>-Klasse.
    ''' </summary>
    Public Sub New()

        CD = New Windows.Forms.ColorDialog
        'Standardwerte setzen, falls das Windows.Forms.ColorDialog sein Verhalten ändert
        CD.AllowFullOpen = True
        CD.AnyColor = False
        CD.Color = Drawing.Color.Black
        CD.FullOpen = False
        CD.SolidColorOnly = False

    End Sub

    ''' <summary>
    ''' Setzt alle Optionen auf die Standardwerte zurück, die zuletzt ausgewählte Farbe auf Schwarz und die benutzerdefinierten Farben auf die Standardwerte.
    ''' </summary>
    Public Sub reset()

        CD.Reset()
        'Standardwerte setzen, falls das Windows.Forms.ColorDialog sein Verhalten ändert
        CD.AllowFullOpen = True
        CD.AnyColor = False
        CD.Color = Drawing.Color.Black
        CD.FullOpen = False
        CD.SolidColorOnly = False

    End Sub

    ''' <summary>
    ''' Führt ein Standarddialogfeld mit einem Standardbesitzer aus.
    ''' </summary>
    Public Function ShowDialog() As DialogResult

        Return CD.ShowDialog()

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

        Return CD.ShowDialog(owner)

    End Function
#End Region

#Region "Eigenschaften"

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob im Dialogfeld benutzerdefinierte Farben definiert werden können, oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn der Benutzer benutzerdefinierte Farben definieren kann, andernfalls <c>False</c>. Die Standardeinstellung ist <c>True</c>.
    ''' </returns>
    Public Property AllowFullOpen() As Boolean
        Get
            Return CD.AllowFullOpen
        End Get
        Set(value As Boolean)
            CD.AllowFullOpen = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob im Dialogfeld bei den Grundfarben alle verfügbaren Farben angezeigt werden, oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn das Dialogfeld bei den Grundfarben alle verfügbare Farben angezeigt, andernfalls <c>False</c>. Der Standardwert ist <c>False</c>.
    ''' </returns>
    Public Property AnyColor() As Boolean
        Get
            Return CD.AnyColor
        End Get
        Set(value As Boolean)
            CD.AnyColor = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft den Wert der Alphakomponente der ausgewählten Farbe ab.
    ''' </summary>
    ''' <returns>
    ''' Der Wert der Alphakomponente der ausgewählten Farbe.
    ''' </returns>
    Public ReadOnly Property ColorAlpha() As Byte
        Get
            Return CD.Color.A
        End Get
    End Property

    ''' <summary>
    ''' Ruft den 32-Bit-ARGB-Wert der ausgewählten Farbe ab.
    ''' </summary>
    ''' <returns>
    ''' Die 32-Bit-ARGB-Wert der ausgewählten Farbe.
    ''' </returns>
    Public ReadOnly Property ColorArgb() As Integer
        Get
            Return CD.Color.ToArgb()
        End Get
    End Property

    ''' <summary>
    ''' Ruft den Wert des Blauanteils der ausgewählten Farbe ab.
    ''' </summary>
    ''' <returns>
    ''' Der Wert des Blauanteils der ausgewählten Farbe.
    ''' </returns>
    Public ReadOnly Property ColorBlue() As Byte
        Get
            Return CD.Color.B
        End Get
    End Property

    ''' <summary>
    ''' Ruft den Wert des Grünanteils der ausgewählten Farbe ab.
    ''' </summary>
    ''' <returns>
    ''' Der Wert des Grünanteils der ausgewählten Farbe.
    ''' </returns>
    Public ReadOnly Property ColorGreen() As Byte
        Get
            Return CD.Color.G
        End Get
    End Property

    ''' <summary>
    ''' Ruft die hexadezimale Darstellung der ausgewählten Farbe ab.
    ''' </summary>
    ''' <returns>
    ''' Die hexadezimale Darstellung der ausgewählten Farbe.
    ''' </returns>
    Public ReadOnly Property ColorHTML() As String
        Get
            Return "#" & Hex(CD.Color.R) & Hex(CD.Color.G) & Hex(CD.Color.B)
        End Get
    End Property

    ''' <summary>
    ''' Ruft den Namen der ausgewählten Farbe ab.
    ''' </summary>
    ''' <returns>
    ''' Der Name der ausgewählten Farbe.
    ''' </returns>
    Public ReadOnly Property ColorName() As String
        Get
            Return CD.Color.Name
        End Get
    End Property

    ''' <summary>
    ''' Ruft den Wert des Rotanteils der ausgewählten Farbe ab.
    ''' </summary>
    ''' <returns>
    ''' Der Wert des Rotanteils der ausgewählten Farbe.
    ''' </returns>
    Public ReadOnly Property ColorRed() As Byte
        Get
            Return CD.Color.R
        End Get
    End Property

    ''' <summary>
    ''' Ruft den im Dialogfeld angezeigten Satz benutzerdefinierter Farben ab oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' Ein Satz benutzerdefinierter Farben, die im Dialogfeld angezeigt werden. Der Standardwert ist <c>Null</c>.
    ''' </returns>
    Public Property CustomColors() As Integer()
        Get
            If CD.CustomColors Is Nothing Then
                Return Nothing
            Else
                Return CD.CustomColors
            End If
        End Get
        Set(value As Integer())
            CD.CustomColors = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob die Steuerelemente für das Erstellen benutzerdefinierter Farben 
    ''' beim Öffnen des Dialogfelds angezeigt werden, oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn die Steuerelemente für benutzerdefinierte Farben beim Öffnen des Dialogfelds verfügbar sind, andernfalls <c>False</c>. Der Standardwert ist <c>False</c>.
    ''' </returns>
    Public Property FullOpen() As Boolean
        Get
            Return CD.FullOpen
        End Get
        Set(value As Boolean)
            CD.FullOpen = value
        End Set
    End Property

    ''' <summary>
    ''' Ruft einen Wert ab der angibt ob Benutzer im Dialogfeld ausschließlich Volltonfarben auswählen können, oder legt diesen fest.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn Benutzer nur Volltonfarben auswählen können, andernfalls <c>False</c>. Der Standardwert ist <c>False</c>.
    ''' </returns>
    Public Property SolidColorOnly() As Boolean
        Get
            Return CD.SolidColorOnly
        End Get
        Set(value As Boolean)
            CD.SolidColorOnly = value
        End Set
    End Property
#End Region

End Class
