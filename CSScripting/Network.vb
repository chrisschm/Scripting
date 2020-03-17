Imports System.Runtime.InteropServices

''' <summary>
''' Stellt eine Eigenschaft und Methoden für die Interaktion mit dem Netzwerk bereit, mit dem der Computer verbunden ist.
''' </summary>
<ComClass(Network.ClassId, Network.InterfaceId, Network.EventId)>
Public Class Network

#Region "Konstante"
    ''' <summary>
    ''' Klassen-ID mit der die <see cref="Network"/>-Klasse eindeutig identifiziert wird.
    ''' </summary>
    Public Const ClassId As String = "FB70ED75-C783-4BAC-A013-7202900EB978"
    ''' <summary>
    ''' Schnittstellen-ID
    ''' </summary>
    Public Const InterfaceId As String = "84477E67-5C7E-4C76-B43B-C746F21D2256"
    ''' <summary>
    ''' Ereignis-ID
    ''' </summary>
    Public Const EventId As String = "78076B65-9929-4CBA-8676-607377AF61D0"
#End Region

    Private NW As Devices.Network

#Region "Auflistungen"

    ''' <summary>
    ''' Gibt an ob eine Ausnahme ausgelöst wird wenn der Benutzer Abbrechen während eines Vorgangs klickt.
    ''' </summary>
    Public Enum UICancelOption As Integer
        ''' <summary>
        ''' Keine Reaktion wenn Abbrechen gedrückt wird.
        ''' </summary>
        DoNothing = 2
        ''' <summary>
        ''' Löst eine Ausnahme aus wenn Abbrechen geklickt wird.
        ''' </summary>
        ThrowException = 3
    End Enum
#End Region

#Region "Methoden"

    ''' <summary>
    ''' Lädt die angegebene Remotedatei herunter und speichert sie in der angegebenen Position.
    ''' </summary>
    ''' <param name="address">Der Pfad zu der Datei die heruntergeladen werden soll, einschließlich des Dateinamens und der Hostadresse.</param>
    ''' <param name="destinationFileName">Dateiname und Pfad unter der die heruntergeladenen Datei abgelegt werden soll.</param>
    ''' <exception cref="ArgumentException">Der Name des Laufwerks ist ungültig.</exception>
    ''' <exception cref="ArgumentException">Der <c>destinationFileName</c> endet mit einem nachgestellten Schrägstrich.</exception>
    ''' <exception cref="TimeoutException">Der Server antwortet nicht innerhalb des Standardtimeouts (100 Sekunden).</exception>
    ''' <exception cref="Security.SecurityException">Die Zielwebsite erfordert Benutzeranmeldeinformationen.</exception>
    ''' <exception cref="Security.SecurityException">Dem Benutzer fehlen die erforderlichen Berechtigungen zum Ausführen.</exception>
    ''' <exception cref="Net.WebException">Der Web-Zielserver wird die Anforderung abgelehnt.</exception>
    Public Sub DownloadFile(address As String, destinationFileName As String)

        NW.DownloadFile(address, destinationFileName)

    End Sub

    ''' <summary>
    ''' Lädt die angegebene Remotedatei herunter und speichert sie in der angegebenen Position.
    ''' </summary>
    ''' <param name="address">Der Pfad zu der Datei die heruntergeladen werden soll, einschließlich des Dateinamens und der Hostadresse.</param>
    ''' <param name="destinationFileName">Dateiname und Pfad unter der die heruntergeladenen Datei abgelegt werden soll.</param>
    ''' <param name="userName">Zu authentifizierender Benutzername.</param>
    ''' <param name="password">Das zu authentifizierende Kennwort.</param>
    ''' <exception cref="ArgumentException">Der Name des Laufwerks ist ungültig.</exception>
    ''' <exception cref="ArgumentException">Der <c>destinationFileName</c> endet mit einem nachgestellten Schrägstrich.</exception>
    ''' <exception cref="TimeoutException">Der Server antwortet nicht innerhalb des Standardtimeouts (100 Sekunden).</exception>
    ''' <exception cref="Security.SecurityException">Die Benutzerauthentifizierung schlägt fehl.</exception>
    ''' <exception cref="Security.SecurityException">Dem Benutzer fehlen die erforderlichen Berechtigungen zum Ausführen.</exception>
    ''' <exception cref="Net.WebException">Der Web-Zielserver wird die Anforderung abgelehnt.</exception>
    Public Sub DownloadFile(address As String, destinationFileName As String, userName As String, password As String)

        NW.DownloadFile(address, destinationFileName, userName, password)

    End Sub

    ''' <summary>
    ''' Lädt die angegebene Remotedatei herunter und speichert sie in der angegebenen Position.
    ''' </summary>
    ''' <param name="address">Der Pfad zu der Datei die heruntergeladen werden soll, einschließlich des Dateinamens und der Hostadresse.</param>
    ''' <param name="destinationFileName">Dateiname und Pfad unter der die heruntergeladenen Datei abgelegt werden soll.</param>
    ''' <param name="userName">Zu authentifizierender Benutzername.</param>
    ''' <param name="password">Das zu authentifizierende Kennwort.</param>
    ''' <param name="showUI"><c>True</c> um den Fortschritt des Vorgangs anzuzeigen, andernfalls <c>False</c>.</param>
    ''' <param name="connectionTimeout">Timeoutintervall in Millisekunden.</param>
    ''' <param name="overwrite"><c>True</c> wenn eine vorhandene Datei überschrieben werden soll, andernfalls <c>False</c>.</param>
    ''' <exception cref="ArgumentException">Der Name des Laufwerks ist ungültig.</exception>
    ''' <exception cref="ArgumentException">Der <c>destinationFileName</c> endet mit einem nachgestellten Schrägstrich.</exception>
    ''' <exception cref="IO.IOException">Wenn der <c>overwrite</c> Wert <c>False</c> und die Zieldatei bereits vorhanden ist.</exception>
    ''' <exception cref="TimeoutException">Der Server antwortet nicht innerhalb des angegebenen <c>connectionTimeout</c>.</exception>
    ''' <exception cref="Security.SecurityException">Die Benutzerauthentifizierung schlägt fehl.</exception>
    ''' <exception cref="Security.SecurityException">Dem Benutzer fehlen die erforderlichen Berechtigungen zum Ausführen.</exception>
    ''' <exception cref="Net.WebException">Der Web-Zielserver wird die Anforderung abgelehnt.</exception>
    Public Sub DownloadFile(address As String, destinationFileName As String, userName As String, password As String, showUI As Boolean, connectionTimeout As Integer, overwrite As Boolean)

        NW.DownloadFile(address, destinationFileName, userName, password, showUI, connectionTimeout, overwrite)

    End Sub

    ''' <summary>
    ''' Lädt die angegebene Remotedatei herunter und speichert sie in der angegebenen Position.
    ''' </summary>
    ''' <param name="address">Der Pfad zu der Datei die heruntergeladen werden soll, einschließlich des Dateinamens und der Hostadresse.</param>
    ''' <param name="destinationFileName">Dateiname und Pfad unter der die heruntergeladenen Datei abgelegt werden soll.</param>
    ''' <param name="userName">Zu authentifizierender Benutzername.</param>
    ''' <param name="password">Das zu authentifizierende Kennwort.</param>
    ''' <param name="showUI"><c>True</c> um den Fortschritt des Vorgangs anzuzeigen, andernfalls <c>False</c>.</param>
    ''' <param name="connectionTimeout">Timeoutintervall in Millisekunden.</param>
    ''' <param name="overwrite"><c>True</c> wenn eine vorhandene Datei überschrieben werden soll, andernfalls <c>False</c>.</param>
    ''' <param name="onUserCancel">Gibt an ob eine Ausnahme ausgelöst werden soll wenn der Download abgebrochen wird.</param>
    ''' <exception cref="ArgumentException">Der Name des Laufwerks ist ungültig.</exception>
    ''' <exception cref="ArgumentException">Der <c>destinationFileName</c> endet mit einem nachgestellten Schrägstrich.</exception>
    ''' <exception cref="IO.IOException">Wenn der <c>overwrite</c> Wert <c>False</c> und die Zieldatei bereits vorhanden ist.</exception>
    ''' <exception cref="TimeoutException">Der Server antwortet nicht innerhalb des angegebenen <c>connectionTimeout</c>.</exception>
    ''' <exception cref="Security.SecurityException">Die Benutzerauthentifizierung schlägt fehl.</exception>
    ''' <exception cref="Security.SecurityException">Dem Benutzer fehlen die erforderlichen Berechtigungen zum Ausführen.</exception>
    ''' <exception cref="Net.WebException">Der Web-Zielserver wird die Anforderung abgelehnt.</exception>
    Public Sub DownloadFile(address As String, destinationFileName As String, userName As String, password As String, showUI As Boolean, connectionTimeout As Integer, overwrite As Boolean, onUserCancel As UICancelOption)

        NW.DownloadFile(address, destinationFileName, userName, password, showUI, connectionTimeout, overwrite, onUserCancel)

    End Sub

    ''' <summary>
    ''' Initialisiert eine neue Instanz der <see cref="Network"/>-Klasse.
    ''' </summary>
    Public Sub New()

        NW = New Devices.Network

    End Sub

    ''' <summary>
    ''' Pingt den angegebenen Server.
    ''' </summary>
    ''' <param name="hostNameOrAddress">Die URL, der Computername oder die IP-Adresse des Servers der gepingt werden soll.</param>
    ''' <returns>
    ''' <c>True</c> wenn der Vorgang erfolgreich war, andernfalls <c>False</c>.
    ''' </returns>
    ''' <exception cref="InvalidOperationException">Es ist keine Verbindung zum Netzwerk verfügbar.</exception>
    ''' <exception cref="Net.NetworkInformation.PingException">URL war ungültig.</exception>
    Public Function Ping(hostNameOrAddress As String) As Boolean

        Return NW.Ping(hostNameOrAddress)

    End Function

    ''' <summary>
    ''' Pingt den angegebenen Server.
    ''' </summary>
    ''' <param name="hostNameOrAddress">Die URL, der Computername oder die IP-Adresse des Servers der gepingt werden soll.</param>
    ''' <param name="timeout">Die Zeitschwelle in Millisekunden für die Kontaktaufnahme mit dem Ziel. Standard ist 500.</param>
    ''' <returns>
    ''' <c>True</c> wenn der Vorgang erfolgreich war, andernfalls <c>False</c>.
    ''' </returns>
    ''' <exception cref="InvalidOperationException">Es ist keine Verbindung zum Netzwerk verfügbar.</exception>
    ''' <exception cref="Net.NetworkInformation.PingException">URL war ungültig.</exception>
    Public Function Ping(hostNameOrAddress As String, timeout As Integer) As Boolean

        Return NW.Ping(hostNameOrAddress, timeout)

    End Function

#End Region

#Region "Eigenschaften"

    ''' <summary>
    ''' Gibt an, ob ein Computer mit einem Netzwerk verbunden ist.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> wenn der Computer mit einem Netzwerk verbunden ist, andernfalls <c>True</c>.
    ''' </returns>
    Public ReadOnly Property IsAvailable() As Boolean
        Get
            Return NW.IsAvailable
        End Get
    End Property

#End Region

End Class