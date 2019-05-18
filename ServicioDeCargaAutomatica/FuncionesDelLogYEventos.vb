Imports ZKSoftwareAPI
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO.TextReader
Imports System.IO.TextWriter
Imports System
Imports System.IO
Imports System.Collections
Imports System.Diagnostics
Imports System.Threading.Thread
Imports System.ComponentModel
Module FuncionesDelLogYEventos
    '"Variables globales"
    Const MyUbicationOfLogs As String = "C:\LogsDeCargaDeLosRelojes"
    'Funciones
    Sub Informe(ByVal Txt2Inform As String)
        Dim HoraAhora As DateTime = Now
        Dim RutaDeArchivo As String = MyUbicationOfLogs + "\Registro" + Date2stringArchivo(HoraAhora) + ".txt"
        If Directory.Exists(MyUbicationOfLogs) Then
            If Not File.Exists(RutaDeArchivo) Then
                Try
                    Dim MyWriter As StreamWriter = New StreamWriter(RutaDeArchivo)
                    MyWriter.WriteLine("### --- Este es un informe del servicio de segundo plando de la carga de los relojes --- ###")
                    MyWriter.WriteLine("->Evento sucedido a las:" + vbTab + Date2stringSQL(Now))
                    MyWriter.WriteLine("-- Reporte ----------------------------------------------------------------------------------")
                    MyWriter.WriteLine(Txt2Inform)
                    MyWriter.WriteLine("----------------------------------------Servicio creado por Alan Fernando Santacruz Rodríguez")
                    MyWriter.Close()
                Catch MyErrorEnEscritura As System.Exception
                    InformeDeErrorAlEscribirElLog("No se ha podido realizar la escritura del Log a las " + Date2stringSQL(Now))
                End Try
            Else
                Directory.CreateDirectory("C:\LogsDeCargaDeLosRelojes")
            End If
        End If
    End Sub
    Sub InformeDeErrorAlEscribirElLog(ByVal sEvent As String)
        'Const strOrigen = "Servicio de carga del reloj checador"
        'Const strRegistro = "No se ha podido el text del log de los errores"
        'Const strEvento = "Ejemplo de un evento en el visor de sucesos de Windows"
        'Const strEquipo = "Servidor de la base de datos del reloj checador"
        'If (Not EventLog.Exists(strOrigen, strEquipo)) Then

        '    Dim objOrigenEvento As EventSourceCreationData = New EventSourceCreationData(strOrigen, strRegistro)
        'End If
        'Dim objEvento As EventLog = New EventLog(strRegistro, strEquipo, strOrigen)
        'objEvento.WriteEntry(seve)
        'objEvento.WriteEntry(strOrigen, "No se podido", EventLogEntryType.Error, 234, 6)
        Dim sSource As String
        Dim sLog As String
        'Dim sEvent As String
        Dim sMachine As String

        sSource = "Servicio de carga de fondo del reloj"
        sLog = "Application"
        sEvent = "EscrituraDelLog"
        sMachine = "."

        If Not EventLog.SourceExists(sSource, sMachine) Then
            EventLog.CreateEventSource(sSource, sLog, sMachine)
        End If

        Dim ELog As New EventLog(sLog, sMachine, sSource)
        ELog.WriteEntry(sEvent)
        ELog.WriteEntry(sEvent, EventLogEntryType.Error, 234, CType(3, Short))
    End Sub
End Module
