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

Public Class Service1
    Protected Friend MyContadorDeInicio As New Timers.Timer
    Protected Friend MyContadorDeCarga As New Timers.Timer

    Protected Friend MyClocks As List(Of ZKSoftware) = New List(Of ZKSoftware)

    Protected Friend MySqlCon As AdmSQL = New AdmSQL("192.168.0.100", "RelojChecador", "usuarios", "1234")
    Protected Const MyConsultaInicio As String = "Select * from ParametricosDeRelojes where Parametro like 'PuntoI%' order by Valor ASC"
    Protected Friend MyPointsInicio As List(Of Integer) = New List(Of Integer)

    Const MyUbicationOfLogs As String = "C:\LogsDeCargaDeLosRelojes"

    Protected Friend MyCategoriasDelError As List(Of String) = New List(Of String) From {
    "No se ha podido el text del log de los errores",
    "No se ha podido establecer la comunicación con la base de datos",
    "No se ha establecido la comunicación con un reloj checador",
    "No se han modificado los campos de la base de datos"
    }

    Protected WithEvents MyCargaOnBackGround As BackgroundWorker = New BackgroundWorker

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.

        MyContadorDeInicio = New Timers.Timer
        AddHandler MyContadorDeInicio.Elapsed, AddressOf ProcesandoElInicio
        MyContadorDeCarga = New Timers.Timer
        MyContadorDeCarga.AutoReset = True
        AddHandler MyContadorDeCarga.Elapsed, AddressOf ConexionYCargaDeRelojes
        ProcesandoElInicio()
        MyCargaOnBackGround.WorkerReportsProgress = True
        MyCargaOnBackGround.WorkerSupportsCancellation = True
    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
        MyCargaOnBackGround.CancelAsync()
        MyContadorDeInicio.Stop()
        MyContadorDeCarga.Stop()
        Me.Finalize()
    End Sub

    Protected Friend Sub ProcesandoElInicio()
        Dim ListaTemporal As List(Of String) = MySqlCon.SqlReaderDown2List(MyConsultaInicio)
        If ListaTemporal.Count > 0 Then
            MyPointsInicio = ListStr2Num(ListaTemporal)
            Dim NextMinute As Integer = FindTheNextNumber2Next(Now.Minute, MyPointsInicio)
            MyContadorDeInicio.Interval = NextMinute
        Else
            Informe("No se han obtenido los datos de los relojes checadores")
        End If
    End Sub

    Protected Friend Sub ConexionYCargaDeRelojes()
        If MyContadorDeInicio.Enabled Then
            MyContadorDeInicio.Enabled = False
            Informe("Se ha detenido el contador de inicio que inicia las cargas en las bases de datos")
        End If
        Informe("Se ha iniciado el trabajador de segundo plano")
        MyCargaOnBackGround.RunWorkerAsync()
    End Sub
    Protected Sub IniciaElTimerDeCarga()

    End Sub
    Protected Sub LecturaDeConfiguraciones()

    End Sub
    Protected Friend Sub Informe(ByVal Txt2Inform As String)
        Dim HoraAhora As DateTime = Now
        Dim RutaDeArchivo As String = MyUbicationOfLogs + "\Registro" + Date2stringArchivo(HoraAhora) + ".txt"
        If Directory.Exists(MyUbicationOfLogs) Then
            If Not File.Exists(RutaDeArchivo) Then
                Try
                    Dim MyWriter As StreamWriter = New StreamWriter(RutaDeArchivo)
                    MyWriter.WriteLine("---Este es un informe del servicio de segundo plando de la carga de los relojes---")
                    MyWriter.WriteLine("->Evento sucedido a las:" + vbTab + Date2stringSQL(Now))
                    MyWriter.WriteLine("Reporte:")
                    MyWriter.WriteLine(Txt2Inform)
                    MyWriter.WriteLine("Servicio creado por Alan Fernando Santacruz Rodríguez")
                    MyWriter.Close()
                Catch MyErrorEnEscritura As System.Exception
                    InformeDeErrorAlEscribirElLog("No se ha podido realizar la escritura del Log a las " + Date2stringSQL(Now))
                End Try
            Else
                Directory.CreateDirectory("C:\LogsDeCargaDeLosRelojes")
            End If
        End If
    End Sub
    Protected Sub InformeDeErrorAlEscribirElLog(ByVal sEvent As String)
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

    Private Sub MyCargaOnBackGround_DoWork(sender As Object, e As DoWorkEventArgs) Handles MyCargaOnBackGround.DoWork

    End Sub
End Class
