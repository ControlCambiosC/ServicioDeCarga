Imports ZKSoftwareAPI
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO.TextReader
Imports System.IO.TextWriter
Imports System
Imports System.IO
Imports System.Collections
Imports System.Diagnostics

Public Class Service1
    Protected Friend MyContadorDeInicio As New Timers.Timer
    Protected Friend MyContadorDeCarga As New Timers.Timer

    Protected Friend MyClocks As List(Of ZKSoftware) = New List(Of ZKSoftware)

    Protected Friend MySqlCon As AdmSQL = New AdmSQL("192.168.0.100", "RelojChecador", "usuarios", "1234")
    Protected Const MyConsultaInicio As String = "Select * from ParametricosDeRelojes where Parametro like 'PuntoI%' order by Valor ASC"
    Protected Friend MyPointsInicio As List(Of Integer) = New List(Of Integer)

    Dim Const MyUbicationOfLogs As String = "C:\LogsDeCargaDeLosRelojes"

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.

        MyContadorDeInicio = New Timers.Timer
        AddHandler MyContadorDeInicio.Elapsed, AddressOf ProcesandoElInicio
        MyContadorDeCarga = New Timers.Timer
        AddHandler MyContadorDeCarga.Elapsed, AddressOf ConexionYCargaDeRelojes
        ProcesandoElInicio()
    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
    End Sub

    Protected Friend Sub ProcesandoElInicio()
        Dim ListaTemporal As List(Of String) = MySqlCon.SqlReaderDown2List(MyConsultaInicio)
        If ListaTemporal.Count > 0 Then
            MyPointsInicio = ListStr2Num(ListaTemporal)
            Dim NextMinute As Integer = FindTheNextNumber2Next(Now.Minute, MyPointsInicio)
            MyContadorDeInicio.Interval = NextMinute
        End If
    End Sub

    Protected Friend Sub ConexionYCargaDeRelojes()

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

                End Try
            Else
                Directory.CreateDirectory("C:\LogsDeCargaDeLosRelojes")
            End If
        End If
    End Sub

End Class
