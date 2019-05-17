Imports ZKSoftwareAPI
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO.TextReader
Imports System.IO.TextWriter

Public Class CargaDiaria
    Protected Friend MyContadorDeInicio As New Timers.Timer
	Protected Friend MyContadorDeCarga As New Timers.Timer
    Protected Friend MyClocks As List(Of ZKSoftware) = New List(Of ZKSoftware)
    Protected Friend MySqlCon As AdmSQL = New AdmSQL("192.168.0.100", "RelojChecador", "usuarios", "1234")
    Protected Const MyConsultaInicio As String = "Select * from ParametricosDeRelojes where Parametro like 'PuntoI%' order by Valor ASC"
    Protected Friend MyPointsInicio As List(Of Integer) = New List(Of Integer)

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.

        MyContadorDeInicio = New Timers.Timer
        AddHandler MyContadorDeInicio.Elapsed, AddressOf ProcesoGeneral
		MyContadorDeCarga=New Timers.Timer
		AddHandler MyContadorDeCarga
    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
    End Sub
    Protected Friend Sub ProcesandoElInicio()
        Dim ListaTemporal As List(Of String) = MySqlCon.SqlReaderDown2List(MyConsultaInicio)
        MyPointsInicio = ListStr2Num(ListaTemporal)
        Dim NextMinute As Integer = FindTheNextNumber2Next(Now.Minute, MyPointsInicio)
		MyContadorDeInicio.Interval=NextMinute
		
    End Sub
    Protected Sub IniciaElTimerDeCarga()
		
    End Sub
    Protected Sub LecturaDeConfiguraciones()

    End Sub

    Sub Main()

    End Sub
End Class
