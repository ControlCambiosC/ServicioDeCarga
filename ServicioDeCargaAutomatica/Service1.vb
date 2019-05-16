Imports ZKSoftwareAPI
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO.TextReader
Imports System.IO.TextWriter

Public Class CargaDiaria
    Protected Friend MyContadorDeCarga As New Timers.Timer
    Protected Friend MyClocks As List(Of ZKSoftware) = New List(Of ZKSoftware)
    Protected Friend MyTimeToCheck As Integer = 10000 'It's on seconds
    Protected Friend MySqlCon As AdmSQL = New AdmSQL("192.168.0.100", "usuarios", "1234")


    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.
        MyContadorDeCarga = New Timers.Timer
        AddHandler MyContadorDeCarga.Elapsed, AddressOf ProcesoGeneral
    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
    End Sub
    Protected Friend Sub LecturaDeDatos()

    End Sub
    Protected Sub ProcesoGeneral()

    End Sub
    Protected Sub LecturaDeConfiguraciones()

    End Sub

    Sub Main()

    End Sub
End Class
