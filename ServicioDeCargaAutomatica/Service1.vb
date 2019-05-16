Imports ZKSoftwareAPI
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO.TextReader
Imports System.IO.TextWriter

Public Class CargaDiaria
    Protected Friend MyContadorDeCarga As New Timers.Timer
    Protected Friend MyClocks As List(Of ZKSoftware) = New List(Of ZKSoftware)
    Protected Friend MyTimeToCheck As Integer = 10000 'It's on seconds

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.
        MyContadorDeCarga = New Timers.Timer

    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
    End Sub
    Protected Friend LecturaDelArchivoConfig()

    End Sub
    Protected Sub ProcesoGeneral()

    End Sub
    Protected Sub LecturaDeConfiguraciones()

    End Sub
    Protected Sub ConexionAlReloj()

    End Sub
    Protected Sub ProcesoDeAdquisicion()

    End Sub
    Protected Sub ProcesoDeDesconexion()

    End Sub
End Class
