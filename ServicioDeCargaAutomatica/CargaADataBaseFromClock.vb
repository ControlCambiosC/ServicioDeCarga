Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports ZKSoftwareAPI
Imports System.Net
Public Class CargaADataBaseFromClock
    'Trabajador del proceso
    Private WithEvents ProcesoCompleto As BackgroundWorker = New BackgroundWorker()
    'Cambiadores
    Dim MyClvDes As MyTriosDeConsulta = New MyTriosDeConsulta("Relojes", "Descripcion", "Clave_Reloj")
    Dim MyIP_Des As MyTriosDeConsulta = New MyTriosDeConsulta("Relojes", "Descripcion", "IP_asignada")
    Dim MyClv As String = ""
    Dim MyIpC As String = ""
    Dim MyDes As String = ""
    'Conexion a la base de datos
    Const MyTable As String = "Relojes"
    Const MyTReg As String = "RegistroDeCargas"
    Protected Friend MySqlCon As AdmSQL = New AdmSQL("192.168.0.100", "RelojChecador", "usuarios", "1234")
    'Reloj Checador
    Protected Friend MyClock As ZKSoftware = New ZKSoftware(Modelo.X628C)
    'Error de nivel jerarquico
    Dim ErrorDeNlvJ As Integer = 0
    Dim MyErroresOnTxt As List(Of String) = New List(Of String) From {
        "Se ha realizado la carga de la informacion sin ningún problema",
        "Ha habído algún problema al registrar la información en la base de datos",
        "No se han podido obtener los registros de asistencia del reloj checador",
        "No se podido establecer la comunicación con el reloj checador",
        "No es válida la clave o la IP que esta registrada en el sistema"
        }
    'Fechas de inicio y finalizacion
    Dim FechaI As DateTime = New DateTime
    Dim FechaT As DateTime = New DateTime()
    'Listas de columnas
    Dim MyColumns As List(Of String) = New List(Of String)
    Dim MyColRegs As List(Of String) = New List(Of String)
    Dim Respuesta As Integer
    Sub New(ByVal MyClockDes As String, ByRef SeAlmacenaLaRespuesta As Integer)
        ProcesoCompleto.WorkerSupportsCancellation = True
        ProcesoCompleto.WorkerReportsProgress = True
        MyDes = MyClockDes
        FechaI = Now
        MyColumns = MySqlCon.SqlReaderDown2List("Select Column_name from INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" + MyTable + "'")
        MyColRegs = MySqlCon.SqlReaderDown2List("Select Column_name from INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" + MyTReg + "'")
        Respuesta = SeAlmacenaLaRespuesta
    End Sub
    Function IsValid()
        MyClv = MyClvDes.Des2Clv(MyDes, MySqlCon)
        MyIpC = MyIP_Des.Des2Clv(MyDes, MySqlCon)
        Dim IpVacia() As Byte = {0, 0, 0, 0}
        Dim IPchange As IPAddress = New IPAddress(IpVacia)
        If MyClv.Length = 8 And MyIpC.Length > 0 Then
            If IPAddress.TryParse(MyIpC, IPchange) Then
                MyIpC = IPchange.ToString()
                Return True
            End If
        End If
        Return False
    End Function

    Private Sub ProcesoCompleto_DoWork(sender As Object, e As DoWorkEventArgs) Handles ProcesoCompleto.DoWork
        Dim Respuesta As Integer = 0
        Dim ProgresoT As Integer = 0
        Dim ContadorA As Integer = 0
        Dim ContadorO As Integer = 0
        Dim ContadorE As Integer = 0
        Dim Maximo As Integer = 0
        ProcesoCompleto.ReportProgress(1)
        If IsValid() Then
            ProcesoCompleto.ReportProgress(10)
            If MyClock.DispositivoConectar(MyIpC, 5, False) Then
                ProcesoCompleto.ReportProgress(25)
                If MyClock.DispositivoObtenerRegistrosAsistencias() Then
                    ProcesoCompleto.ReportProgress(30)
                    Maximo = MyClock.ListaMarcajes.Count
                    Try
                        For Each Registro As UsuarioMarcaje In MyClock.ListaMarcajes
                            Respuesta = Marcaje2DataBase(Registro, MyColumns, MySqlCon, MyClv, MyIpC)
                            If Respuesta = 0 Then
                                ContadorA = ContadorA + 1
                            ElseIf Respuesta = 4 Or Respuesta = 3 Then
                                ContadorO = ContadorO + 1
                            Else
                                ContadorE = ContadorE + 1
                            End If
                            ProgresoT = ProgresoT + 1
                            'MyProgress.SetProgress(ProgresoTotal)
                            ProcesoCompleto.ReportProgress(Convert.ToInt16((100 * ProgresoT / Maximo) / 2 + 30))
                        Next
                        FechaT = Now
                        ProcesoCompleto.ReportProgress(81)
                        'We need to load this information to the database so ...
                        ErrorDeNlvJ = 0
                    Catch MyE As System.Exception
                        ErrorDeNlvJ = 1
                    End Try
                Else
                    ErrorDeNlvJ = 2
                End If
            Else
                ErrorDeNlvJ = 3
            End If
        Else
            ErrorDeNlvJ = 4
        End If
        '-----Registro de las acciones realizadas
        If Not CargaLog(FechaI, FechaT, Maximo, ContadorA, ContadorO, ContadorE, MyClv, MyIpC, MyErroresOnTxt(ErrorDeNlvJ)) = 1 Then
            ErrorDeNlvJ = 5
        End If
        MyClock.DispositivoDesconectar()
        ProcesoCompleto.ReportProgress(100)
    End Sub

    Function CargaLog(FechaI, FechaT, Maximo, ContadorA, ContadorO, ContadorE, MyClv, MyIpC, MyError)
        Dim DataToRegister As List(Of String) = New List(Of String)
        Dim NuevasC As List(Of String) = Cut2List(MyColRegs, 1, MyColRegs.Count - 1)
        Respuesta = ErrorDeNlvJ
        DataToRegister.Add(Date2stringSQL(FechaI))
        DataToRegister.Add(Date2stringSQL(FechaT))
        DataToRegister.Add(Maximo)
        DataToRegister.Add(ContadorA.ToString)
        DataToRegister.Add(ContadorO.ToString)
        DataToRegister.Add(ContadorE.ToString)
        DataToRegister.Add(MyClv)
        DataToRegister.Add(MyIpC)
        DataToRegister.Add(MyError)
        ProcesoCompleto.ReportProgress(85)
        Dim ListaAInsertar As List(Of String) = MySqlCon.InsertComillas(DataToRegister)
        If ListaAInsertar.Count > 0 Then
            If MySqlCon.InsertaEnSql(MyTReg, ListaAInsertar) Then
                Return 1
            Else
                Return 0
            End If
        Else
            Return -1
        End If
    End Function

End Class
