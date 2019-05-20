Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports ZKSoftwareAPI
Imports System.Net
Public Class CargaADataBaseFromClock
    'Trabajador del proceso
    'Private WithEvents ProcesoCompleto As BackgroundWorker = New BackgroundWorker()
    'Cambiadores
    Dim MyClvDes As MyTriosDeConsulta = New MyTriosDeConsulta("Relojes", "Descripcion", "Clave_Reloj")
    Dim MyClv_Ip As MyTriosDeConsulta = New MyTriosDeConsulta("Relojes", "IP_asignada", "Clave_Reloj")
    Dim MyClv As String = ""
    Dim MyIpC As String = ""
    Dim MyDes As String = ""
    'Conexion a la base de datos
    Const MyTable As String = "Relojes"
    Const MyTReg As String = "RegistroDeCargas"
    Const RegAsi As String = "RegReloj"

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
    Dim MyColAsis As List(Of String) = New List(Of String)
    Dim Respuesta As Integer = 0
    Sub New(ByVal MyClockClv As String) ' ByRef SeAlmacenaLaRespuesta As Integer)
        'ProcesoCompleto.WorkerSupportsCancellation = True
        'ProcesoCompleto.WorkerReportsProgress = True
        MyClv = MyClockClv
        FechaI = Now
        MyColumns = MySqlCon.SqlReaderDown2List("Select Column_name from INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" + MyTable + "'")
        MyColRegs = MySqlCon.SqlReaderDown2List("Select Column_name from INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" + MyTReg + "'")
        MyColAsis = MySqlCon.SqlReaderDown2List("Select Column_name from INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" + RegAsi + "'")
        'Respuesta = SeAlmacenaLaRespuesta
    End Sub
    Function IsValid()
        MyDes = MyClvDes.Clv2Desc(MyClv, MySqlCon)
        MyIpC = MyClv_Ip.Clv2Desc(MyClv, MySqlCon)
        Dim IpVacia() As Byte = {0, 0, 0, 0}
        Dim IPchange As IPAddress = New IPAddress(IpVacia)
        If MyClv.Length = 8 And MyIpC.Length > 0 And MyDes.Length > 0 Then
            If IPAddress.TryParse(MyIpC, IPchange) Then
                MyIpC = IPchange.ToString()
                Return True
                Exit Function
            End If
        End If
        Informe("Los datos del reloj son invalidos: " + vbNewLine +
                "Clave        : " + MyClv + vbNewLine +
                "Descripcion  : " + MyDes + vbNewLine +
                "Ip registrada: " + MyIpC
                )
        Return False
    End Function

    Public Sub ProcesoCompleto()
        Dim Respuesta As Integer = 0
        Dim ProgresoT As Integer = 0
        Dim ContadorA As Integer = 0
        Dim ContadorO As Integer = 0
        Dim ContadorPorOmisionCom As Integer = 0
        Dim ContadorE As Integer = 0
        Dim ContadorDeNoInsercion = 0
        Dim ContadorDeErrorOnSqlS = 0
        Dim ContadorDeListasIncom = 0
        Dim ContadorDeMissinCompa = 0
        Dim Maximo As Integer = 0
        'ProcesoCompleto.ReportProgress(1)
        If IsValid() Then
            'ProcesoCompleto.ReportProgress(10)
            If MyClock.DispositivoConectar(MyIpC, 5, False) Then
                ' ProcesoCompleto.ReportProgress(25)
                If MyClock.DispositivoObtenerRegistrosAsistencias() Then
                    'ProcesoCompleto.ReportProgress(30)
                    Maximo = MyClock.ListaMarcajes.Count
                    Try
                        For Each Registro As UsuarioMarcaje In MyClock.ListaMarcajes
                            Respuesta = Marcaje2DataBase(Registro, MyColAsis, MySqlCon, MyClv, MyIpC)
                            If Respuesta = 0 Then
                                ContadorA = ContadorA + 1
                            ElseIf Respuesta = 1 Then
                                ContadorDeNoInsercion = ContadorDeNoInsercion + 1
                            ElseIf Respuesta = -1 Then
                                ContadorDeErrorOnSqlS = ContadorDeErrorOnSqlS + 1
                            ElseIf Respuesta = 2 Then
                                ContadorO = ContadorO + 1
                            ElseIf Respuesta = 3 Then
                                ContadorPorOmisionCom = ContadorPorOmisionCom + 1
                            ElseIf Respuesta = 4 Then
                                ContadorDeMissinCompa = ContadorDeMissinCompa + 1
                            ElseIf Respuesta = 5 Then
                                ContadorDeListasIncom = ContadorDeListasIncom + 1
                            Else
                                ContadorE = ContadorE + 1
                            End If
                            ProgresoT = ProgresoT + 1
                            'MyProgress.SetProgress(ProgresoTotal)
                            'ProcesoCompleto.ReportProgress(Convert.ToInt16((100 * ProgresoT / Maximo) / 2 + 30))
                        Next
                        FechaT = Now
                        ContadorE = ContadorE + ContadorDeErrorOnSqlS + ContadorDeNoInsercion + ContadorDeListasIncom + ContadorDeMissinCompa
                        Informe("Resumen de operaciones del reloj:" + vbTab + MyClv + vbNewLine +
                                "Se han añadido: " + vbTab + ContadorA.ToString + vbNewLine +
                                "Se han omitido: " + vbTab + ContadorO.ToString + vbNewLine +
                                "Por omisión Co: " + vbTab + ContadorPorOmisionCom.ToString + vbNewLine +
                                "Con errores   : " + vbTab + ContadorE.ToString + vbNewLine +
                                "ErrorSqlScript: " + vbTab + ContadorDeErrorOnSqlS.ToString + vbNewLine +
                                "ErrorNoInserci:" + vbTab + ContadorDeNoInsercion.ToString + vbNewLine +
                                "ErrorEnListaCo:" + vbTab + ContadorDeMissinCompa.ToString + vbNewLine +
                                "ErrorEnObtenLi:" + vbTab + ContadorDeListasIncom.ToString)
                        'ProcesoCompleto.ReportProgress(81)
                        'We need to load this information to the database so ...
                        ErrorDeNlvJ = 0
                    Catch MyE As System.Exception
                        Informe("Error al subir a la base de datos del reloj: " + vbTab + MyClv + vbTab + MyDes + vbNewLine +
                                "Error: " + MyE.Message.ToString)
                        ErrorDeNlvJ = 1
                    End Try
                Else
                    Informe("No se ha podido obtener los registros de asistencia del reloj: " + vbTab + MyClv + vbTab + MyDes + vbNewLine +
                            "Error: " + MyClock.ERROR.ToString)
                    ErrorDeNlvJ = 2
                End If
            Else
                Informe("No se ha podido conectar al reloj: " + vbTab + MyClv + vbTab + MyDes + vbNewLine +
                            "Error: " + MyClock.ERROR.ToString)
                ErrorDeNlvJ = 3
            End If
        Else
            ErrorDeNlvJ = 4
        End If
        '-----Registro de las acciones realizadas
        If Not CargaLog(FechaI, FechaT, Maximo, ContadorA, ContadorO, ContadorE, MyClv, MyIpC, MyErroresOnTxt(ErrorDeNlvJ)) = 1 Then
            Dim DataToRegister As List(Of String) = New List(Of String)
            DataToRegister.Add(Date2stringSQL(FechaI).ToString)
            DataToRegister.Add(Date2stringSQL(FechaT).ToString)
            DataToRegister.Add(Maximo.ToString)
            DataToRegister.Add(ContadorA.ToString)
            DataToRegister.Add(ContadorO.ToString)
            DataToRegister.Add(ContadorE.ToString)
            DataToRegister.Add(MyClv.ToString)
            DataToRegister.Add(MyIpC.ToString)
            DataToRegister.Add(ErrorDeNlvJ.ToString + vbTab + " - " + MyErroresOnTxt(ErrorDeNlvJ))
            Informe("No se ha podido cargar la informacion al log de cargas " + List2Secciones(DataToRegister))
        End If
        MyClock.DispositivoDesconectar()
        'ProcesoCompleto.ReportProgress(100)
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
        'ProcesoCompleto.ReportProgress(85)
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
