Imports System.ComponentModel
Imports ZKSoftwareAPI

Public Class ServicioDeCargaDeRelojes
    Protected Friend MyContadorDeInicio As Timers.Timer = New Timers.Timer
    Protected Friend MyContadorDeCarga As Timers.Timer = New Timers.Timer

    Protected Const MyConsultaDeClocks As String = "Select Clave_Reloj from Relojes order by Descripcion"
    Protected CambiadorDeClvToDes As MyTriosDeConsulta = New MyTriosDeConsulta("Relojes", "Clave_Reloj", "Descripcion")
    Protected Friend MyClocksClv As List(Of String) = New List(Of String)

    Protected Friend MySqlCon As AdmSQL = New AdmSQL("192.168.0.100", "RelojChecador", "usuarios", "1234")
    Protected Const MyConsultaInicio As String = "Select Valor from ParametricosDeRelojes where Parametro like 'PuntoI%' order by Valor ASC"
    Protected Friend MyPointsInicio As List(Of Integer) = New List(Of Integer)

    Protected Const MyConsultaDeIntervalo = "Select Valor from ParametricosDeRelojes where Parametro='IntervaloEnMin'"
    Protected Const IntervaloMinimo As Integer = 10 'Minutos Es el mínimo intervalo aceptado
    Protected Const IntervaloMaximo As Integer = 2880 'Es el máximo valor del intervalo aceptado que corresponde a dos días
    Protected Friend MyIntervaloDeCarga As Integer = 90  'Este es el valor por defecto

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
        Try
            'MyContadorDeInicio = New Timers.Timer
            AddHandler MyContadorDeInicio.Elapsed, AddressOf ConfiguracionDePrimeraCarga
            'MyContadorDeCarga = New Timers.Timer
            MyContadorDeCarga.AutoReset = True
            AddHandler MyContadorDeCarga.Elapsed, AddressOf ACargarDatos
            ProcesandoElInicio()
            MyCargaOnBackGround.WorkerReportsProgress = True
            MyCargaOnBackGround.WorkerSupportsCancellation = True
        Catch MyErro As System.Exception
            Informe("Ha habído un error al iniciar el servicio: " + vbNewLine + MyErro.Message.ToString)
            Me.Finalize()
        End Try
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
            If MyPointsInicio.Count > 0 Then
                Dim NextMinute As Integer = FindTheNextNumber2Next(Now.Minute, MyPointsInicio)
                If NextMinute > 0 Then
                    Dim MyEspera As Integer = NextMinute - Now.Minute
                    MyContadorDeInicio.Interval = Minutes2Milis(MyEspera)
                    MyContadorDeInicio.Start()
                    Informe("Se ha iniciado la espera para empezar las cargas." + vbNewLine +
                            "Iniciando el servicio en " + MyEspera.ToString + " minutos, considerando el inicio al minuto " + NextMinute.ToString)
                ElseIf NextMinute = 0 Then
                    Dim MyEspera As Integer = 60 - Now.Minute
                    MyContadorDeInicio.Interval = Minutes2Milis(MyEspera)
                    MyContadorDeInicio.Start()
                    Informe("Se ha iniciado la espera para empezar las cargas." + vbNewLine +
                            "Iniciando el servicio en " + MyEspera.ToString + "  minutos, considerando el inicio al minuto " + NextMinute.ToString)
                Else
                    Informe("Se ha devuelto de la busqueda del intervalo siguiente el valor de :" + NextMinute.ToString)
                End If
            Else
                Informe("No se ha podido convertir la información de los tiempos de inicio a números: datos" + +vbNewLine + List2Secciones(ListaTemporal))
            End If
        Else
            Informe("No se han obtenido los datos de los relojes checadores")
        End If
    End Sub

    Protected Friend Sub ConfiguracionDePrimeraCarga()
        Try
            MyContadorDeInicio.Stop()
            Informe("Se ha detenido el contador de inicio que inicia las cargas en las bases de datos a las " + DateOnlyTime2stringSQL(Now) + "del " + DateOnlyDate2stringSQL(Now))
            'Preparativos del intervalo
            Dim LecturaDelIntervalo As List(Of String) = MySqlCon.SqlReaderDown2List(MyConsultaDeIntervalo)
            If LecturaDelIntervalo.Count > 0 AndAlso LecturaDelIntervalo(0).Length > 0 Then
                Dim InterT As Integer = 0
                If Integer.TryParse(LecturaDelIntervalo(0), InterT) Then
                    If InterT < IntervaloMinimo Then
                        Informe("Se ha registrado un intervalo menor al mínimo, se usará el menor registrado en la servicio" + vbNewLine +
                            "Intervalo mínimo:" + vbTab + IntervaloMinimo.ToString + vbTab + "Intervalo registrado:" + vbTab + InterT.ToString)

                        MyIntervaloDeCarga = IntervaloMinimo
                    ElseIf InterT > IntervaloMaximo Then
                        Informe("Se ha registrado un intervalo mayor al máximo, se usará el máximo registrado en la servicio" + vbNewLine +
                        "Intervalo máximo:" + vbTab + IntervaloMaximo.ToString + vbTab + "Intervalo registrado:" + vbTab + InterT.ToString)

                        MyIntervaloDeCarga = IntervaloMaximo
                    Else
                        MyIntervaloDeCarga = InterT
                        Informe("Se usara el intervalo de tiempo de " + MyIntervaloDeCarga.ToString)
                    End If
                    MyContadorDeCarga.Interval = Minutes2Milis(MyIntervaloDeCarga)
                    MyContadorDeCarga.AutoReset = True
                    MyContadorDeCarga.Start()
                    MyCargaOnBackGround.RunWorkerAsync()
                Else
                    Informe("No se ha podido convertir el valor leído de la base de datos a entero, dato leído:  " + LecturaDelIntervalo(0) + vbNewLine _
                                + "Se detendrá el servicio ")
                    OnStop()
                End If
            Else
                Informe("No se han podido obtener el intervalo de la base de datos, se detendrá el servicio.")
                OnStop()
            End If
        Catch er As System.Exception
            Informe("Se ha encontrado un error en la configuracion de primera carga, a continuacion se muestra la descripcion del error: " + vbNewLine + er.Message.ToString)
            OnStop()
        End Try

    End Sub
    Protected Friend Sub ACargarDatos()
        MyCargaOnBackGround.RunWorkerAsync()
    End Sub

    Private Sub MyCargaOnBackGround_DoWork(sender As Object, e As DoWorkEventArgs) Handles MyCargaOnBackGround.DoWork
        Informe("Se ha iniciado el proceso de carga de los relojes checadores, es un proceso que puede utilizar recursos importantes por un momento")
        MyClocksClv = MySqlCon.SqlReaderDown2List(MyConsultaDeClocks)
        If MyClocksClv.Count > 0 Then
            For Each Clv As String In MyClocksClv
                'Aquí entra el proceso de la magia eterna
                Dim MyCarga As CargaADataBaseFromClock = New CargaADataBaseFromClock(Clv)
                MyCarga.ProcesoCompleto()
                If MyCarga.ErrorDeNlvJ = 0 Then

                End If
            Next
            'Informe("Se termino la carga de: " + List2Secciones(MyClocksClv))
        Else
            Informe("No se ha podido obtener la información de los relojes checadores, se intentará en el siguiente intervalo de carga")
        End If
    End Sub

End Class