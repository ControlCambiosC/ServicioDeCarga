Imports ZKSoftwareAPI
Imports System.Data.SqlClient
Module FuncioneDeBusquedaEnMarcajes
    Class DatosDeMarcajeClase
        Property NumeroCredencial As Integer = 0
        Property Year As Integer = 1
        Property Mes As Integer = 2
        Property Dia As Integer = 3
        Property Hora As Integer = 4
        Property Minuto As Integer = 5
        Property Segundo As Integer = 6
        Property MarcajeO As Integer = 7
    End Class
    Public DatosDeMarcaje As DatosDeMarcajeClase = New DatosDeMarcajeClase()

    ''' <summary>
    ''' Permite consultar un párametro de la clase  UsuarioMarcaje
    ''' </summary>
    ''' <returns></returns>
    Function ConsultaUnParametroDeLM(ByRef LMarcaje As UsuarioMarcaje, ByVal Parametro As Integer)
        Select Case Parametro
            Case 0
                Return LMarcaje.NumeroCredencial
            Case 1
                Return LMarcaje.Anio
            Case 2
                Return LMarcaje.Mes
            Case 3
                Return LMarcaje.Dia
            Case 4
                Return LMarcaje.Hora
            Case 5
                Return LMarcaje.Minuto
            Case 6
                Return LMarcaje.Segundo
            Case 7
                Return LMarcaje.MarcajeOperativo
            Case Else
                Return ""
        End Select
    End Function
    Function Marcaje2str(ByRef MarcajeOp As MarcajeOperativo)
        Dim MyStr As String = ""
        MyStr = MarcajeOp.NumeroDispositivo + " " + MarcajeOp.Parametro1 + " " + MarcajeOp.Parametro2 + " " + MarcajeOp.Parametro3 + " " + MarcajeOp.Parametro4 + " " + MarcajeOp.Operacion
        Return MyStr
    End Function
    Function ComparaLM(ByRef LMarcaje As UsuarioMarcaje, ByVal Parametro As Integer, ByVal Comparativo As Integer)
        If ConsultaUnParametroDeLM(LMarcaje, Parametro) = Comparativo Then
            Return True
        Else
            Return False
        End If
    End Function
    Function ComparaLM(ByRef LMarcaje As UsuarioMarcaje, ByVal Comparativo As MarcajeOperativo)
        Try
            If Marcaje2str(ConsultaUnParametroDeLM(LMarcaje, DatosDeMarcaje.MarcajeO)) = Marcaje2str(Comparativo) Then
                Return True
                Exit Function
            Else
                Return False
                Exit Function
            End If
        Catch e As System.Exception
            MsgBox("No se ha podido convertir la clase MarcajeOperativo a Str",, "ModuloFuncionesDeBusquedaEnMarcajes->ComparaLM")
        End Try
        Return False
    End Function
    ''' <summary>
    ''' The name it's only to refer to find the labels for every thing
    ''' </summary>
    ''' <returns></returns>
    Function GetTheKClusters()
        Return 0
    End Function
    ''' <summary>
    ''' Convierte un marcaje en una lista de strings con su respectivo
    ''' En caso de error regresa una lista con un -1
    ''' </summary>
    ''' <param name="LMarcaje"></param>
    ''' <returns></returns>
    Function Marcaje2ListOnlyDateTime(ByRef LMarcaje As UsuarioMarcaje)
        Dim MyData As List(Of Integer) = New List(Of Integer)
        Try
            For Index = 1 To 7
                MyData.Add(Convert.ToInt16(ConsultaUnParametroDeLM(LMarcaje, Index)))
            Next
            Return MyData
        Catch
            MyData.Clear()
            MyData.Add(-1)
            Return MyData
        End Try
    End Function
    Function Marcaje2FormatSqlInList(ByRef LMarcaje As UsuarioMarcaje)
        Dim Persona As String = Convert.ToString(LMarcaje.Anio)
        Dim FechaOnList As List(Of Integer) = Marcaje2ListOnlyDateTime(LMarcaje)
        Dim MyError As String = ""
        Dim MyData As List(Of String) = New List(Of String)
        If FechaOnList.Count = 6 Then
            Dim MyDate As String = Date2stringSQL(List2date(FechaOnList))
            MyData.Add(MyDate + Persona)
            MyData.Add(Persona)
            MyData.Add(MyDate)
            Return MyData
        Else
            MyData.Clear()
            Return MyData
        End If
    End Function
    ''' <summary>
    ''' Descripcion de los errores 
    ''' -> (-4), no se afecto la base de datos
    ''' -> (-3), ya existía ese registro
    ''' -> (-2), las listas eran de tamaños diferentes, muy raro
    ''' -> (-1), no se ha podido obtener esa información
    ''' </summary>
    ''' <param name="LMarcaje"></param>
    ''' <param name="Columnas"></param>
    ''' <param name="MySqlCon"></param>
    ''' <returns></returns>
    Function Marcaje2DataBase(ByRef LMarcaje As UsuarioMarcaje, ByVal Columnas As List(Of String), ByRef MySqlCon As AdmSQL)
        Dim ListOfData As List(Of String) = Marcaje2FormatSqlInList(LMarcaje)
        Const MyTable As String = "RelojChecador"
        If ListOfData.Count > 0 Then
            Dim Igualdades As List(Of String) = MySqlCon.RetornaIgualdades(Columnas, ListOfData)
            If Igualdades.Count > 0 Then
                Dim Consulta As String = MySqlCon.ArmaConSql(MyTable, Columnas, Igualdades)
                Dim MyType As Integer
                If Not MySqlCon.ExisteLaconsulta(Consulta, "ClvUsuario", MyType) Then
                    If MySqlCon.InsertaEnSql(MyTable, ListOfData) Then
                        Return 1
                    Else
                        Return -4

                    End If
                Else
                    Return -3
                End If
            Else
                Return -2
            End If
        Else
            Return -1
        End If
    End Function

End Module
