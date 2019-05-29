Imports System.Data.SqlClient

''' <summary>
''' Created by Alan , this class contains all the refernt functions or subs to do
''' it whatever your want, the mayority of all work with List(Of String)''
''' - Retorna igualdades
''' - Arma consulta
''' </summary>
Public Class AdmSQL
    Protected ConnectionString As String = ""

    '"Data Source = 192.168.0.100;Initial Catalog=ManProd;Persist Security Info=True;User ID=usuarios;Password=1234 "
    Sub New(ByVal MyIp As String, ByVal BaseD As String, ByVal Usu As String, ByVal Pwd As String)
        ConnectionString = "Data Source = " + MyIp + ";Initial Catalog=" + BaseD + ";Persist Security Info=True;User ID=" + Usu + ";Password=" + Pwd
    End Sub

    ''' <summary>
    ''' Permite retornar el string para la conexión creado en base al constructor de la clase.
    ''' </summary>
    ''' <returns></returns>
    Function RetornaElConnectionString()
        Return ConnectionString
    End Function

    ''' <summary>
    ''' Permite de una manera tener las comparaciones necesarias para relizar el cambio o actualizacion
    ''' regresa una lista de manera que quedan las columnas igualadas de esta manera Columna(n)='Lis(n)'
    ''' </summary>
    ''' <param name="Columna"></param>
    ''' <param name="Lis2"></param>
    ''' <returns></returns>
    Function RetornaIgualdades(ByVal Columna As List(Of String), ByVal Lis2 As List(Of String))
        Dim Devuelta As List(Of String) = New List(Of String)
        If Columna.Count = Lis2.Count Then
            For index As Integer = 0 To Columna.Count - 1
                Devuelta.Add(Columna(index) + " = '" + Lis2(index) + "'")
            Next
        Else
            Devuelta.Clear()
            'MsgBox("Error las listas no son del mismo tamaño")
        End If
        Return Devuelta
    End Function

    Function RetornaIgualdadesSinComillas(ByVal Columna As List(Of String), ByVal Lis2 As List(Of String))
        Dim Devuelta As List(Of String) = New List(Of String)
        If Columna.Count = Lis2.Count Then
            For index As Integer = 0 To Columna.Count - 1
                Devuelta.Add(Columna(index) + " = " + Lis2(index) + " ")
            Next
        Else
            MsgBox("Error las listas no son del mismo tamaño")
        End If
        Return Devuelta
    End Function

    'Function ConsSql2Table(ByVal Consulta As String)
    '    Dim Dt As New DataTable
    '    Dt.Clear()
    '    Try
    '        Using con As New SqlConnection(ConnectionString)
    '            con.Open()
    '            Dim da As New SqlDataAdapter(Consulta, con)
    '            con.Close()
    '            da.Fill(Dt)
    '        End Using
    '    Catch ex As System.Exception
    '        MsgBox("Error " + ex.Message)
    '    End Try
    '    Return Dt
    'End Function

    ''' <summary>
    ''' Retorna el string de consulta basado en
    ''' Select C1,C2,...,Cn From TablaSQL Where Cond1 Or/And Cond2 ...Or/And CondN
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="lColum"></param>
    ''' <param name="CondBusqueda"></param>
    ''' <param name="Condicionante"></param>
    ''' <returns></returns>
    Function ArmaConSql(ByVal TablaSQL As String, ByVal lColum As List(Of String), ByVal CondBusqueda As List(Of String), ByVal Condicionante As List(Of String))
        Dim StrConsulta = "Select "
        Dim MaxIndex = lColum.Count - 1
        Try
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    StrConsulta = StrConsulta + lColum(index)
                Else
                    StrConsulta = StrConsulta + lColum(index) + ", "
                End If
            Next
            StrConsulta = StrConsulta + " From " + TablaSQL + " Where "

            MaxIndex = CondBusqueda.Count - 1
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    StrConsulta = StrConsulta + CondBusqueda(index)
                Else
                    StrConsulta = StrConsulta + CondBusqueda(index) + Condicionante(index) + " "
                End If
            Next
        Catch er As System.Exception
            'MsgBox(er.Message,, "Módulo de armado de consulta condicionada")
            Return ""
        End Try
        Return StrConsulta.Trim
    End Function

    ''' <summary>
    ''' Permite estructurar una consulta condicionada
    ''' Select colum1,colum2,...,columnN form TablaSQL where Con1 and Cond2 ... and CondN
    ''' </summary>
    ''' <param name="lColum"></param>
    ''' <param name="TablaSQL"></param>
    ''' <param name="CondBusqueda"></param>
    ''' <returns></returns>
    Function ArmaConSql(ByVal TablaSQL As String, ByVal lColum As List(Of String), ByVal CondBusqueda As List(Of String))
        Dim StrConsulta = "Select "
        Dim MaxIndex = lColum.Count - 1
        Try
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    StrConsulta = StrConsulta + lColum(index)
                Else
                    StrConsulta = StrConsulta + lColum(index) + ", "
                End If
            Next
            StrConsulta = StrConsulta + " From " + TablaSQL + " Where "

            MaxIndex = CondBusqueda.Count - 1
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    StrConsulta = StrConsulta + CondBusqueda(index)
                Else
                    StrConsulta = StrConsulta + CondBusqueda(index) + "And "
                End If
            Next
        Catch er As System.Exception
            'MsgBox(er.Message,, "Módulo de armado de consulta condicionada")
            Return ""
        End Try
        Return StrConsulta.Trim
    End Function

    ''' <summary>
    ''' Estructura una busqueda de manera que select * from TablaSQL where Cond1 and Cond2 ... and ConN
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="CondBusqueda"></param>
    ''' <returns></returns>
    Function ArmaConSql(ByVal TablaSQL As String, ByVal CondBusqueda As List(Of String))
        Dim StrConsulta = "Select * From " + TablaSQL + " Where "
        Dim MaxIndex = CondBusqueda.Count - 1
        Try
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    StrConsulta = StrConsulta + CondBusqueda(index)
                Else
                    StrConsulta = StrConsulta + CondBusqueda(index) + " And "
                End If
            Next
        Catch er As System.Exception
            'MsgBox("Ha habido un error en la busqueda condicionada" + vbCrLf + er.Message + vbCrLf + StrConsulta,, "Error en el módulo armaSQL")
            Return ""
        End Try
        Return StrConsulta.Trim
    End Function

    ''' <summary>
    ''' Retorna Select * from TablaSQL
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <returns></returns>
    Function ArmaConSQL(ByVal TablaSQL As String)
        Return "Select * From " + TablaSQL
    End Function

    ''' <summary>
    ''' Simplificada para el uso de simplemente el String de condicion
    ''' Select * from TablaSQL where CondBusqueda
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="CondBusqueda"></param>
    ''' <returns></returns>
    Function ArmaConSql(ByVal TablaSQL As String, ByVal CondBusqueda As String)
        Dim StrConsulta = "Select * From " + TablaSQL + " Where "
        Try
            StrConsulta = StrConsulta + CondBusqueda
        Catch er As System.Exception
            'MsgBox("Ha habido un error en ArmaSql" + vbCrLf + er.Message,, "Error en el módulo armaSQL")
            Return ""
        End Try
        Return StrConsulta.Trim
    End Function

    '-------------Funciones de UPDATE de la informacion
    'UPDATE table_name
    'Set column1 = value1, column2 = value2, ...
    'WHERE condition;
    ''' <summary>
    ''' Realizara una instrucción basada en
    ''' Update TablaSQL set C1 ='CambiosSQL1', C2 ='CambiosSQL2',..., CN ='CambiosSQLN' Where Con1 And Con2 And Con3
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="CambiosSQL"></param>
    ''' <param name="Condiciones"></param>
    Function UpdateOnSQL(ByVal TablaSQL As String, ByVal CambiosSQL As List(Of String), ByVal Condiciones As List(Of String))
        Dim Comando As String = "Update " + TablaSQL + " set "
        Dim Exito As Integer = 0
        Try
            Dim MaxIndex = CambiosSQL.Count - 1
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    Comando = Comando + CambiosSQL(index)
                Else
                    Comando = Comando + CambiosSQL(index) + ", "
                End If
            Next
            Comando = Comando + " where "
            MaxIndex = Condiciones.Count - 1
            For index As Integer = 0 To MaxIndex
                If index = MaxIndex Then
                    Comando = Comando + Condiciones(index)
                Else
                    Comando = Comando + Condiciones(index) + " and "
                End If
            Next
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Comando, con)
                Dim respuesta = cmd.ExecuteNonQuery()
                If respuesta = 0 Then
                    Exito = -1
                    'MsgBox("No se ha modificado ninguna columna")
                End If
                con.Close()
            End Using
        Catch er As System.Exception
            'MsgBox("Ha habido un error en la busqueda condicionada" + vbCrLf + er.Message,, "Error en el módulo armaSQL")
            Exito = -1
        End Try
        Return Exito
    End Function

    ''' <summary>
    ''' Examina en la Tabla SQL en base a la consulta y para acelerar el proceso solo se toma una columna, si esta existe existe el resto
    ''' para ello es necesario recuperar el contenido de la columna en base a un tipo propuesto.
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="Consulta"></param>
    ''' <param name="MyColumn"></param>
    ''' <param name="MyType"></param>
    ''' <returns></returns>
    Function ExisteLaconsulta(ByVal TablaSQL As String, ByVal Consulta As String, ByVal MyColumn As String, ByVal MyType As Object)
        Dim existe As Integer = 0
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                cmd.Parameters.AddWithValue("@" + MyColumn, MyType)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    existe = 1
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            existe = -1
        End Try
        Return existe
    End Function
    Function ExisteLaconsulta(ByVal Consulta As String, ByVal MyColumn As String, ByVal MyType As Object)
        Dim existe As Integer = 0
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                cmd.Parameters.AddWithValue("@" + MyColumn, MyType)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    existe = 1
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            existe = -1
        End Try
        Return existe
    End Function
    ''' <summary>
    ''' Se regresa la informacion de la consulta en una lista con el orden previsto según el orden de la otras dos listas
    ''' </summary>
    ''' <param name="Consulta"></param>
    ''' <param name="MyColumns"></param>
    ''' <returns></returns>
    Function SqlReader2List(ByVal Consulta As String, ByVal MyColumns As List(Of String))
        Dim listaC As List(Of String) = New List(Of String)
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If (dr.Read()) Then
                    For Each str As String In MyColumns
                        If dr(str) Is DBNull.Value Then
                            listaC.Add("")
                        Else
                            'listaC.Add(dr(str).ToString)
                            listaC.Add(Convert.ToString(dr(str)))
                        End If

                    Next
                End If
                con.Close()
            End Using
        Catch ex As System.Exception
            'MsgBox("Se ha detectado un error: " + vbCrLf + ex.Message + vbCrLf + Consulta, MsgBoxStyle.Critical, "Error en el modúlo ConsultaDeLinea Consulta-Columnas")
            listaC.Clear()
        End Try
        Return listaC
    End Function

    ''' <summary>
    ''' Esta función hace uso de
    ''' Insert Inot TablaSQL Values(MyValues1,MyValues2,...MyValuesN)
    ''' </summary>
    ''' <param name="TablaSQL"></param>
    ''' <param name="MyValues"></param>
    ''' <returns></returns>
    Function InsertaEnSql(ByVal TablaSQL As String, ByVal MyValues As List(Of String))
        'Dim MyError = False
        Dim Exito As Integer = 0
        Try
            Dim Comando As String = "Insert Into " + TablaSQL + " Values ("
            Dim Limite = MyValues.Count - 1
            For Index As Integer = 0 To Limite
                If Index = Limite Then
                    Comando = Comando + MyValues(Index) + ")"
                Else
                    Comando = Comando + MyValues(Index) + ", "
                End If
            Next
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Comando, con)
                Dim Respuesta = cmd.ExecuteNonQuery()
                con.Close()
                If Respuesta > 0 Then
                    Exito = 1
                Else
                    Exito = -1
                End If
            End Using
        Catch ex As System.Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error en el modúlo InsertaEnSQL")
            'MyError = True
            Exito = -2
        End Try
        Return Exito
    End Function

    ''' <summary>
    ''' Permite pasar de
    ''' Dato-> 'Dato'
    ''' Pero con elementos de una lista
    ''' </summary>
    ''' <param name="MyList"></param>
    ''' <returns></returns>
    Function InsertComillas(ByVal MyList As List(Of String))
        Dim Retorno As List(Of String) = New List(Of String)
        For Each str As String In MyList
            Retorno.Add("'" + str + "'")
        Next
        Return Retorno
    End Function

    Function InsertComillas(ByVal MyList As List(Of String), ByVal Pini As Integer, Pfin As Integer)
        Dim Retorno As List(Of String) = New List(Of String)
        Try
            For Index As Integer = 0 To MyList.Count - 1
                If (Pini <= Index And Index <= Pfin) Then
                    Retorno.Add(" '" + MyList(Index) + "' ")
                Else
                    Retorno.Add(MyList(Index))
                End If
            Next
        Catch er As System.Exception
            'MsgBox(er.Message)
            Retorno.Clear()
        End Try
        Return Retorno
    End Function

    ''' <summary>
    ''' Permite ejecutar la consulta definida, siendo la TablaSQL unicamente como dato informativo.
    ''' </summary>
    ''' <param name="Consulta"></param>
    ''' <returns></returns>
    Function SqlReaderDown2List(ByVal Consulta As String)
        Dim listaC As List(Of String) = New List(Of String)
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Consulta, con)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While (dr.Read())
                    If Not dr(0) Is DBNull.Value Then
                        listaC.Add(dr(0))
                    End If
                    'listaC.Add(dr(0))
                End While
                con.Close()
            End Using
        Catch ex As System.Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Error en el modúlo SqlReaderDown2List")
            listaC.Clear()
        End Try
        Return listaC
    End Function

    ''' <summary>
    ''' Permite realizar la consulta de esta manera
    ''' Select MyColR from MyTable where MyColB='MyStr'
    ''' </summary>
    ''' <param name="Mytable"></param>
    ''' <param name="MyColR"></param>
    ''' <param name="MyColB"></param>
    ''' <param name="MyStr"></param>
    ''' <returns></returns>
    Function BusquedaDeCambioEnSql(ByVal Mytable As String, ByVal MyColR As String, ByVal MyColB As String, ByVal MyStr As String)
        Dim MyError As Boolean = True
        Dim Consulta As String = "Select " + MyColR + " from " + Mytable + " where " + MyColB + " ='" + MyStr + "'"
        Dim MyR As List(Of String) = SqlReaderDown2List(Consulta)
        If MyR.Count = 0 Then
            Return "" 'e de error
        Else
            Return MyR(0)
        End If
    End Function

    ''' <summary>
    ''' Aquí se cargan los registros que consideramos valiosos, recuerda encomillar aquellos datos que de verdad
    ''' lo necesiten
    ''' </summary>
    ''' <returns></returns>
    Function InsertaEnSql(ByVal MyTable As String, ByVal MyColumns As List(Of String), ByVal MyData As List(Of String))
        Dim Exito As Integer = 0
        Try
            If MyColumns.Count = MyData.Count Then
                Dim Comando As String = "Insert Into " + MyTable + "("
                Dim Index As Integer = 0
                Dim UltIn As Integer = MyColumns.Count - 1
                For Index = 0 To UltIn
                    If Index = UltIn Then
                        Comando = Comando + MyColumns(Index) + ")"
                    Else
                        Comando = Comando + MyColumns(Index) + ", "
                    End If
                Next
                Comando = Comando + " values("
                For Index = 0 To UltIn
                    If Index = UltIn Then
                        Comando = Comando + MyTable(Index) + ")"
                    Else
                        Comando = Comando + MyTable(Index) + ", "
                    End If
                Next
                Using con As New SqlConnection(ConnectionString)
                    con.Open()
                    Dim cmd As New SqlCommand(Comando, con)
                    Dim Respuesta = cmd.ExecuteNonQuery()
                    If Respuesta > 0 Then
                        Exito = 1
                    Else
                        Exito = -1
                    End If
                    con.Close()
                End Using
            Else
                'MsgBox("Las listas no tienen el mismo tamaño en InsertaEnSql",, "AdminSQL")
                Exito = -2
            End If
        Catch er As System.Exception
            'MsgBox(er.Message,, "AdminSQL")
            Exito = -3
        End Try
        Return Exito
    End Function
    ''' <summary>
    ''' Esta función intentará crear una nueva base de datos
    ''' -1 La base de datos ya existe
    ''' -0 Hubo un error
    ''' 1 Creada sin problemas
    ''' </summary>
    Function CreaBD(ByVal NombreBD As String)
        Const ComandoI As String = "Create database "
        Dim Comando As String = ComandoI + NombreBD.Trim

        Return 0
    End Function
    ''' <summary>
    ''' Esta funcion retorna un uno en caso de exito, -3 en caso de que ya exista
    ''' -2 En caso de que exista una tabla con el número de columnas incompletas
    ''' -1 En el caso de un error en SQL y no se haya creado
    ''' </summary>
    ''' <param name="NombreTb"></param>
    ''' <param name="Columnas"></param>
    ''' <returns></returns>
    Function CreaTb(ByVal NombreTb As String, ByVal Columnas As List(Of String))
        'Inicial setting of variables'
        Dim Existe = IsTableExists(NombreTb, Columnas.Count)
        If Existe = -1 Then
            Dim MyStartStr = "Create table " + NombreTb + " ("
            Dim MaxiIndex As Integer = Columnas.Count - 1
            For Index As Integer = 0 To MaxiIndex
                If Index <> MaxiIndex Then
                    MyStartStr = MyStartStr + Columnas(Index) + ", "
                Else
                    MyStartStr = MyStartStr + Columnas(Index) + ") "
                End If
            Next
            Dim Comando = MyStartStr
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Dim cmd As New SqlCommand(Comando, con)
                Dim Respuesta = cmd.ExecuteNonQuery()
                If Respuesta > 0 Then
                    Return 1
                Else
                    Return 0
                End If
                con.Close()
            End Using
            Return -1
        ElseIf Existe = 0 Then
            Return -2
        Else
            Return -3
        End If
    End Function
    ''' <summary>
    ''' Permite conocer si una tabla en Sql existe de manera contraria, regresa un -1
    ''' El caso 0 es si no contiene el número esperado de columnas la tabla
    ''' </summary>
    ''' <param name="NombreTb"></param>
    ''' <param name="NumberOfColums"></param>
    ''' <returns></returns>
    Function IsTableExists(ByVal NombreTb As String, ByVal NumberOfColums As Integer)
        Dim BusquedaDeTablaExistente = "Select Column_name From INFORMATION_SCHEMA.COLUMNS Where TABLE_NAME = '" + NombreTb + "'"
        Dim MyExistence As List(Of String) = SqlReaderDown2List(BusquedaDeTablaExistente)
        If MyExistence.Count = NumberOfColums Then
            Return 1
        ElseIf MyExistence.Count > 0 Then
            Return 0
        Else
            Return -1
        End If
    End Function
End Class