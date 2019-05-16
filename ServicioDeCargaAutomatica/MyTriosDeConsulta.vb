Public Class MyTriosDeConsulta
    Dim MyTable As String
    Dim MyColDe As String
    Dim MyCoClv As String

    ''' <summary>
    ''' Armado de la consulta, esto es para evitar definir tantas variables
    ''' Tabla
    ''' Columna que contiene la descripcion My_CoDe
    ''' Columna que contiene la clave   My_CoClv
    ''' </summary>
    ''' <param name="My_Table"></param>
    ''' <param name="My_CoDe"></param>
    ''' <param name="My_CoClv"></param>
    Sub New(ByVal My_Table As String, ByVal My_CoDe As String, ByVal My_CoClv As String)
        MyTable = My_Table
        MyColDe = My_CoDe
        MyCoClv = My_CoClv
    End Sub

    Function ReturnList(ByRef MySqlCon As AdmSQL)
        Dim Consulta = "Select " + MyColDe + " from " + MyTable
        Return MySqlCon.SqlReaderDown2List(Consulta)
    End Function

    Function Clv2Desc(ByVal MyStr As String, ByRef MySqlCon As AdmSQL)
        Return MySqlCon.BusquedaDeCambioEnSql(MyTable, MyColDe, MyCoClv, MyStr)
    End Function

    Function Des2Clv(ByVal MyStr As String, ByRef MySqlCon As AdmSQL)
        Return MySqlCon.BusquedaDeCambioEnSql(MyTable, MyCoClv, MyColDe, MyStr)
    End Function

End Class

Public Class ChecadorDeExistencia
    Dim MyTabla As String
    Dim MyColumna As String
    Dim WithColum As String
    Dim MyMetodo As Integer

    Sub New(ByVal Table As String, ByVal Columna As String)
        MyTabla = Table
        MyColumna = Columna
        MyMetodo = 1
    End Sub

    Sub New(ByVal Table As String, ByVal Columna As String, ByVal AndColumn As String)
        MyTabla = Table
        MyColumna = Columna
        WithColum = AndColumn
        MyMetodo = 2
    End Sub

    Function RetMetodo()
        Return MyMetodo
    End Function

    Function Existe(ByVal MyString As String, ByRef MySqlCon As AdmSQL)
        If MyMetodo = 1 Then
            Dim Consulta As String = "Select " + MyColumna + " From " + MyTabla + " where " + MyColumna + " = '" + MyString + "'"
            Dim MyList As List(Of String) = MySqlCon.SqlReaderDown2List(Consulta)
            If MyList.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            MsgBox("El método a usar es incorrecto en Checador de Existencia")
            Return False
        End If
    End Function

    Function ExisteWith(ByVal MyString As String, ByVal Cond2 As String, ByRef MySqlCon As AdmSQL)
        If MyMetodo = 2 Then
            Dim Consulta As String = "Select " + MyColumna + " From " + MyTabla + " where " + MyColumna + " = '" + MyString + "' And " + WithColum + " = '" + Cond2 + "'"
            Dim MyList As List(Of String) = MySqlCon.SqlReaderDown2List(Consulta)
            If MyList.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            MsgBox("El método a usar es incorrecto en Checador de Existencia")
            Return False
        End If
    End Function

End Class