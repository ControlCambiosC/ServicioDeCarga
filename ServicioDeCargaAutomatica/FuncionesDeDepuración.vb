Module FuncionesDeDepuración

    ''' <summary>
    ''' Concatena una lista simplemente de manera que:
    ''' Elemento1 Elemento2 ... ElementoN
    ''' </summary>
    ''' <param name="MyList"></param>
    ''' <returns></returns>
    Function List2String(ByVal MyList As List(Of String))
        Dim MyStr As String = ""
        For Each str As String In MyList
            MyStr = MyStr + str + " "
        Next
        Return MyStr
    End Function

    ''' <summary>
    ''' Permite pasar de una lista a un string de la siguiente forma:
    ''' - Elemento 1
    ''' - Elemento 2
    '''     .
    '''     .
    '''     .
    ''' - Elemento n
    ''' </summary>
    ''' <param name="MyList"></param>
    ''' <returns></returns>
    Function List2Secciones(ByVal MyList As List(Of String))
        Dim MyStr As String = ""
        For Each str As String In MyList
            MyStr = MyStr + vbCrLf + "- " + str
        Next
        Return MyStr
    End Function

End Module