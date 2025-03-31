'creando los caracteres para la coneccion con sql server atraves de una constante
'Si en caso tu servidor no contiene contraseña
'Completar con los datos de tu servidor (Source) y base de datos(Catalog)
Const strConn = "Provider=SQLOLEDB;Data Source=COMPLETAR;Initial Catalog=COMPLETAR;Integrated Security=SSPI;"

Sub Excel_a_SQL()
    Dim Sheetdata As Worksheet
    Dim lastrow As Long, i As Long
    Dim connection As Object
    Dim command As Object
    
    ' Configuración de la hoja
    Set Sheetdata = Workbooks("NOMBRE_DE_TU_ARCHIVO.xlsm").Sheets("HOJA") ' CAMBIAR LOS DATOS
    lastrow = Sheetdata.Range("A1").End(xlDown).Row
    
    ' Crear la conexión
    Set connection = CreateObject("ADODB.Connection")
    connection.Open strConn
    
    ' Activar comando
    Set command = CreateObject("ADODB.Command")
    command.ActiveConnection = connection
    
    ' Iniciar la transacción
    connection.BeginTrans
    
    On Error GoTo ManejarError
    
    'Limpiar Tabla de CREDITOS_VIGENTES
    command.CommandText = "DELETE FROM CREDITOS_VIGENTES"
    command.Execute
    
    ' Activar comando
    Set command = CreateObject("ADODB.Command")
    command.ActiveConnection = connection
    
    ' Recorrer filas e insertar datos
    For i = 2 To lastrow
        Dim sqlInsert As String, j As Integer
        sqlInsert = "INSERT INTO CREDITOS_VIGENTES VALUES ("
        
        ' Leer celdas y armar el query
        For j = 1 To 27
            If Trim(Sheetdata.Cells(i, j).Value) = "" Then ' Para valores vacios
                sqlInsert = sqlInsert & "NULL,"
            ElseIf j = 8 Or j = 14 Or j = 27 Then ' Para Fechas
                sqlInsert = sqlInsert & "'" & Format(Sheetdata.Cells(i, j).Value, "yyyy-MM-dd") & "',"
            ElseIf j = 1 Or j = 2 Or j = 3 Or j = 4 Or j = 5 Or j = 6 Then
                sqlInsert = sqlInsert & "'" & Sheetdata.Cells(i, j).Text & "',"
            Else
                sqlInsert = sqlInsert & Sheetdata.Cells(i, j).Value & ","
            End If
        Next j
        
        ' Eliminar la última coma si existe y cerrar paréntesis
        If Trim(sqlInsert) <> "INSERT INTO CREDITOS_VIGENTES VALUES (" Then
            sqlInsert = Left(sqlInsert, Len(sqlInsert) - 1) & ")"
            command.CommandText = sqlInsert
            command.Execute
        End If

    Next i
    
    'confirmar la transacción
    connection.CommitTrans
   
    ' Mensaje de éxito
    MsgBox "Datos insertados correctamente en CREDITOS_VIGENTES", vbInformation
    
    ' Cerrar conexión y terminar procedimiento
    connection.Close
    Set command = Nothing
    Set connection = Nothing
    Exit Sub
    
'Si ocurrio un error que devuelva todo como estaba antes de iniciar la transaccion
ManejarError:
    connection.RollbackTrans
    MsgBox "Error: " & Err.Description, vbCritical
    conn.Close
    Set command = Nothing
    Set connection = Nothing
    

End Sub
