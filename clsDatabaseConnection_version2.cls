'**********************************************************************************
' This is a class module I designed for ease of using/writing queries on a vba
'application.
' -Emmanuel Valencia
'**********************************************************************************

Option Base 0
Private queryState As Integer
Private dbConnection As Integer

Private Function ConnectToDatabaseServer(dbQuerySetting As String) As Boolean
    Dim dbCon As ADODB.Connection
    Dim dbRecset As ADODB.Recordset
    Dim localConnectionSetting As String
    Dim strPath As String
    Dim strConnection As String
    Dim counter As Integer
    Dim index As Integer
    Dim arrayCombinedValues As Variant
    
    On Error GoTo Errorhandler:
    
    'LIST YOUR SOURCE PATHS HERE - COMMENT OUT THE CURRENT DATABASE BEING USED
    '=======================================================================================================
    If IsNull(dbConnection) Then
        'dev default repository
        strPath = "" 'location of your default dev backend
    
    Else
    
        Select Case dbConnection
        Case 0
            strPath = "" 'location of 1st repository
        Case 1
            strPath = "" 'location of 2nd repository
        Case 2
            strPath = "" 'location of 3rd repository
        Case Else
            strPath = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name  'specify -1, to connect to the local workbook
        End Select
    
    End If

    '=======================================================================================================
    
    If dbConnection = -1 Then
        strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                        "Data Source=" & strPath & ";" & _
                        "Extended Properties=" & """" & "Excel 12.0 Macro;HDR=YES;" & """"
    
    Else
        strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                        "Data Source=" & strPath & ";" & _
                        "Jet OLEDB:Engine Type=5;" & _
                        "Persist Security Info=TRUE;"
    End If

    Set dbCon = New ADODB.Connection
    dbCon.Open strConnection
    Set dbRecset = New ADODB.Recordset
    
    If dbCon.State = 1 Then
        
        ConnectToDatabaseServer = True
        'LOG_EVENT "Connection Successful to Parallon shared."
        Select Case queryState
        Case 1
            Set dbRecset = dbCon.Execute(dbQuerySetting)
            counter = 0
            index = 0
            
            If Not dbRecset.EOF Then
                While Not dbRecset.EOF
                    ReDim Preserve HeaderValuesArray(dbRecset.Fields.Count - 1) As Variant
                    ReDim Preserve OuterValuesArray(index) As Variant
                    
                    For counter = 0 To dbRecset.Fields.Count - 1
                        HeaderValuesArray(counter) = dbRecset.Fields(counter).Name
                        ReDim Preserve InnerValuesArray(counter) As Variant
                        InnerValuesArray(counter) = dbRecset.Fields(counter).Value
                    Next counter
                    OuterValuesArray(index) = InnerValuesArray
                    index = index + 1
                    dbRecset.MoveNext
                Wend
            Else
                    ReDim Preserve InnerValuesArray(0) As Variant
                    ReDim Preserve OuterValuesArray(0) As Variant
                    InnerValuesArray(0) = Null
                    OuterValuesArray(0) = InnerValuesArray
            End If
        Case Else
            dbCon.Execute dbQuerySetting
        End Select
        
    End If
    
    If dbRecset.State = 1 Then
        dbRecset.Close
    End If
    Set dbRecset = Nothing
    dbCon.Close
    Set dbCon = Nothing

    Exit Function

Errorhandler:
    'LOG_EVENT "Connection Error"
    ConnectToDatabaseServer = False
    MsgBox Err.Description & vbCrLf & vbCrLf & "Please contact MIS regarding this issue.", vbCritical + vbOKOnly, "Upload"
    ConnectToDatabaseServer = False
    
End Function

Public Function QueryDatabase(sqlQuery As String, Optional ConnectToDb As Integer) As Boolean
    dbConnection = ConnectToDb
    If ParseSQLQuery(sqlQuery) = True Then
        If ConnectToDatabaseServer(sqlQuery) = True Then
            QueryDatabase = True
        Else
            QueryDatabase = False
        End If
    End If
End Function

Private Function ParseSQLQuery(sqlQuery As String) As Boolean
    ParseSQLQuery = False
    sqlQuery = UCase(sqlQuery)
    
    
    If InStr(1, sqlQuery, "INSERT") Then
        
        ParseSQLQuery = True
        queryState = 2
    
    ElseIf InStr(1, sqlQuery, "DELETE") Then
    
        ParseSQLQuery = True
        queryState = 4
        
    ElseIf (Not InStr(1, sqlQuery, "SELECT") = 0) And (Not InStr(1, sqlQuery, "FROM") = 0) Then
        
        ParseSQLQuery = True
        queryState = 1
        
    ElseIf InStr(1, sqlQuery, "UPDATE") Then
    
        ParseSQLQuery = True
        queryState = 3
        
    Else
    
        ParseSQLQuery = True
        queryState = 5
    
    End If

End Function

Public Sub DestroyArray()
    ReDim Preserve InnerValuesArray(0) As Variant
    ReDim Preserve OuterValuesArray(0) As Variant
    InnerValuesArray(0) = Null
    OuterValuesArray(0) = InnerValuesArray
End Sub
