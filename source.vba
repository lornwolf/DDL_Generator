Option Explicit

' Folder selection using built-in FileDialog (64-bit compatible)
Function BrowseForFolder(Optional title As String = "フォルダを選択") As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .title = title
        .AllowMultiSelect = False
        .InitialFileName = ""
        
        If .Show = -1 Then
            BrowseForFolder = .SelectedItems(1)
        Else
            BrowseForFolder = ""
        End If
    End With
End Function

Sub SelectInputFolder()
    Dim folderPath As String
    folderPath = BrowseForFolder("入力用フォルダを選択")
    If folderPath <> "" Then
        Me.OLEObjects("TextBox_Input").Object.Text = folderPath
    End If
End Sub

Sub SelectOutputFolder()
    Dim folderPath As String
    folderPath = BrowseForFolder("出力用フォルダを選択")
    If folderPath <> "" Then
        Me.OLEObjects("TextBox_Output").Object.Text = folderPath
    End If
End Sub

Sub Button_Input_Click()
    Call SelectInputFolder
End Sub

Sub Button_Output_Click()
    Call SelectOutputFolder
End Sub

Sub Button_Execute_Click()
    Call ExecuteDDLGeneration
End Sub

Sub ExecuteDDLGeneration()
    Dim inputFolder As String
    Dim outputFolder As String
    Dim fso As Object
    Dim file As Object
    Dim files As Object
    Dim excelApp As Object
    Dim Workbook As Object
    Dim ws As Object
    Dim tableId As String
    Dim tableName As String
    Dim ddlContent As String
    Dim warningMsg As String
    Dim errorMsg As String
    Dim fileCount As Integer
    Dim i As Long
    Dim fieldName As String
    Dim fieldId As String
    Dim fieldType As String
    Dim fieldLen As String
    Dim fieldDec As String
    Dim isPK As Boolean
    Dim isNullable As Boolean
    Dim nullStr As String
    Dim primaryKeys As String
    Dim sqlType As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    warningMsg = ""
    errorMsg = ""
    ddlContent = ""
    fileCount = 0
    
    On Error Resume Next
    inputFolder = Me.OLEObjects("TextBox_Input").Object.Text
    If Err.Number <> 0 Then
        MsgBox "TextBox_Input コントロールを追加してください！", vbCritical, "エラー"
        Exit Sub
    End If
    On Error GoTo 0
    
    On Error Resume Next
    outputFolder = Me.OLEObjects("TextBox_Output").Object.Text
    If Err.Number <> 0 Then
        MsgBox "TextBox_Output コントロールを追加してください！", vbCritical, "エラー"
        Exit Sub
    End If
    On Error GoTo 0
    
    If inputFolder = "" Or Not fso.FolderExists(inputFolder) Then
        MsgBox "エラー：入力用フォルダのパスが無効または存在しません！", vbCritical, "エラー"
        Exit Sub
    End If
    
    If outputFolder = "" Or Not fso.FolderExists(outputFolder) Then
        MsgBox "エラー：出力用フォルダのパスが無効または存在しません！", vbCritical, "エラー"
        Exit Sub
    End If
    
    Set files = fso.GetFolder(inputFolder).files
    Dim hasExcel As Boolean
    hasExcel = False
    For Each file In files
        If LCase(fso.GetExtensionName(file.Name)) = "xls" Or LCase(fso.GetExtensionName(file.Name)) = "xlsx" Then
            hasExcel = True
            Exit For
        End If
    Next
    
    If Not hasExcel Then
        MsgBox "エラー：入力用フォルダにExcelファイル(.xlsまたは.xlsx)がありません！", vbCritical, "エラー"
        Exit Sub
    End If
    
    If MsgBox("DDL生成を開始しますか？", vbYesNo + vbQuestion, "確認") = vbNo Then
        Exit Sub
    End If
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    
    On Error Resume Next
    
    For Each file In files
        If LCase(fso.GetExtensionName(file.Name)) = "xls" Or LCase(fso.GetExtensionName(file.Name)) = "xlsx" Then
            fileCount = fileCount + 1
            
            Set Workbook = excelApp.Workbooks.Open(file.path, False, True)
            
            If Workbook.Sheets.Count >= 2 Then
                Set ws = Workbook.Sheets(2)
                
                tableId = Trim(CStr(ws.Range("D2").Value))
                tableName = Trim(CStr(ws.Range("E2").Value))
                
                If tableId = "" Or tableName = "" Then
                    errorMsg = errorMsg & "ファイル " & file.Name & " のテーブルIDまたはテーブル名が空です！" & vbCrLf
                    Workbook.Close False
                    GoTo NextFile
                End If
                
                ddlContent = ddlContent & "-- Table: " & tableId & " (" & tableName & ")" & vbCrLf
                ddlContent = ddlContent & "DROP TABLE IF EXISTS " & tableId & ";" & vbCrLf
                ddlContent = ddlContent & "CREATE TABLE " & tableId & " (" & vbCrLf
                
                primaryKeys = ""
                i = 6
                
                Do While True
                    fieldName = Trim(CStr(ws.Cells(i, 4).Value))
                    fieldId = Trim(CStr(ws.Cells(i, 5).Value))
                    
                    If fieldName = "" And fieldId = "" Then
                        Exit Do
                    End If
                    
                    Dim pkMark As String
                    pkMark = Trim(CStr(ws.Cells(i, 2).Value))
                    If pkMark <> "" And pkMark <> "PK" And Left(pkMark, 1) <> "P" Then
                        warningMsg = warningMsg & "ファイル " & file.Name & " の" & i & "行目のB列に無効な値 '" & pkMark & "' が含まれています。無視しました！" & vbCrLf
                        pkMark = ""
                    End If
                    isPK = (pkMark = "PK" Or Left(pkMark, 1) = "P")
                    
                    Dim nullMark As String
                    nullMark = Trim(CStr(ws.Cells(i, 3).Value))
                    If nullMark <> "" And nullMark <> "Y" Then
                        warningMsg = warningMsg & "ファイル " & file.Name & " の" & i & "行目のC列に無効な値 '" & nullMark & "' が含まれています。無視しました！" & vbCrLf
                        nullMark = ""
                    End If
                    isNullable = (nullMark = "Y")
                    
                    If fieldName = "" Then
                        errorMsg = errorMsg & "ファイル " & file.Name & " の" & i & "行目のフィールド名が空です。処理中止！" & vbCrLf
                        Workbook.Close False
                        GoTo NextFile
                    End If
                    
                    fieldId = Trim(CStr(ws.Cells(i, 5).Value))
                    If fieldId = "" Then
                        errorMsg = errorMsg & "ファイル " & file.Name & " の" & i & "行目のフィールドIDが空です。処理中止！" & vbCrLf
                        Workbook.Close False
                        GoTo NextFile
                    End If
                    If InStr(fieldId, " ") > 0 Then
                        errorMsg = errorMsg & "ファイル " & file.Name & " の" & i & "行目のフィールドIDにスペースが含まれています。処理中止！" & vbCrLf
                        Workbook.Close False
                        GoTo NextFile
                    End If
                    
                    fieldType = Trim(CStr(ws.Cells(i, 6).Value))
                    If fieldType = "" Then
                        errorMsg = errorMsg & "ファイル " & file.Name & " の" & i & "行目のフィールドタイプが空です。処理中止！" & vbCrLf
                        Workbook.Close False
                        GoTo NextFile
                    End If
                    
                    sqlType = ConvertToMySQLType(fieldType)
                    If sqlType = "" Then
                        errorMsg = errorMsg & "ファイル " & file.Name & " の" & i & "行目のフィールドタイプ '" & fieldType & "' はMySQLでサポートされていません。処理中止！" & vbCrLf
                        Workbook.Close False
                        GoTo NextFile
                    End If
                    
                    fieldLen = Trim(CStr(ws.Cells(i, 7).Value))
                    If fieldLen <> "" Then
                        If Not IsNumeric(fieldLen) Then
                            errorMsg = errorMsg & "ファイル " & file.Name & " の" & i & "行目のフィールド長が有効数字ではありません。処理中止！" & vbCrLf
                            Workbook.Close False
                            GoTo NextFile
                        End If
                    End If
                    
                    fieldDec = Trim(CStr(ws.Cells(i, 8).Value))
                    If fieldDec <> "" Then
                        If Not IsNumeric(fieldDec) Then
                            errorMsg = errorMsg & "ファイル " & file.Name & " の" & i & "行目の小数点以下桁数が有効数字ではありません。処理中止！" & vbCrLf
                            Workbook.Close False
                            GoTo NextFile
                        End If
                    End If
                    
                    Dim mysqlType As String
                    mysqlType = BuildMySQLType(sqlType, fieldLen, fieldDec)
                    
                    If isNullable Then
                        nullStr = "NULL"
                    Else
                        nullStr = "NOT NULL"
                    End If
                    
                    ddlContent = ddlContent & "    " & fieldId & " " & mysqlType & " " & nullStr
                    
                    If isPK Then
                        If primaryKeys <> "" Then
                            primaryKeys = primaryKeys & ", "
                        End If
                        primaryKeys = primaryKeys & fieldId
                    End If
                    
                    ddlContent = ddlContent & "," & vbCrLf
                    
                    i = i + 1
                Loop
                
                If primaryKeys <> "" Then
                    ddlContent = ddlContent & "    PRIMARY KEY (" & primaryKeys & ")" & vbCrLf
                End If
                
                ddlContent = ddlContent & ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;" & vbCrLf & vbCrLf
            Else
                warningMsg = warningMsg & "ファイル " & file.Name & " はSheetが1つしかないため、スキップしました！" & vbCrLf
            End If
            
            Workbook.Close False
        End If
        
NextFile:
    Next
    
    On Error GoTo 0
    
    excelApp.Quit
    Set excelApp = Nothing
    
    If ddlContent <> "" Then
        Dim ddlFile As Object
        Set ddlFile = fso.CreateTextFile(outputFolder & "\DDL.sql", True, True)
        ddlFile.Write ddlContent
        ddlFile.Close
    End If
    
    If warningMsg <> "" Then
        Dim warnFile As Object
        Set warnFile = fso.CreateTextFile(outputFolder & "\WARN.log", True, True)
        warnFile.Write warningMsg
        warnFile.Close
    End If
    
    If errorMsg <> "" Then
        Dim errFile As Object
        Set errFile = fso.CreateTextFile(outputFolder & "\ERROR.log", True, True)
        errFile.Write errorMsg
        errFile.Close
    End If
    
    Dim summary As String
    summary = "処理完了！" & vbCrLf & vbCrLf
    summary = summary & "合計 " & fileCount & " 件のExcelファイルを処理しました。" & vbCrLf
    
    If ddlContent <> "" Then
        summary = summary & "DDL出力先: DDL.sql" & vbCrLf
    End If
    
    If warningMsg <> "" Then
        summary = summary & "警告内容: WARN.log" & vbCrLf
    End If
    
    If errorMsg <> "" Then
        summary = summary & "エラー内容: ERROR.log" & vbCrLf
    End If
    
    MsgBox summary, vbInformation, "完了"
    
    Set fso = Nothing
End Sub

Function ConvertToMySQLType(fieldType As String) As String
    Dim t As String
    t = LCase(Trim(fieldType))
    
    Select Case t
        Case "varchar", "nvarchar", "char", "nchar"
            ConvertToMySQLType = "VARCHAR"
        Case "text", "ntext"
            ConvertToMySQLType = "TEXT"
        Case "int", "integer"
            ConvertToMySQLType = "INT"
        Case "bigint"
            ConvertToMySQLType = "BIGINT"
        Case "smallint"
            ConvertToMySQLType = "SMALLINT"
        Case "tinyint"
            ConvertToMySQLType = "TINYINT"
        Case "decimal", "numeric"
            ConvertToMySQLType = "DECIMAL"
        Case "float", "real"
            ConvertToMySQLType = "FLOAT"
        Case "double"
            ConvertToMySQLType = "DOUBLE"
        Case "date"
            ConvertToMySQLType = "DATE"
        Case "datetime", "timestamp"
            ConvertToMySQLType = "DATETIME"
        Case "time"
            ConvertToMySQLType = "TIME"
        Case "blob", "binary", "varbinary"
            ConvertToMySQLType = "BLOB"
        Case "bit"
            ConvertToMySQLType = "BIT"
        Case "boolean", "bool"
            ConvertToMySQLType = "TINYINT"
        Case Else
            ConvertToMySQLType = ""
    End Select
End Function

Function BuildMySQLType(sqlType As String, fieldLen As String, fieldDec As String) As String
    BuildMySQLType = sqlType
    
    If sqlType = "VARCHAR" Or sqlType = "CHAR" Then
        If fieldLen <> "" Then
            BuildMySQLType = sqlType & "(" & fieldLen & ")"
        Else
            BuildMySQLType = sqlType & "(255)"
        End If
    ElseIf sqlType = "DECIMAL" Then
        If fieldLen <> "" And fieldDec <> "" Then
            BuildMySQLType = sqlType & "(" & fieldLen & "," & fieldDec & ")"
        ElseIf fieldLen <> "" Then
            BuildMySQLType = sqlType & "(" & fieldLen & ",0)"
        Else
            BuildMySQLType = sqlType & "(10,0)"
        End If
    ElseIf sqlType = "BIT" Then
        If fieldLen <> "" Then
            BuildMySQLType = sqlType & "(" & fieldLen & ")"
        Else
            BuildMySQLType = sqlType & "(1)"
        End If
    End If
End Function

Private Sub Label1_Click()

End Sub



Private Sub Input_Label_Click()

End Sub

Private Sub Output_Label_Click()

End Sub
