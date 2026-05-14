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
    Dim singleDdl As String
    Dim singleDdlFolder As String
    
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
    
    singleDdlFolder = outputFolder & "\CREATE文（テーブル単位）"
    If Not fso.FolderExists(singleDdlFolder) Then
        fso.CreateFolder singleDdlFolder
    End If
    
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

                'Excel SHA-256 ハッシュを計算（差分判定用ヘッダー）
                Dim sourceHash As String
                Dim generatedAt As String
                sourceHash = GetFileSHA256(file.path)
                generatedAt = Format(Now, "yyyy-mm-dd")

                ddlContent = ddlContent & "-- Table: " & tableId & " (" & tableName & ")" & vbCrLf
                ddlContent = ddlContent & "DROP TABLE IF EXISTS " & tableId & ";" & vbCrLf
                ddlContent = ddlContent & "CREATE TABLE " & tableId & " (" & vbCrLf
                
                '単体DDLファイルには差分判定用ヘッダーを追加
                singleDdl = "-- @source-hash: " & sourceHash & vbCrLf
                singleDdl = singleDdl & "-- @generated-at: " & generatedAt & vbCrLf
                singleDdl = singleDdl & "-- Table: " & tableId & " (" & tableName & ")" & vbCrLf
                singleDdl = singleDdl & "DROP TABLE IF EXISTS " & tableId & ";" & vbCrLf
                singleDdl = singleDdl & "CREATE TABLE " & tableId & " (" & vbCrLf
                
                ' K5に「AUTO_INCREMENT」ヘッダーが設定されているか確認
                Dim hasAutoIncrementCol As Boolean
                hasAutoIncrementCol = (UCase(Trim(CStr(ws.Cells(5, 11).Value))) = "AUTO_INCREMENT")

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

                    ' K列のAUTO_INCREMENT判定（K5に「AUTO_INCREMENT」がある場合のみ）
                    Dim isAutoIncrement As Boolean
                    isAutoIncrement = False
                    If hasAutoIncrementCol Then
                        Dim aiMark As String
                        aiMark = Trim(CStr(ws.Cells(i, 11).Value))
                        If aiMark <> "" And aiMark <> "〇" Then
                            warningMsg = warningMsg & "ファイル " & file.Name & " の" & i & "行目のK列に無効な値 '" & aiMark & "' が含まれています。無視しました！" & vbCrLf
                            aiMark = ""
                        End If
                        isAutoIncrement = (aiMark = "〇")
                    End If

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

                    ' DATETIME(n) パターン検出：F列に括弧付きで指定された場合、nを抽出
                    Dim datetimePrecisionFromType As String
                    datetimePrecisionFromType = ""
                    Dim fieldTypeLower As String
                    fieldTypeLower = LCase(fieldType)
                    If Len(fieldTypeLower) >= 10 And Left(fieldTypeLower, 9) = "datetime(" And Right(fieldTypeLower, 1) = ")" Then
                        datetimePrecisionFromType = Trim(Mid(fieldType, 10, Len(fieldType) - 10))
                        If datetimePrecisionFromType = "" Then datetimePrecisionFromType = "0"
                        fieldType = "DATETIME"
                    End If

                    sqlType = ConvertToMySQLType(fieldType)
                    If sqlType = "" Then
                        errorMsg = errorMsg & "ファイル " & file.Name & " の" & i & "行目のフィールドタイプ '" & fieldType & "' はMySQLでサポートされていません。処理中止！" & vbCrLf
                        Workbook.Close False
                        GoTo NextFile
                    End If
                    
                    fieldLen = Trim(CStr(ws.Cells(i, 7).Value))
                    If sqlType = "DATETIME" Then
                        ' DATETIME(n) が指定された場合、F列のnを使用しG列は無視する
                        If datetimePrecisionFromType <> "" Then
                            fieldLen = datetimePrecisionFromType
                        End If
                        ' 0～6以外（非数値含む）は0として処理
                        Dim dtPrecision As Integer
                        If Not IsNumeric(fieldLen) Then
                            fieldLen = "0"
                        Else
                            dtPrecision = CInt(fieldLen)
                            If dtPrecision < 0 Or dtPrecision > 6 Then
                                fieldLen = "0"
                            Else
                                fieldLen = CStr(dtPrecision)
                            End If
                        End If
                    Else
                        If fieldLen <> "" Then
                            If Not IsNumeric(fieldLen) Then
                                errorMsg = errorMsg & "ファイル " & file.Name & " の" & i & "行目のフィールド長が有効数字ではありません。処理中止！" & vbCrLf
                                Workbook.Close False
                                GoTo NextFile
                            End If
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

                    Dim columnTypeStr As String
                    Dim columnConstraints As String
                    If isAutoIncrement Then
                        ' AUTO_INCREMENTはUNSIGNED NOT NULLを強制
                        columnTypeStr = mysqlType & " UNSIGNED"
                        columnConstraints = "NOT NULL AUTO_INCREMENT"
                    Else
                        columnTypeStr = mysqlType
                        If isNullable Then
                            nullStr = "NULL"
                        Else
                            nullStr = "NOT NULL"
                        End If
                        columnConstraints = nullStr
                    End If

                    ddlContent = ddlContent & "    " & fieldId & " " & columnTypeStr & " " & columnConstraints

                    If fieldName <> "" Then
                        ddlContent = ddlContent & " COMMENT '" & Replace(fieldName, "'", "''") & "'"
                    End If

                    singleDdl = singleDdl & "    " & fieldId & " " & columnTypeStr & " " & columnConstraints

                    If fieldName <> "" Then
                        singleDdl = singleDdl & " COMMENT '" & Replace(fieldName, "'", "''") & "'"
                    End If
                    
                    If isPK Then
                        If primaryKeys <> "" Then
                            primaryKeys = primaryKeys & ", "
                        End If
                        primaryKeys = primaryKeys & fieldId
                    End If
                    
                    ddlContent = ddlContent & "," & vbCrLf
                    singleDdl = singleDdl & "," & vbCrLf
                    
                    i = i + 1
                Loop
                
                If primaryKeys <> "" Then
                    ddlContent = ddlContent & "    PRIMARY KEY (" & primaryKeys & ")" & vbCrLf
                    singleDdl = singleDdl & "    PRIMARY KEY (" & primaryKeys & ")" & vbCrLf
                End If
                
                ddlContent = ddlContent & ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='" & Replace(tableName, "'", "''") & "';" & vbCrLf & vbCrLf
                singleDdl = singleDdl & ") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COMMENT='" & Replace(tableName, "'", "''") & "';" & vbCrLf
                
                Call WriteUtf8NoBom(singleDdlFolder & "\" & tableId & "_" & tableName & ".sql", singleDdl)
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
        Call WriteUtf8NoBom(outputFolder & "\CREATE文.sql", ddlContent)
    End If

    If warningMsg <> "" Then
        Call WriteUtf8NoBom(outputFolder & "\WARN.log", warningMsg)
    End If

    If errorMsg <> "" Then
        Call WriteUtf8NoBom(outputFolder & "\ERROR.log", errorMsg)
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
        Case "longblob"
            ConvertToMySQLType = "LONGBLOB"
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
    ElseIf sqlType = "DATETIME" Then
        If fieldLen <> "" And fieldLen <> "0" Then
            BuildMySQLType = sqlType & "(" & fieldLen & ")"
        Else
            BuildMySQLType = sqlType
        End If
    End If
End Function

' 文字列をUTF-8（BOMなし）でファイルに書き込む
' ADODB.Stream を利用し、UTF-8 BOM（EF BB BF）の3バイトをスキップして保存する
Sub WriteUtf8NoBom(filePath As String, content As String)
    Dim utfStream As Object
    Dim binStream As Object

    Set utfStream = CreateObject("ADODB.Stream")
    utfStream.Type = 2 ' adTypeText
    utfStream.Mode = 3 ' adModeReadWrite
    utfStream.Charset = "utf-8"
    utfStream.Open
    utfStream.WriteText content

    ' バイナリモードに切り替えてBOMをスキップ
    utfStream.Position = 0
    utfStream.Type = 1 ' adTypeBinary
    utfStream.Position = 3 ' UTF-8 BOM (EF BB BF) をスキップ

    Set binStream = CreateObject("ADODB.Stream")
    binStream.Type = 1 ' adTypeBinary
    binStream.Mode = 3
    binStream.Open

    utfStream.CopyTo binStream
    binStream.SaveToFile filePath, 2 ' adSaveCreateOverWrite

    binStream.Close
    utfStream.Close
    Set binStream = Nothing
    Set utfStream = Nothing
End Sub

' ファイルのSHA-256ハッシュを計算する（Windows標準のcertutilを利用）
' 戻り値: 64文字の小文字16進文字列（失敗時は空文字列）
' Node.js の crypto.createHash('sha256') と同じ結果を返すため、生成ツール間で互換性あり
Function GetFileSHA256(filePath As String) As String
    Dim shell As Object
    Dim exec As Object
    Dim output As String
    Dim lines() As String
    Dim hashLine As String
    Dim i As Integer

    GetFileSHA256 = ""

    On Error GoTo ErrHandler

    Set shell = CreateObject("WScript.Shell")
    Set exec = shell.exec("cmd /c certutil -hashfile """ & filePath & """ SHA256")

    ' プロセス完了を待つ
    Do While exec.Status = 0
        DoEvents
    Loop

    output = exec.StdOut.ReadAll()

    ' certutil の出力例:
    '   SHA256 ハッシュ (ファイル <path>):
    '   8c4d89ceae943a7f1c48337640e918c982bc29643e3266bd02499cc7e182ac0c
    '   CertUtil: -hashfile コマンドは正常に完了しました。
    ' 2行目（インデックス1）がハッシュ
    lines = Split(output, vbCrLf)

    For i = 0 To UBound(lines)
        hashLine = Trim(lines(i))
        ' 16進文字のみで構成された64文字の行を探す（環境差吸収）
        If Len(Replace(hashLine, " ", "")) = 64 Then
            If IsHexString(Replace(hashLine, " ", "")) Then
                GetFileSHA256 = LCase(Replace(hashLine, " ", ""))
                Exit Function
            End If
        End If
    Next i

    Exit Function

ErrHandler:
    GetFileSHA256 = ""
End Function

' 文字列が16進数のみで構成されているか判定する
Function IsHexString(s As String) As Boolean
    Dim i As Integer
    Dim c As String
    IsHexString = False
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = LCase(Mid(s, i, 1))
        If Not ((c >= "0" And c <= "9") Or (c >= "a" And c <= "f")) Then
            Exit Function
        End If
    Next i
    IsHexString = True
End Function