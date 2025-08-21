Attribute VB_Name = "Module1"
'======================================================================================
' ■ 機能      : Excelシート名の禁則文字を置換し、長さを31文字以内に調整する
' ■ 引数      : sheetName (String) - 元のシート名
' ■ 戻り値    : String - 整形後のシート名
'======================================================================================
Private Function SanitizeSheetName(ByVal sheetName As String) As String
    Dim invalidChars As String
    Dim i As Long
    invalidChars = "[]*/\?:" ' Excelのシート名で使用が禁止されている文字
    
    ' 禁則文字を"_"に一括置換
    For i = 1 To Len(invalidChars)
        sheetName = Replace(sheetName, Mid(invalidChars, i, 1), "_")
    Next i
    
    ' シート名の長さを31文字以内に切り詰める
    If Len(sheetName) > 31 Then
        SanitizeSheetName = Left(sheetName, 31)
    Else
        SanitizeSheetName = sheetName
    End If
End Function


'======================================================================================
' ■ メイン処理: 外部ブックのデータをひな形に転記し、個別のファイルとして保存する
'======================================================================================
Sub 外部ブックからファイル転記を実行_ひな形利用版()

    '//--------------------------------------------------------------------------------
    '// 変数宣言
    '//--------------------------------------------------------------------------------
    ' --- オブジェクト変数 ---
    Dim wbA As Workbook         ' 転記元ブック (ブックA)
    Dim wbB As Workbook         ' 転記先ブック (ひな形ブック)
    Dim wsA As Worksheet        ' 転記元シート (ループ用)
    Dim wsB As Worksheet        ' 転記先シート
    Dim templateSheet As Worksheet ' ひな形となるシート
    Dim transferRange As Range  ' 転記するデータ範囲
    
    ' --- 文字列・数値変数 ---
    Dim folderPath As String    ' 転記元フォルダのパス
    Dim templatePath As String  ' ひな形ブックのパス
    Dim destFolderPath As String ' 保存先フォルダのパス
    Dim fileName As String      ' 処理中の転記元ファイル名
    Dim lastRowA As Long        ' 転記元シートの最終行
    Dim startCol As String      ' 転記開始列
    Dim endCol As String        ' 転記終了列
    Dim newFileName As String   ' 保存用の新しいファイル名
    Dim dotPos As Long          ' 拡張子の位置
    Dim baseName As String      ' 拡張子を除いたファイル名
    Dim extension As String     ' 拡張子
    Dim direction As String     ' ユーザーが選択する方向 (u/d)
    
    ' ひな形シート名を定数として定義
    Const TEMPLATE_SHEET_NAME As String = "JZXXXXXX　電文定義書"

    '//--------------------------------------------------------------------------------
    '// STEP 1: 事前設定とユーザーによるパス・オプションの選択
    '//--------------------------------------------------------------------------------
    Application.ScreenUpdating = False

    ' 1-1. 転記元フォルダの選択
    MsgBox "はじめに、転記元のデータが入ったフォルダを選択してください。", vbInformation
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "転記元のExcelファイルが入っているフォルダを選択してください"
        If .Show = True Then folderPath = .SelectedItems(1) & Application.PathSeparator Else Exit Sub
    End With

    ' 1-2. ひな形ブックの選択
    MsgBox "次に、レイアウトが設定された「ひな形ブック」を選択してください。", vbInformation
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "ひな形となるExcelブックを選択してください"
        .Filters.Clear
        .Filters.Add "Excel ファイル", "*.xlsx; *.xlsm; *.xls"
        If .Show = True Then templatePath = .SelectedItems(1) Else Exit Sub
    End With

    ' 1-3. 保存先フォルダの選択
    MsgBox "最後に、作成したファイルの「保存先フォルダ」を選択してください。", vbInformation
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ファイルの保存先フォルダを選択してください"
        If .Show = True Then destFolderPath = .SelectedItems(1) & Application.PathSeparator Else Exit Sub
    End With
    
    ' 1-4. 転記方向の選択
    direction = InputBox("Upですか？Downですか？ 「u」または「d」で入力してください", "方向の選択")
    Select Case UCase(Trim(direction))
        Case "U": startCol = "A": endCol = "T"
        Case "D": startCol = "A": endCol = "M"
        Case Else
            MsgBox "入力が「u」または「d」ではありません。" & vbCrLf & "処理を中断します。"
            Exit Sub
    End Select

    '//--------------------------------------------------------------------------------
    '// STEP 2: フォルダ内のファイルを巡回処理 (メインループ)
    '//--------------------------------------------------------------------------------
    fileName = Dir(folderPath & "*.xls*")

    Do While fileName <> ""
        ' エラーが発生しても処理を中断せず、次のファイルへ進む
        On Error Resume Next
        Set wbA = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
        Set wbB = Workbooks.Open(templatePath)
        On Error GoTo 0

        If Not wbA Is Nothing And Not wbB Is Nothing Then
            ' 2-1. ひな形シートの存在を確認
            On Error Resume Next
            Set templateSheet = wbB.Worksheets(TEMPLATE_SHEET_NAME)
            On Error GoTo 0
            
            If templateSheet Is Nothing Then
                MsgBox "ひな形ブックに「" & TEMPLATE_SHEET_NAME & "」シートが見つかりません。" & vbCrLf & "ファイル「" & fileName & "」の処理をスキップします。", vbExclamation
            Else
                ' 2-2. 転記元ブックの全シートをループ処理
                For Each wsA In wbA.Worksheets
                    ' ひな形シートをコピーして新しいシートを作成
                    templateSheet.Copy After:=wbB.Worksheets(wbB.Worksheets.Count)
                    Set wsB = wbB.Worksheets(wbB.Worksheets.Count)
                    
                    lastRowA = wsA.Cells(wsA.Rows.Count, startCol).End(xlUp).Row

                    ' 2行目以降にデータが存在する場合のみ転記処理を実行
                    If lastRowA >= 2 Then
                        Set transferRange = wsA.Range(startCol & "2:" & endCol & lastRowA)
                        wsB.Range(startCol & "2").Resize(transferRange.Rows.Count, transferRange.Columns.Count).Value = transferRange.Value
                    End If
                    
                    ' 2-3. 特定のセルに固定文字列を太字で入力
                    Select Case UCase(Trim(direction))
                        Case "U"
                            With wsB.Range("T8")
                                .Value = "imオブジェクト名"
                                .Font.Bold = True
                            End With
                        Case "D"
                            With wsB.Range("M8")
                                .Value = "imオブジェクト名"
                                .Font.Bold = True
                            End With
                    End Select
                    
                    ' 2-4. Z1セルを一時利用してシート名を変更後、クリア
                    wsB.Range("Z1").Value = wsA.Name
                    On Error Resume Next
                    wsB.Name = SanitizeSheetName(wsB.Range("Z1").Value)
                    On Error GoTo 0
                    wsB.Range("Z1").ClearContents
                Next wsA
                
                ' 2-5. 元のひな形シートを削除
                Application.DisplayAlerts = False
                templateSheet.Delete
                Application.DisplayAlerts = True
                
                ' 2-6. 保存ファイル名を作成 ("_転記済み"を付与)
                dotPos = InStrRev(fileName, ".")
                If dotPos > 0 Then
                    baseName = Left(fileName, dotPos - 1)
                    extension = Mid(fileName, dotPos)
                    newFileName = baseName & "_転記済み" & extension
                Else
                    newFileName = fileName & "_転記済み"
                End If
                
                ' 2-7. 新しいブックとして保存
                wbB.SaveAs fileName:=destFolderPath & newFileName, FileFormat:=wbA.FileFormat
            End If
            
            ' 開いたブックを閉じる
            wbB.Close SaveChanges:=False
            wbA.Close SaveChanges:=False
        Else
            ' ファイルが開けなかった場合、デバッグ用にログを出力
            If wbA Is Nothing Then Debug.Print "ファイルが開けませんでした: " & folderPath & fileName
            If wbB Is Nothing Then Debug.Print "ひな形ブックが開けませんでした: " & templatePath
        End If
        
        ' オブジェクト変数を解放し、次のループに備える
        Set wbA = Nothing: Set wsA = Nothing: Set wbB = Nothing: Set wsB = Nothing: Set templateSheet = Nothing
        
        ' 次のファイルを取得
        fileName = Dir
    Loop

    '//--------------------------------------------------------------------------------
    '// STEP 3: 完了処理
    '//--------------------------------------------------------------------------------
    Application.ScreenUpdating = True
    MsgBox "処理が完了しました。" & vbCrLf & "ファイルは """ & destFolderPath & """ に保存されています。"

End Sub

