Attribute VB_Name = "Module1"
'======================================================================================
' �� �@�\      : Excel�V�[�g���֑̋�������u�����A������31�����ȓ��ɒ�������
' �� ����      : sheetName (String) - ���̃V�[�g��
' �� �߂�l    : String - ���`��̃V�[�g��
'======================================================================================
Private Function SanitizeSheetName(ByVal sheetName As String) As String
    Dim invalidChars As String
    Dim i As Long
    invalidChars = "[]*/\?:" ' Excel�̃V�[�g���Ŏg�p���֎~����Ă��镶��
    
    ' �֑�������"_"�Ɉꊇ�u��
    For i = 1 To Len(invalidChars)
        sheetName = Replace(sheetName, Mid(invalidChars, i, 1), "_")
    Next i
    
    ' �V�[�g���̒�����31�����ȓ��ɐ؂�l�߂�
    If Len(sheetName) > 31 Then
        SanitizeSheetName = Left(sheetName, 31)
    Else
        SanitizeSheetName = sheetName
    End If
End Function


'======================================================================================
' �� ���C������: �O���u�b�N�̃f�[�^���ЂȌ`�ɓ]�L���A�ʂ̃t�@�C���Ƃ��ĕۑ�����
'======================================================================================
Sub �O���u�b�N����t�@�C���]�L�����s_�ЂȌ`���p��()

    '//--------------------------------------------------------------------------------
    '// �ϐ��錾
    '//--------------------------------------------------------------------------------
    ' --- �I�u�W�F�N�g�ϐ� ---
    Dim wbA As Workbook         ' �]�L���u�b�N (�u�b�NA)
    Dim wbB As Workbook         ' �]�L��u�b�N (�ЂȌ`�u�b�N)
    Dim wsA As Worksheet        ' �]�L���V�[�g (���[�v�p)
    Dim wsB As Worksheet        ' �]�L��V�[�g
    Dim templateSheet As Worksheet ' �ЂȌ`�ƂȂ�V�[�g
    Dim transferRange As Range  ' �]�L����f�[�^�͈�
    
    ' --- ������E���l�ϐ� ---
    Dim folderPath As String    ' �]�L���t�H���_�̃p�X
    Dim templatePath As String  ' �ЂȌ`�u�b�N�̃p�X
    Dim destFolderPath As String ' �ۑ���t�H���_�̃p�X
    Dim fileName As String      ' �������̓]�L���t�@�C����
    Dim lastRowA As Long        ' �]�L���V�[�g�̍ŏI�s
    Dim startCol As String      ' �]�L�J�n��
    Dim endCol As String        ' �]�L�I����
    Dim newFileName As String   ' �ۑ��p�̐V�����t�@�C����
    Dim dotPos As Long          ' �g���q�̈ʒu
    Dim baseName As String      ' �g���q���������t�@�C����
    Dim extension As String     ' �g���q
    Dim direction As String     ' ���[�U�[���I��������� (u/d)
    
    ' �ЂȌ`�V�[�g����萔�Ƃ��Ē�`
    Const TEMPLATE_SHEET_NAME As String = "JZXXXXXX�@�d����`��"

    '//--------------------------------------------------------------------------------
    '// STEP 1: ���O�ݒ�ƃ��[�U�[�ɂ��p�X�E�I�v�V�����̑I��
    '//--------------------------------------------------------------------------------
    Application.ScreenUpdating = False

    ' 1-1. �]�L���t�H���_�̑I��
    MsgBox "�͂��߂ɁA�]�L���̃f�[�^���������t�H���_��I�����Ă��������B", vbInformation
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�]�L����Excel�t�@�C���������Ă���t�H���_��I�����Ă�������"
        If .Show = True Then folderPath = .SelectedItems(1) & Application.PathSeparator Else Exit Sub
    End With

    ' 1-2. �ЂȌ`�u�b�N�̑I��
    MsgBox "���ɁA���C�A�E�g���ݒ肳�ꂽ�u�ЂȌ`�u�b�N�v��I�����Ă��������B", vbInformation
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "�ЂȌ`�ƂȂ�Excel�u�b�N��I�����Ă�������"
        .Filters.Clear
        .Filters.Add "Excel �t�@�C��", "*.xlsx; *.xlsm; *.xls"
        If .Show = True Then templatePath = .SelectedItems(1) Else Exit Sub
    End With

    ' 1-3. �ۑ���t�H���_�̑I��
    MsgBox "�Ō�ɁA�쐬�����t�@�C���́u�ۑ���t�H���_�v��I�����Ă��������B", vbInformation
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�t�@�C���̕ۑ���t�H���_��I�����Ă�������"
        If .Show = True Then destFolderPath = .SelectedItems(1) & Application.PathSeparator Else Exit Sub
    End With
    
    ' 1-4. �]�L�����̑I��
    direction = InputBox("Up�ł����HDown�ł����H �uu�v�܂��́ud�v�œ��͂��Ă�������", "�����̑I��")
    Select Case UCase(Trim(direction))
        Case "U": startCol = "A": endCol = "T"
        Case "D": startCol = "A": endCol = "M"
        Case Else
            MsgBox "���͂��uu�v�܂��́ud�v�ł͂���܂���B" & vbCrLf & "�����𒆒f���܂��B"
            Exit Sub
    End Select

    '//--------------------------------------------------------------------------------
    '// STEP 2: �t�H���_���̃t�@�C�������񏈗� (���C�����[�v)
    '//--------------------------------------------------------------------------------
    fileName = Dir(folderPath & "*.xls*")

    Do While fileName <> ""
        ' �G���[���������Ă������𒆒f�����A���̃t�@�C���֐i��
        On Error Resume Next
        Set wbA = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
        Set wbB = Workbooks.Open(templatePath)
        On Error GoTo 0

        If Not wbA Is Nothing And Not wbB Is Nothing Then
            ' 2-1. �ЂȌ`�V�[�g�̑��݂��m�F
            On Error Resume Next
            Set templateSheet = wbB.Worksheets(TEMPLATE_SHEET_NAME)
            On Error GoTo 0
            
            If templateSheet Is Nothing Then
                MsgBox "�ЂȌ`�u�b�N�Ɂu" & TEMPLATE_SHEET_NAME & "�v�V�[�g��������܂���B" & vbCrLf & "�t�@�C���u" & fileName & "�v�̏������X�L�b�v���܂��B", vbExclamation
            Else
                ' 2-2. �]�L���u�b�N�̑S�V�[�g�����[�v����
                For Each wsA In wbA.Worksheets
                    ' �ЂȌ`�V�[�g���R�s�[���ĐV�����V�[�g���쐬
                    templateSheet.Copy After:=wbB.Worksheets(wbB.Worksheets.Count)
                    Set wsB = wbB.Worksheets(wbB.Worksheets.Count)
                    
                    lastRowA = wsA.Cells(wsA.Rows.Count, startCol).End(xlUp).Row

                    ' 2�s�ڈȍ~�Ƀf�[�^�����݂���ꍇ�̂ݓ]�L���������s
                    If lastRowA >= 2 Then
                        Set transferRange = wsA.Range(startCol & "2:" & endCol & lastRowA)
                        wsB.Range(startCol & "2").Resize(transferRange.Rows.Count, transferRange.Columns.Count).Value = transferRange.Value
                    End If
                    
                    ' 2-3. ����̃Z���ɌŒ蕶����𑾎��œ���
                    Select Case UCase(Trim(direction))
                        Case "U"
                            With wsB.Range("T8")
                                .Value = "im�I�u�W�F�N�g��"
                                .Font.Bold = True
                            End With
                        Case "D"
                            With wsB.Range("M8")
                                .Value = "im�I�u�W�F�N�g��"
                                .Font.Bold = True
                            End With
                    End Select
                    
                    ' 2-4. Z1�Z�����ꎞ���p���ăV�[�g����ύX��A�N���A
                    wsB.Range("Z1").Value = wsA.Name
                    On Error Resume Next
                    wsB.Name = SanitizeSheetName(wsB.Range("Z1").Value)
                    On Error GoTo 0
                    wsB.Range("Z1").ClearContents
                Next wsA
                
                ' 2-5. ���̂ЂȌ`�V�[�g���폜
                Application.DisplayAlerts = False
                templateSheet.Delete
                Application.DisplayAlerts = True
                
                ' 2-6. �ۑ��t�@�C�������쐬 ("_�]�L�ς�"��t�^)
                dotPos = InStrRev(fileName, ".")
                If dotPos > 0 Then
                    baseName = Left(fileName, dotPos - 1)
                    extension = Mid(fileName, dotPos)
                    newFileName = baseName & "_�]�L�ς�" & extension
                Else
                    newFileName = fileName & "_�]�L�ς�"
                End If
                
                ' 2-7. �V�����u�b�N�Ƃ��ĕۑ�
                wbB.SaveAs fileName:=destFolderPath & newFileName, FileFormat:=wbA.FileFormat
            End If
            
            ' �J�����u�b�N�����
            wbB.Close SaveChanges:=False
            wbA.Close SaveChanges:=False
        Else
            ' �t�@�C�����J���Ȃ������ꍇ�A�f�o�b�O�p�Ƀ��O���o��
            If wbA Is Nothing Then Debug.Print "�t�@�C�����J���܂���ł���: " & folderPath & fileName
            If wbB Is Nothing Then Debug.Print "�ЂȌ`�u�b�N���J���܂���ł���: " & templatePath
        End If
        
        ' �I�u�W�F�N�g�ϐ���������A���̃��[�v�ɔ�����
        Set wbA = Nothing: Set wsA = Nothing: Set wbB = Nothing: Set wsB = Nothing: Set templateSheet = Nothing
        
        ' ���̃t�@�C�����擾
        fileName = Dir
    Loop

    '//--------------------------------------------------------------------------------
    '// STEP 3: ��������
    '//--------------------------------------------------------------------------------
    Application.ScreenUpdating = True
    MsgBox "�������������܂����B" & vbCrLf & "�t�@�C���� """ & destFolderPath & """ �ɕۑ�����Ă��܂��B"

End Sub

