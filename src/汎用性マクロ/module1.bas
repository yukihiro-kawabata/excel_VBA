'##########################################################
'
' �f�B���N�g�����ɂ���t�@�C���S�Ă��擾����
'
' == �d�l ===
' �t�H���_Path���G�N�Z�����ɋL�ڂ���
' �L�ڂ����Z����I��������ԂŎ��s�����
' �L�ڃZ�����牺2�s�ڂɃt�@�C���ꗗ�������o�����
'
'##########################################################
Sub getFileListInDir()

    '�I���ꏊ���擾���ăf�B���N�g��Path���擾����
    dirPath = ActiveSheet.Cells(Selection.Row, Selection.Column) & "\"

    '���̍s���珑���o��
    '���ۂɂ͋󗓂��܂ނ̂ŁA���̎����珑���o��
    cnt = Selection.Row + 1

    '�t�@�C�������ׂď����o��
    filePath = Dir(dirPath & "*")
    Do While filePath <> ""
        cnt = cnt + 1
        Cells(cnt, 1) = filePath
        filePath = Dir()
    Loop
End Sub
