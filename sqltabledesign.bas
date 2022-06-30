Attribute VB_Name = "Module1"
Option Explicit

Sub generate_sql()
    Dim started As Boolean  ' �e�e�[�u���ɓ����Ă��邩�ǂ���
    started = False
    Dim startrow As Integer ' �e�e�[�u���̃X�^�[�g�s
    Dim endrow As Integer   ' �e�e�[�u���̏I���s
    
    Dim endsheet As Long    ' �V�[�g�̏I���s
    endsheet = Cells(Rows.Count, 2).End(xlUp).Row + 1
    
    Dim i As Integer
    For i = 1 To endsheet
        
        ' (���܂ތ����Z���̏ꍇ(�e�[�u�����̃Z�����m)
        If Cells(i, 2).MergeCells And InStr(Cells(i, 2).Value, "(") > 1 Then
            startrow = i
            started = True
        End If
      
        ' �e�[�u���ɓ����Ă��Ă��񖼂��󂾂����ꍇ(�e�[�u���̏I��茟�m)
        If Cells(i, 2).Value = "" And started Then
            endrow = i - 1
            ' SQL�𐶐����ĊJ�n�s�̉��̃Z���ɑ��
            Cells(startrow, 8).Value = convert_table(startrow, endrow)
            started = False
        End If
        
    Next i
End Sub

Function convert_table(ByVal startrow As Integer, ByVal endrow As Integer) As String
    Dim sql As String   ' ��������SQL
    Dim table_name_string As String ' �e�[�u�������������Z���̕�����
    table_name_string = Cells(startrow, 2).Value
    
    ' SQL���͂��� (�Ƃ��������̍��̕����܂ł��e�[�u�����Ƃ��Ď��
    sql = "CREATE TABLE " & Left(table_name_string, InStr(table_name_string, "(") - 1) & "(" & vbLf
    
    Dim i As Integer
    ' �e�[�u�����s��2��(���o���s���΂�����)
    For i = startrow + 2 To endrow
        ' �񖼂���łȂ����
        If Cells(i, 2).Value <> "" Then
            
            ' �ŏ��̃J�����ȊO�̓J���}��t���ċ�؂�
            If i <> startrow + 2 Then
                sql = sql & "," & vbLf
            End If
            
            sql = sql & Cells(i, 2).Value & " " ' �J������
            sql = sql & Cells(i, 3).Value   ' �f�[�^�^
            
            ' ����͂���ꍇ�̂݌�������
            If Cells(i, 4).Value <> "" Then
                sql = sql & " " & Cells(i, 4).Value
            End If
            
        End If
    Next i
    
    ' SQL���I���
    sql = sql & vbLf & ");"
    ' return
    convert_table = sql
End Function
