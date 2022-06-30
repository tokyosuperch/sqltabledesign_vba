Attribute VB_Name = "Module1"
Option Explicit

Sub generate_sql()
    Dim started As Boolean  ' 各テーブルに入っているかどうか
    started = False
    Dim startrow As Integer ' 各テーブルのスタート行
    Dim endrow As Integer   ' 各テーブルの終了行
    
    Dim endsheet As Long    ' シートの終了行
    endsheet = Cells(Rows.Count, 2).End(xlUp).Row + 1
    
    Dim i As Integer
    For i = 1 To endsheet
        
        ' (を含む結合セルの場合(テーブル名のセル検知)
        If Cells(i, 2).MergeCells And InStr(Cells(i, 2).Value, "(") > 1 Then
            startrow = i
            started = True
        End If
      
        ' テーブルに入っていてかつ列名が空だった場合(テーブルの終わり検知)
        If Cells(i, 2).Value = "" And started Then
            endrow = i - 1
            ' SQLを生成して開始行の横のセルに代入
            Cells(startrow, 8).Value = convert_table(startrow, endrow)
            started = False
        End If
        
    Next i
End Sub

Function convert_table(ByVal startrow As Integer, ByVal endrow As Integer) As String
    Dim sql As String   ' 生成するSQL
    Dim table_name_string As String ' テーブル名が入ったセルの文字列
    table_name_string = Cells(startrow, 2).Value
    
    ' SQL文はじめ (という文字の左の文字までをテーブル名として取る
    sql = "CREATE TABLE " & Left(table_name_string, InStr(table_name_string, "(") - 1) & "(" & vbLf
    
    Dim i As Integer
    ' テーブル名行の2つ下(見出し行を飛ばすため)
    For i = startrow + 2 To endrow
        ' 列名が空でなければ
        If Cells(i, 2).Value <> "" Then
            
            ' 最初のカラム以外はカンマを付けて区切る
            If i <> startrow + 2 Then
                sql = sql & "," & vbLf
            End If
            
            sql = sql & Cells(i, 2).Value & " " ' カラム名
            sql = sql & Cells(i, 3).Value   ' データ型
            
            ' 制約はある場合のみ結合する
            If Cells(i, 4).Value <> "" Then
                sql = sql & " " & Cells(i, 4).Value
            End If
            
        End If
    Next i
    
    ' SQL文終わり
    sql = sql & vbLf & ");"
    ' return
    convert_table = sql
End Function
