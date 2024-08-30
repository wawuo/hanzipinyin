Function ConvertHanziToPinyin(hanzi As String) As String
    Dim conn As Object
    Dim rs As Object
    Dim pinyin As String
    Dim char As String
    Dim tempPinyin As String
    Dim dbPath As String
    Dim j As Integer
    Dim notFound As String
    
    ' 初始化pinyin和notFound
    pinyin = ""
    notFound = ""
    
    ' 获取当前工作簿路径
    dbPath = ThisWorkbook.Path & "\hanzi.accdb"
    
    ' 检查数据库文件是否存在
    If Dir(dbPath) = "" Then
        ConvertHanziToPinyin = "数据库文件未找到，请将数据库文件与Excel工作簿放在同一目录中。"
        Exit Function
    End If
    
    ' 创建ADO连接
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    
    ' 遍历每个字符
    For i = 1 To Len(hanzi)
        char = Mid(hanzi, i, 1)
        
        ' 查询拼音
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open "SELECT 拼音 FROM Hzpyutf8 WHERE 汉字 = '" & char & "'", conn, 1, 1
        
        If Not rs.EOF Then
            ' 处理多音字（你需要决定如何处理）
            tempPinyin = rs("拼音")
            If notFound <> "" Then
                pinyin = pinyin & " " & notFound & " "
                notFound = ""
            End If
            pinyin = pinyin & " " & tempPinyin
        Else
            ' 如果未找到，保留原字符
            notFound = notFound & char
        End If
        
        rs.Close
    Next i
    
    ' 处理最后一段未找到的字符
    If notFound <> "" Then
        pinyin = pinyin & " " & notFound & " "
    End If
    
    conn.Close
    ConvertHanziToPinyin = Trim(pinyin)
End Function

