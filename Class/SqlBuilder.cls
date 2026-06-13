VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Lang VBA

Option Compare Database
Option Explicit


Private TopSqlString As String
Private SelectAllField As Boolean
Private IsDerivedTable As Boolean
Private DistinctSqlString As String
Private FromTableName As String
Private IntoTableName As String
Private DerivedTableName As String

Private FieldCollection As Collection
Private ValueCollection As Collection
Private JoinCollection As Collection
Private WhereCollection As Collection
Private GroupCollection As Collection
Private HavingCollection As Collection
Private OrderCollection As Collection
Private ParamDefCollection As Collection
Private ParamTypeCollection As Collection
Private ParamCollection As Collection


'判断验证
Private Function HasField() As Boolean
    If SelectAllField Then
        HasField = True
    ElseIf FieldCollection Is Nothing Then
        HasField = False
    Else
        HasField = (FieldCollection.Count > 0)
    End If
End Function
Private Function Validation() As Boolean
    If FieldCollection Is Nothing Or ValueCollection Is Nothing Then
        Validation = False
    Else
        Validation = (ValueCollection.Count > 0 And FieldCollection.Count = ValueCollection.Count)
    End If
End Function
Private Function HasInto() As Boolean
    HasInto = Not StringBase.IsWhiteSpace(IntoTableName)
End Function
Private Function HasSource() As Boolean
    HasSource = Not StringBase.IsWhiteSpace(FromTableName)
End Function
Private Function HasWhere() As Boolean
    If WhereCollection Is Nothing Then
        HasWhere = False
    Else
        HasWhere = (WhereCollection.Count > 0)
    End If
End Function

'是否使用参数
Public Function HasParam() As Boolean
    If ParamDefCollection Is Nothing Or ParamTypeCollection Is Nothing Or ParamCollection Is Nothing Then
        HasParam = False
    Else
        HasParam = (ParamDefCollection.Count > 0 And ParamDefCollection.Count = ParamTypeCollection.Count And ParamDefCollection.Count = ParamCollection.Count)
    End If
End Function
'是否作为派生表（子查询表）
Public Function HasDerived() As Boolean
    HasDerived = IsDerivedTable
End Function


'指定记录数
Public Sub Top(ByVal value As Long)
    If value > 0 Then TopSqlString = "TOP " & value & " "
End Sub
'去除重复值
Public Sub Distinct(Optional ByVal value As Boolean = True)
    If value = True Then
        DistinctSqlString = "DISTINCT "
    Else
        DistinctSqlString = ""
    End If
End Sub
'选择全部字段
Public Sub SelectAll(Optional ByVal value As Boolean = True)
    SelectAllField = value
End Sub


' 数据源
Public Sub From(ByRef FromTable As Variant, Optional ByVal Alias As String = "")
    If VBA.VarType(FromTable) = VBA.vbString And Not StringBase.IsWhiteSpace(FromTable) Then
        If Not StringBase.IsWhiteSpace(Alias) Then
            FromTableName = FromTable & " AS " & Alias
        Else
            FromTableName = FromTable
        End If
    ElseIf TypeOf FromTable Is SqlBuilder Then
        If Not StringBase.IsWhiteSpace(Alias) Then FromTable.Derived Alias
        If FromTable.HasDerived Then
            FromTableName = FromTable.ToSqlString
        Else
            Err.Raise VBA.vbObjectError + 3075, "SqlBuilder.From", "子查询必须设置别名"
        End If
    Else
        Err.Raise VBA.vbObjectError + 3075, "SqlBuilder.From", "FROM 语句错误"
    End If
End Sub

Public Sub Into(ByVal TableName As String)
    If Not StringBase.IsWhiteSpace(TableName) Then IntoTableName = TableName
End Sub

Public Function Alias() As String
    Alias = DerivedTableName
End Function

'开启子查询的派生表，设置别名
Public Sub Derived(ByVal Alias As String)
    If Not StringBase.IsWhiteSpace(Alias) Then
        IsDerivedTable = True
        DerivedTableName = Alias
    Else
        IsDerivedTable = False
        DerivedTableName = ""
    End If
End Sub


' 表连接
Public Sub Join(ByRef JoinTable As Variant, ByVal SourceTable As String, ByVal JoinField As String, Optional ByVal SourceField As String = "")
    If StringBase.IsWhiteSpace(SourceTable) Then Exit Sub
    If StringBase.IsWhiteSpace(JoinField) Then Exit Sub
    If StringBase.IsWhiteSpace(SourceField) Then SourceField = JoinField

    If JoinCollection Is Nothing Then Set JoinCollection = New Collection

    If VBA.VarType(JoinTable) = VBA.vbString And Not StringBase.IsWhiteSpace(JoinTable) Then
        JoinCollection.Add " INNER JOIN " & JoinTable & " ON " & SourceTable & "." & SourceField & "=" & JoinTable & "." & JoinField
    ElseIf TypeOf JoinTable Is SqlBuilder Then
        If JoinTable.HasDerived Then
            JoinCollection.Add " INNER JOIN " & JoinTable.ToSqlString & " ON " & SourceTable & "." & SourceField & "=" & JoinTable.Alias & "." & JoinField
        Else
            Err.Raise VBA.vbObjectError + 3075, "SqlBuilder.Join", "子查询必须设置别名"
        End If
    Else
        Err.Raise VBA.vbObjectError + 3075, "SqlBuilder.Join", "INNER JOIN 语句错误"
    End If
End Sub

' 左连接
Public Sub LeftJoin(ByRef JoinTable As Variant, ByVal SourceTable As String, ByVal JoinField As String, Optional ByVal SourceField As String = "")
    If StringBase.IsWhiteSpace(SourceTable) Then Exit Sub
    If StringBase.IsWhiteSpace(JoinField) Then Exit Sub
    If StringBase.IsWhiteSpace(SourceField) Then SourceField = JoinField

    If JoinCollection Is Nothing Then Set JoinCollection = New Collection

    If VBA.VarType(JoinTable) = VBA.vbString And Not StringBase.IsWhiteSpace(JoinTable) Then
        JoinCollection.Add " LEFT JOIN " & JoinTable & " ON " & SourceTable & "." & SourceField & "=" & JoinTable & "." & JoinField
    ElseIf TypeOf JoinTable Is SqlBuilder Then
        If JoinTable.HasDerived Then
            JoinCollection.Add " LEFT JOIN " & JoinTable.ToSqlString & " ON " & SourceTable & "." & SourceField & "=" & JoinTable.Alias & "." & JoinField
        Else
            Err.Raise VBA.vbObjectError + 3075, "SqlBuilder.LeftJoin", "子查询必须设置别名"
        End If
    Else
        Err.Raise VBA.vbObjectError + 3075, "SqlBuilder.LeftJoin", "LEFT JOIN 语句错误"
    End If
End Sub

Private Function ToSourceSqlString(Optional ByVal IsUpdate As Boolean = False) As String
    If HasSource Then
        Dim SourceSqlString As String
        If Not JoinCollection Is Nothing Then
            Dim length As Long
            length = JoinCollection.Count
            If length > 0 Then
                If length = 1 Then
                    SourceSqlString = FromTableName & JoinCollection(1)
                Else
                    SourceSqlString = "(" & FromTableName & JoinCollection(1) & ")"
                    Dim i As Long
                    For i = 2 To length
                        If i < length Then
                            SourceSqlString = "(" & SourceSqlString & JoinCollection(i) & ")"
                        Else
                            SourceSqlString = SourceSqlString & JoinCollection(i)
                        End If
                    Next i
                End If
            Else
                SourceSqlString = FromTableName
            End If
        Else
            SourceSqlString = FromTableName
        End If

        If IsUpdate = False Then SourceSqlString = " FROM " & SourceSqlString
        ToSourceSqlString = SourceSqlString
    Else
        Err.Raise VBA.vbObjectError + 3078, "SqlBuilder.ToSourceSqlString", "SQL 语法错误"
    End If
End Function


'字段
Public Sub Field(ByVal FieldName As String, Optional ByRef value As Variant, Optional ByVal Alias As String)
    If StringBase.IsWhiteSpace(FieldName) Then Exit Sub
    If FieldCollection Is Nothing Then Set FieldCollection = New Collection
    If Not IsMissing(Alias) And Not StringBase.IsWhiteSpace(Alias) Then
        FieldCollection.Add FieldName & " AS " & Alias
    Else
        FieldCollection.Add FieldName
    End If
    If Not IsMissing(value) Then
        If ValueCollection Is Nothing Then Set ValueCollection = New Collection
        Select Case VBA.VarType(value)
            Case VBA.vbString
                If Left(value, 4) = "[$$]" Then
                    ValueCollection.Add Mid(value, 5)
                Else
                    ValueCollection.Add "'" & Replace(value, "'", "''") & "'"
                End If
            Case VBA.vbDate
                ValueCollection.Add "#" & Format(value, "yyyy-mm-dd") & "#"
            Case VBA.vbNull
                ValueCollection.Add "NULL"
            Case Else
                ValueCollection.Add value
        End Select
    End If
End Sub

Private Function ToFieldSqlString(Optional ByVal IsUpdate As Boolean = False) As String
    If HasField Then
        Dim FieldSqlString As String
        If SelectAllField And HasSource Then
            FieldSqlString = FromTableName & ".*"
        ElseIf Not FieldCollection Is Nothing Then
            If FieldCollection.Count > 0 Then
                Dim i As Long
                Dim length As Long
                length = FieldCollection.Count
                For i = 1 To length
                    If i < length Then
                        If IsUpdate = True And Validation Then
                            FieldSqlString = FieldSqlString & FieldCollection(i) & "=" & ValueCollection(i) & ", "
                        Else
                            FieldSqlString = FieldSqlString & FieldCollection(i) & ", "
                        End If
                    ElseIf IsUpdate = True And Validation Then
                        FieldSqlString = FieldSqlString & FieldCollection(i) & "=" & ValueCollection(i)
                    Else
                        FieldSqlString = FieldSqlString & FieldCollection(i)
                    End If
                Next i
            End If
        End If
        ToFieldSqlString = FieldSqlString
    Else
        Err.Raise VBA.vbObjectError + 3085, "SqlBuilder.ToFieldSqlString", "查询字段未定义"
    End If
End Function

Private Function ToValueSqlString(Optional ByVal IsUpdate As Boolean = False) As String
    If Not ValueCollection Is Nothing And ValueCollection.Count > 0 Then
        Dim i As Long
        Dim length As Long
        Dim ValueSqlString As String
        length = ValueCollection.Count
        For i = 1 To length
            If i < length Then
                ValueSqlString = ValueSqlString & ValueCollection(i) & ", "
            Else
                ValueSqlString = ValueSqlString & ValueCollection(i)
            End If
        Next i
        ToValueSqlString = ValueSqlString
    Else
        Err.Raise VBA.vbObjectError + 3155, "SqlBuilder.ToValueSqlString", "字段值未定义"
    End If
End Function

' 筛选
Public Sub Where(ByVal Condition As String, Optional ByRef value As Variant)
    If StringBase.IsWhiteSpace(Condition) Then Exit Sub
    If WhereCollection Is Nothing Then Set WhereCollection = New Collection
    If Not IsMissing(value) Then
        Select Case VBA.VarType(value)
            Case VBA.vbString
                If Left(value, 4) = "[$$]" Then
                    WhereCollection.Add Condition & " " & Mid(value, 5)
                Else
                    WhereCollection.Add Condition & " '" & Replace(value, "'", "''") & "'"
                End If
            Case VBA.vbDate
                WhereCollection.Add Condition & " #" & Format(value, "yyyy-mm-dd") & "#"
            Case VBA.vbNull
                WhereCollection.Add Condition & " NULL"
            Case Else
                WhereCollection.Add Condition & " " & value
        End Select
    Else
        WhereCollection.Add Condition
    End If
End Sub

Private Function ToWhereSqlString() As String
    If Not WhereCollection Is Nothing Then
        If WhereCollection.Count > 0 Then
            Dim i As Long
            Dim length As Long
            Dim WhereSqlString As String
            length = WhereCollection.Count
            For i = 1 To length
                If i < length Then
                    WhereSqlString = WhereSqlString & "(" & WhereCollection(i) & ") And "
                Else
                    WhereSqlString = WhereSqlString & "(" & WhereCollection(i) & ")"
                End If
            Next i
            If length > 1 Then WhereSqlString = "(" & WhereSqlString & ")"
            If Not StringBase.IsWhiteSpace(WhereSqlString) Then WhereSqlString = " WHERE " & WhereSqlString
            ToWhereSqlString = WhereSqlString
        End If
    End If
    Exit Function
End Function


' 参数
Public Sub Param(ByVal ParamName As String, ByVal value As Variant, Optional ByVal ParamType As String)
    ParamName = Replace(ParamName, " ", "_")
    If StringBase.IsWhiteSpace(ParamName) Then Exit Sub
    If ParamDefCollection Is Nothing Then Set ParamDefCollection = New Collection
    If ParamCollection Is Nothing Then Set ParamCollection = New Collection
    If ParamTypeCollection Is Nothing Then Set ParamTypeCollection = New Collection
    ParamName = Replace(ParamName, ".", "_")
    ParamName = Replace(ParamName, "!", "_")
    ParamDefCollection.Add "Param_" & ParamName
    ParamCollection.Add value

    If IsMissing(ParamType) Or StringBase.IsWhiteSpace(ParamType) Then
        Select Case VarType(value)
            Case 2
                ParamTypeCollection.Add "Short"
            Case 3
                ParamTypeCollection.Add "Long"
            Case 4
                ParamTypeCollection.Add "Single"
            Case 5
                ParamTypeCollection.Add "Double"
            Case 6
                ParamTypeCollection.Add "Currency"
            Case 7
                ParamTypeCollection.Add "DateTime"
            Case 8
                ParamTypeCollection.Add "Text(255)"
            Case 11
                ParamTypeCollection.Add "Bit"
            Case 14
                ParamTypeCollection.Add "Decimal"
            Case 17
                ParamTypeCollection.Add "Byte"
            Case 20
                ParamTypeCollection.Add "LongLong"
            Case Else
                ParamTypeCollection.Add "Text(255)"
        End Select
    Else
        ParamTypeCollection.Add ParamType
    End If
End Sub

Public Sub SetQueryDef(ByRef Def As DAO.QueryDef)
    If HasParam Then
        Dim i As Long
        Dim length As Long
        length = ParamDefCollection.Count
        For i = 1 To length
            Def.Parameters(ParamDefCollection(i)) = ParamCollection(i)
        Next i
        Set ParamDefCollection = Nothing
        Set ParamTypeCollection = Nothing
        Set ParamCollection = Nothing
    End If
End Sub

Private Function ToParamDefSqlString() As String
    If HasParam Then
        If Not ParamDefCollection Is Nothing Then
            If ParamDefCollection.Count > 0 Then
                Dim i As Long
                Dim length As Long
                Dim ParamDefSqlString As String
                length = ParamDefCollection.Count
                For i = 1 To length
                    If i < length Then
                        ParamDefSqlString = ParamDefSqlString & "[" & ParamDefCollection(i) & "] " & ParamTypeCollection(i) & ", "
                    Else
                        ParamDefSqlString = ParamDefSqlString & "[" & ParamDefCollection(i) & "] " & ParamTypeCollection(i)
                    End If
                Next i
                If Not StringBase.IsWhiteSpace(ParamDefSqlString) Then ParamDefSqlString = "PARAMETERS " & ParamDefSqlString & "; "

                ToParamDefSqlString = ParamDefSqlString
            End If
        End If
    End If
End Function


' 分组
Public Sub Group(ByVal FieldName As String)
    If Not StringBase.IsWhiteSpace(FieldName) Then
        If GroupCollection Is Nothing Then Set GroupCollection = New Collection
        GroupCollection.Add FieldName
    End If
End Sub

Private Function ToGroupSqlString() As String
    If Not GroupCollection Is Nothing Then
        If GroupCollection.Count > 0 Then
            Dim i As Long
            Dim length As Long
            Dim GroupSqlString As String
            length = GroupCollection.Count
            For i = 1 To length
                If i < length Then
                    GroupSqlString = GroupSqlString & GroupCollection(i) & ", "
                Else
                    GroupSqlString = GroupSqlString & GroupCollection(i)
                End If
            Next i
            If Not StringBase.IsWhiteSpace(GroupSqlString) Then GroupSqlString = " GROUP BY " & GroupSqlString
            ToGroupSqlString = GroupSqlString
        End If
    End If
End Function

' 筛选
Public Sub Having(ByVal Condition As String, Optional ByRef value As Variant)
    If StringBase.IsWhiteSpace(Condition) Then Exit Sub
    If HavingCollection Is Nothing Then Set HavingCollection = New Collection
    If Not IsMissing(value) Then
        Select Case VBA.VarType(value)
            Case VBA.vbString
                If Left(value, 4) = "[$$]" Then
                    HavingCollection.Add Condition & " " & Mid(value, 5)
                Else
                    HavingCollection.Add Condition & " '" & Replace(value, "'", "''") & "'"
                End If
            Case VBA.vbDate
                HavingCollection.Add Condition & " #" & Format(value, "yyyy-mm-dd") & "#"
            Case VBA.vbNull
                HavingCollection.Add Condition & " NULL"
            Case Else
                HavingCollection.Add Condition & " " & value
        End Select
    Else
        HavingCollection.Add Condition
    End If
End Sub

Private Function ToHavingSqlString() As String
    If Not HavingCollection Is Nothing Then
        If HavingCollection.Count > 0 Then
            Dim i As Long
            Dim length As Long
            Dim HavingSqlString As String
            length = HavingCollection.Count
            For i = 1 To length
                If i < length Then
                    HavingSqlString = HavingSqlString & "(" & HavingCollection(i) & ") And "
                Else
                    HavingSqlString = HavingSqlString & "(" & HavingCollection(i) & ")"
                End If
            Next i
            If length > 1 Then HavingSqlString = "(" & HavingSqlString & ")"
            If Not StringBase.IsWhiteSpace(HavingSqlString) Then HavingSqlString = " HAVING " & HavingSqlString
            ToHavingSqlString = HavingSqlString
        End If
    End If
End Function

'排序
Public Sub Order(ByVal FieldName As String, Optional ByVal Reverse As Boolean = False)
    If StringBase.IsWhiteSpace(FieldName) Then Exit Sub
    If OrderCollection Is Nothing Then Set OrderCollection = New Collection
    If Reverse = True Then
        OrderCollection.Add FieldName & " DESC"
    Else
        OrderCollection.Add FieldName
    End If
End Sub

Private Function ToOrderSqlString() As String
    If Not OrderCollection Is Nothing Then
        If OrderCollection.Count > 0 Then
            Dim i As Long
            Dim length As Long
            Dim OrderSqlString As String
            length = OrderCollection.Count
            For i = 1 To length
                If i < length Then
                    OrderSqlString = OrderSqlString & OrderCollection(i) & ", "
                Else
                    OrderSqlString = OrderSqlString & OrderCollection(i)
                End If
            Next i
            If Not StringBase.IsWhiteSpace(OrderSqlString) Then OrderSqlString = " ORDER BY " & OrderSqlString
            ToOrderSqlString = OrderSqlString
        End If
    End If
End Function


' 拼接sql字符串
Public Function ToSqlString(Optional ByVal Operation As Byte = 0, Optional ByVal Semicolon As Boolean = True) As String
    Dim EndSemicolon As String
    If Semicolon Then EndSemicolon = ";"
    Dim SqlString As String
    If IsDerivedTable Then
        If HasField And HasSource Then
            SqlString = "(" & ToParamDefSqlString & "SELECT " & DistinctSqlString & TopSqlString & ToFieldSqlString & ToSourceSqlString & ToWhereSqlString & ToGroupSqlString & ToHavingSqlString & ToOrderSqlString & ") AS " & DerivedTableName
        End If
    Else
        '新增
        If Operation = 4 Then
            If StringBase.IsWhiteSpace(IntoTableName) Then
                Err.Raise VBA.vbObjectError + 3155, "SqlBuilder.ToSqlString", "INSERT INTO 语句错误"
                Exit Function
            End If
            If Validation Then
                If HasSource Then
                    SqlString = ToParamDefSqlString & "INSERT INTO " & IntoTableName & " (" & ToFieldSqlString & ") SELECT " & ToValueSqlString & ToSourceSqlString & EndSemicolon
                Else
                    SqlString = ToParamDefSqlString & "INSERT INTO " & IntoTableName & " (" & ToFieldSqlString & ") VALUES (" & ToValueSqlString & ")" & EndSemicolon
                End If
            Else
                Err.Raise VBA.vbObjectError + 3155, "SqlBuilder.ToSqlString", "INSERT 语句错误"
                Exit Function
            End If
        '更改
        ElseIf Operation = 3 Then
            If HasSource Then
                If Validation Then
                    SqlString = ToParamDefSqlString & "UPDATE " & ToSourceSqlString(True) & " SET " & ToFieldSqlString(True) & ToWhereSqlString & EndSemicolon
                Else
                    Err.Raise VBA.vbObjectError + 3085, "SqlBuilder.ToSqlString", "UPDATE 字段未定义"
                    Exit Function
                End If
            Else
                Err.Raise VBA.vbObjectError + 3144, "SqlBuilder.ToSqlString", "UPDATE 语句错误"
                Exit Function
            End If

            If HasSource And Validation Then SqlString = ToParamDefSqlString & "UPDATE " & ToSourceSqlString(True) & " SET " & ToFieldSqlString(True) & ToWhereSqlString & EndSemicolon
        '删除
        ElseIf Operation = 2 Then
            If HasSource Then
                If HasWhere Then
                    SqlString = ToParamDefSqlString & "DELETE" & ToSourceSqlString & ToWhereSqlString & EndSemicolon
                Else
                    Err.Raise VBA.vbObjectError + 3075, "SqlBuilder.ToSqlString", "DELETE 必须定义 WHERE 语句"
                    Exit Function
                End If
            Else
                Err.Raise VBA.vbObjectError + 3078, "SqlBuilder.ToSqlString", "FROM 语句错误"
                Exit Function
            End If
        '计数
        ElseIf Operation = 1 Then
            If HasSource Then
                SqlString = ToParamDefSqlString & "SELECT COUNT(*)" & ToSourceSqlString & ToWhereSqlString & EndSemicolon
            Else
                Err.Raise VBA.vbObjectError + 3078, "SqlBuilder.ToSqlString", "FROM 语句错误"
                Exit Function
            End If
        '查询
        ElseIf HasField And HasSource Then
            SqlString = ToParamDefSqlString & "SELECT " & DistinctSqlString & TopSqlString & ToFieldSqlString & ToSourceSqlString & ToWhereSqlString & ToGroupSqlString & ToHavingSqlString & ToOrderSqlString & EndSemicolon
        End If
    End If

    Set FieldCollection = Nothing
    Set ValueCollection = Nothing
    Set JoinCollection = Nothing
    Set WhereCollection = Nothing
    Set GroupCollection = Nothing
    Set HavingCollection = Nothing
    Set OrderCollection = Nothing

    'MsgBox SqlString
    If VBA.Len(SqlString) > 0 Then
        ToSqlString = SqlString
    Else
        Err.Raise 3075, "SqlBuilder.ToSqlString", "SQL 语法错误"
    End If
End Function
