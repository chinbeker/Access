VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit


Private TopSqlString As String
Private SelectAllField As Boolean
Private DistinctSqlString As String
Private FromTableName As String
Private IntoTableName As String

Private FieldCollection As Collection
Private ValueCollection As Collection
Private InnerCollection As Collection
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
Private Function HasdInto() As Boolean
    HasdInto = Not StringBase.IsWhiteSpace(IntoTableName)
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
Public Function HasParam() As Boolean
    If ParamDefCollection Is Nothing Or ParamTypeCollection Is Nothing Or ParamCollection Is Nothing Then
        HasParam = False
    Else
        HasParam = (ParamDefCollection.Count > 0 And ParamDefCollection.Count = ParamTypeCollection.Count And ParamDefCollection.Count = ParamCollection.Count)
    End If
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


' 表
Public Sub From(ByVal TableName As String)
    If Not StringBase.IsWhiteSpace(TableName) Then FromTableName = TableName
End Sub
Public Sub Into(ByVal TableName As String)
    If Not StringBase.IsWhiteSpace(TableName) Then IntoTableName = TableName
End Sub

' 表连接
Public Sub Inner(ByVal InnerTable As String, ByVal TargetTable As String, ByVal InnerJoinCondition As String, Optional ByVal JoinCondition As String)
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(InnerTable) Then Exit Sub
    If StringBase.IsWhiteSpace(TargetTable) Then Exit Sub
    If StringBase.IsWhiteSpace(InnerJoinCondition) Then Exit Sub
    If IsMissing(JoinCondition) Or StringBase.IsWhiteSpace(JoinCondition) Then JoinCondition = InnerJoinCondition
    If InnerCollection Is Nothing Then Set InnerCollection = New Collection
    InnerCollection.Add " INNER JOIN " & InnerTable & " ON " & TargetTable & "." & JoinCondition & "=" & InnerTable & "." & InnerJoinCondition
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

Private Function ToSourceSqlString(Optional ByVal IsUpdate As Boolean = False) As String
    On Error GoTo ErrorHandler
    If HasSource Then
        Dim SourceSqlString As String
        If Not InnerCollection Is Nothing Then
            Dim length As Long
            length = InnerCollection.Count
            If length > 0 Then
                If length = 1 Then
                    SourceSqlString = FromTableName & InnerCollection(1)
                Else
                    SourceSqlString = "(" & FromTableName & InnerCollection(1) & ")"
                    Dim i As Integer
                    For i = 2 To length
                        If i < length Then
                            SourceSqlString = "(" & SourceSqlString & InnerCollection(i) & ")"
                        Else
                            SourceSqlString = SourceSqlString & InnerCollection(i)
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
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function


'字段
Public Sub Field(ByVal FieldName As String, Optional ByRef value As Variant)
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(FieldName) Then Exit Sub
    If FieldCollection Is Nothing Then Set FieldCollection = New Collection
    FieldCollection.Add FieldName

    If Not IsMissing(value) Then
        If ValueCollection Is Nothing Then Set ValueCollection = New Collection
        If Not StringBase.IsWhiteSpace(value) Then
            If Left(value, 4) = "[$$]" Then
                ValueCollection.Add Mid(value, 5)
            Else
                ValueCollection.Add "'" & value & "'"
            End If
        ElseIf IsDate(value) Then
            ValueCollection.Add "#" & Format(value, "yyyy-mm-dd") & "#"
        Else
            ValueCollection.Add value
        End If
    End If
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

Private Function ToFieldSqlString(Optional ByVal IsUpdate As Boolean = False) As String
    On Error GoTo ErrorHandler
    If HasField Then
        Dim FieldSqlString As String
        If SelectAllField And HasSource Then
            FieldSqlString = FromTableName & ".*"
        ElseIf Not FieldCollection Is Nothing Then
            If FieldCollection.Count > 0 Then
                Dim i As Integer
                Dim length As Integer
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
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

Private Function ToValueSqlString(Optional ByVal IsUpdate As Boolean = False) As String
    On Error GoTo ErrorHandler
    If Not ValueCollection Is Nothing Then
        If ValueCollection.Count > 0 Then
            Dim i As Integer
            Dim length As Integer
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
        End If
    End If
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 筛选
Public Sub Where(ByVal Condition As String, Optional ByRef value As Variant)
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(Condition) Then Exit Sub
    If WhereCollection Is Nothing Then Set WhereCollection = New Collection
    If Not IsMissing(value) Then
        If Not StringBase.IsWhiteSpace(value) Then
            If Left(value, 4) = "[$$]" Then
                WhereCollection.Add Condition & Mid(value, 5)
            Else
                WhereCollection.Add Condition & "'" & value & "'"
            End If
        ElseIf IsDate(value) Then
            WhereCollection.Add Condition & "#" & Format(value, "yyyy-mm-dd") & "#"
        Else
            WhereCollection.Add Condition & value
        End If
    Else
        WhereCollection.Add Condition
    End If
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

Private Function ToWhereSqlString() As String
    On Error GoTo ErrorHandler
    If Not WhereCollection Is Nothing Then
        If WhereCollection.Count > 0 Then
            Dim i As Integer
            Dim length As Integer
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

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function


' 参数
Public Sub Param(ByVal ParamName As String, ByVal value As Variant, Optional ByVal ParamType As String)
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(ParamName) Then Exit Sub
    If ParamDefCollection Is Nothing Then Set ParamDefCollection = New Collection
    If ParamCollection Is Nothing Then Set ParamCollection = New Collection
    If ParamTypeCollection Is Nothing Then Set ParamTypeCollection = New Collection
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

    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub
Public Sub SetParam(ByRef Def As DAO.QueryDef)
    On Error GoTo ErrorHandler
    If HasParam Then
        Dim i As Integer
        Dim length As Integer
        length = ParamDefCollection.Count
        For i = 1 To length
            Def.Parameters(ParamDefCollection(i)) = ParamCollection(i)
        Next i
        Set ParamDefCollection = Nothing
        Set ParamTypeCollection = Nothing
        Set ParamCollection = Nothing
    End If
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

Private Function ToParamDefSqlString() As String
    On Error GoTo ErrorHandler
    If HasParam Then
        If Not ParamDefCollection Is Nothing Then
            If ParamDefCollection.Count > 0 Then
                Dim i As Integer
                Dim length As Integer
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
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function


' 分组
Public Sub Group(ByVal FieldName As String)
    On Error GoTo ErrorHandler
    If Not StringBase.IsWhiteSpace(FieldName) Then
        If GroupCollection Is Nothing Then Set GroupCollection = New Collection
        GroupCollection.Add FieldName
    End If
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

Private Function ToGroupSqlString() As String
    On Error GoTo ErrorHandler
    If Not GroupCollection Is Nothing Then
        If GroupCollection.Count > 0 Then
            Dim i As Integer
            Dim length As Integer
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
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

' 筛选
Public Sub Having(ByVal Condition As String, Optional ByRef value As Variant)
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(Condition) Then Exit Sub
    If HavingCollection Is Nothing Then Set HavingCollection = New Collection
    If Not IsMissing(value) Then
        If Not StringBase.IsWhiteSpace(value) Then
            If Left(value, 4) = "[$$]" Then
                HavingCollection.Add Condition & Mid(value, 5)
            Else
                HavingCollection.Add Condition & "'" & value & "'"
            End If
        ElseIf IsDate(value) Then
            HavingCollection.Add Condition & "#" & Format(value, "yyyy-mm-dd") & "#"
        Else
            HavingCollection.Add Condition & value
        End If
    Else
        HavingCollection.Add Condition
    End If
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

Private Function ToHavingSqlString() As String
    On Error GoTo ErrorHandler
    If Not HavingCollection Is Nothing Then
        If HavingCollection.Count > 0 Then
            Dim i As Integer
            Dim length As Integer
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
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function

'排序
Public Sub Order(ByVal FieldName As String, Optional ByVal Reverse As Boolean = False)
    On Error GoTo ErrorHandler
    If StringBase.IsWhiteSpace(FieldName) Then Exit Sub
    If OrderCollection Is Nothing Then Set OrderCollection = New Collection
    If Reverse = True Then
        OrderCollection.Add FieldName & " DESC"
    Else
        OrderCollection.Add FieldName
    End If
    Exit Sub

ErrorHandler:
    Call Message.Error(Err)
    Exit Sub
End Sub

Private Function ToOrderSqlString() As String
    On Error GoTo ErrorHandler
    If Not OrderCollection Is Nothing Then
        If OrderCollection.Count > 0 Then
            Dim i As Integer
            Dim length As Integer
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
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function


' 拼接sql字符串
Public Function ToSqlString(Optional ByVal Operation As Byte = 0) As String
    On Error GoTo ErrorHandler
    Dim SqlString As String

    '新增
    If Operation = 4 Then
        If HasdInto And Validation Then
            SqlString = ToParamDefSqlString & "INSERT INTO " & IntoTableName & " (" & ToFieldSqlString & ") SELECT " & ToValueSqlString & ToSourceSqlString & ";"
        End If
    '更改
    ElseIf Operation = 3 Then
        If HasSource And Validation Then SqlString = ToParamDefSqlString & "UPDATE " & ToSourceSqlString(True) & " SET " & ToFieldSqlString(True) & ToWhereSqlString & ";"
    '删除
    ElseIf Operation = 2 Then
        If HasSource And HasWhere Then SqlString = ToParamDefSqlString & "DELETE" & ToSourceSqlString & ToWhereSqlString & ";"
    '计数
    ElseIf Operation = 1 Then
        If HasSource Then SqlString = ToParamDefSqlString & "SELECT " & DistinctSqlString & "COUNT(*)" & ToSourceSqlString & ToWhereSqlString & ";"
    '查询
    ElseIf HasField And HasSource Then
        SqlString = ToParamDefSqlString & "SELECT " & DistinctSqlString & TopSqlString & ToFieldSqlString & ToSourceSqlString & ToWhereSqlString & ToGroupSqlString & ToHavingSqlString & ToOrderSqlString & ";"
    End If

    Set FieldCollection = Nothing
    Set ValueCollection = Nothing
    Set InnerCollection = Nothing
    Set WhereCollection = Nothing
    Set GroupCollection = Nothing
    Set HavingCollection = Nothing
    Set OrderCollection = Nothing

    'MsgBox SqlString
    ToSqlString = SqlString
    Exit Function

ErrorHandler:
    Call Message.Error(Err)
    Exit Function
End Function
