Attribute VB_Name = "JSON"
Option Compare Database
Option Explicit

Private Type JSON_Options
    UseDoubleForLargeNumbers As Boolean
    AllowUnquotedKeys As Boolean
    EscapeSolidus As Boolean
End Type


Public Options As JSON_Options

'私有方法
'JSON_ParseObject
Private Function JSON_ParseObject(JSON_String As String, ByRef JSON_Index As Long) As Dictionary
    Dim JSON_Key As String
    Dim JSON_NextChar As String

    Set JSON_ParseObject = New Dictionary
    JSON_SkipSpaces JSON_String, JSON_Index

    If VBA.Mid$(JSON_String, JSON_Index, 1) <> "{" Then
        err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting '{'")
    Else
        JSON_Index = JSON_Index + 1

        Do
            JSON_SkipSpaces JSON_String, JSON_Index

            If VBA.Mid$(JSON_String, JSON_Index, 1) = "}" Then
                JSON_Index = JSON_Index + 1
                Exit Function
            ElseIf VBA.Mid$(JSON_String, JSON_Index, 1) = "," Then
                JSON_Index = JSON_Index + 1
                JSON_SkipSpaces JSON_String, JSON_Index
            End If

            JSON_Key = JSON_ParseKey(JSON_String, JSON_Index)
            JSON_NextChar = JSON_Peek(JSON_String, JSON_Index)

            If JSON_NextChar = "[" Or JSON_NextChar = "{" Then
                Set JSON_ParseObject.Item(JSON_Key) = JSON_ParseValue(JSON_String, JSON_Index)
            Else
                JSON_ParseObject.Item(JSON_Key) = JSON_ParseValue(JSON_String, JSON_Index)
            End If
        Loop
    End If
End Function

'JSON_ParseArray
Private Function JSON_ParseArray(JSON_String As String, ByRef JSON_Index As Long) As Collection
    Set JSON_ParseArray = New Collection

    JSON_SkipSpaces JSON_String, JSON_Index
    If VBA.Mid$(JSON_String, JSON_Index, 1) <> "[" Then
        err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting '['")
    Else
        JSON_Index = JSON_Index + 1

        Do
            JSON_SkipSpaces JSON_String, JSON_Index
            If VBA.Mid$(JSON_String, JSON_Index, 1) = "]" Then
                JSON_Index = JSON_Index + 1
                Exit Function
            ElseIf VBA.Mid$(JSON_String, JSON_Index, 1) = "," Then
                JSON_Index = JSON_Index + 1
                JSON_SkipSpaces JSON_String, JSON_Index
            End If

            JSON_ParseArray.Add JSON_ParseValue(JSON_String, JSON_Index)
        Loop
    End If
End Function

'JSON_ParseValue
Private Function JSON_ParseValue(JSON_String As String, ByRef JSON_Index As Long) As Variant
    JSON_SkipSpaces JSON_String, JSON_Index
    Select Case VBA.Mid$(JSON_String, JSON_Index, 1)
        Case "{"
            Set JSON_ParseValue = JSON_ParseObject(JSON_String, JSON_Index)
        Case "["
            Set JSON_ParseValue = JSON_ParseArray(JSON_String, JSON_Index)
        Case """", "'"
            JSON_ParseValue = JSON_ParseString(JSON_String, JSON_Index)
        Case Else
            If VBA.Mid$(JSON_String, JSON_Index, 4) = "true" Then
                JSON_ParseValue = True
                JSON_Index = JSON_Index + 4
            ElseIf VBA.Mid$(JSON_String, JSON_Index, 5) = "false" Then
                JSON_ParseValue = False
                JSON_Index = JSON_Index + 5
            ElseIf VBA.Mid$(JSON_String, JSON_Index, 4) = "null" Then
                JSON_ParseValue = Null
                JSON_Index = JSON_Index + 4
            ElseIf VBA.InStr("+-0123456789", VBA.Mid$(JSON_String, JSON_Index, 1)) Then
                JSON_ParseValue = JSON_ParseNumber(JSON_String, JSON_Index)
            Else
                err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
            End If
    End Select
End Function

'JSON_ParseString
Private Function JSON_ParseString(JSON_String As String, ByRef JSON_Index As Long) As String
    Dim JSON_Quote As String
    Dim JSON_Char As String
    Dim JSON_Code As String
    Dim JSON_Buffer As String
    Dim JSON_BufferPosition As Long
    Dim JSON_BufferLength As Long

    JSON_SkipSpaces JSON_String, JSON_Index

    JSON_Quote = VBA.Mid$(JSON_String, JSON_Index, 1)
    JSON_Index = JSON_Index + 1

    Do While JSON_Index > 0 And JSON_Index <= Len(JSON_String)
        JSON_Char = VBA.Mid$(JSON_String, JSON_Index, 1)

        Select Case JSON_Char
            Case "\"
                JSON_Index = JSON_Index + 1
                JSON_Char = VBA.Mid$(JSON_String, JSON_Index, 1)

                Select Case JSON_Char
                    Case """", "\", "/", "'"
                        JSON_BufferAppend JSON_Buffer, JSON_Char, JSON_BufferPosition, JSON_BufferLength
                        JSON_Index = JSON_Index + 1
                    Case "b"
                        JSON_BufferAppend JSON_Buffer, vbBack, JSON_BufferPosition, JSON_BufferLength
                        JSON_Index = JSON_Index + 1
                    Case "f"
                        JSON_BufferAppend JSON_Buffer, vbFormFeed, JSON_BufferPosition, JSON_BufferLength
                        JSON_Index = JSON_Index + 1
                    Case "n"
                        JSON_BufferAppend JSON_Buffer, vbCrLf, JSON_BufferPosition, JSON_BufferLength
                        JSON_Index = JSON_Index + 1
                    Case "r"
                        JSON_BufferAppend JSON_Buffer, vbCr, JSON_BufferPosition, JSON_BufferLength
                        JSON_Index = JSON_Index + 1
                    Case "t"
                        JSON_BufferAppend JSON_Buffer, vbTab, JSON_BufferPosition, JSON_BufferLength
                        JSON_Index = JSON_Index + 1
                    Case "u"
                        JSON_Index = JSON_Index + 1
                        JSON_Code = VBA.Mid$(JSON_String, JSON_Index, 4)
                        JSON_BufferAppend JSON_Buffer, VBA.ChrW(VBA.val("&h" + JSON_Code)), JSON_BufferPosition, JSON_BufferLength
                        JSON_Index = JSON_Index + 4
                End Select

            Case JSON_Quote
                JSON_ParseString = JSON_BufferToString(JSON_Buffer, JSON_BufferPosition)
                JSON_Index = JSON_Index + 1
                Exit Function

            Case Else
                JSON_BufferAppend JSON_Buffer, JSON_Char, JSON_BufferPosition, JSON_BufferLength
                JSON_Index = JSON_Index + 1
        End Select
    Loop
End Function

'JSON_ParseNumber
Private Function JSON_ParseNumber(JSON_String As String, ByRef JSON_Index As Long) As Variant
    Dim JSON_Char As String
    Dim JSON_Value As String
    Dim JSON_IsLargeNumber As Boolean

    JSON_SkipSpaces JSON_String, JSON_Index

    Do While JSON_Index > 0 And JSON_Index <= Len(JSON_String)
        JSON_Char = VBA.Mid$(JSON_String, JSON_Index, 1)

        If VBA.InStr("+-0123456789.eE", JSON_Char) Then
            JSON_Value = JSON_Value & JSON_Char
            JSON_Index = JSON_Index + 1
        Else
            JSON_IsLargeNumber = IIf(InStr(JSON_Value, "."), Len(JSON_Value) >= 17, Len(JSON_Value) >= 16)
            If Not JsonOptions.UseDoubleForLargeNumbers And JSON_IsLargeNumber Then
                JSON_ParseNumber = JSON_Value
            Else
                JSON_ParseNumber = VBA.val(JSON_Value)
            End If
            Exit Function
        End If
    Loop
End Function

'JSON_ParseKey
Private Function JSON_ParseKey(JSON_String As String, ByRef JSON_Index As Long) As String
    If VBA.Mid$(JSON_String, JSON_Index, 1) = """" Or VBA.Mid$(JSON_String, JSON_Index, 1) = "'" Then
        JSON_ParseKey = JSON_ParseString(JSON_String, JSON_Index)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Dim JSON_Char As String
        Do While JSON_Index > 0 And JSON_Index <= Len(JSON_String)
            JSON_Char = VBA.Mid$(JSON_String, JSON_Index, 1)
            If (JSON_Char <> " ") And (JSON_Char <> ":") Then
                JSON_ParseKey = JSON_ParseKey & JSON_Char
                JSON_Index = JSON_Index + 1
            Else
                Exit Do
            End If
        Loop
    Else
        err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting '""' or '''")
    End If

    JSON_SkipSpaces JSON_String, JSON_Index
    If VBA.Mid$(JSON_String, JSON_Index, 1) <> ":" Then
        err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JSON_String, JSON_Index, "Expecting ':'")
    Else
        JSON_Index = JSON_Index + 1
    End If
End Function

'JSON_IsUndefined
Private Function JSON_IsUndefined(ByVal JSON_Value As Variant) As Boolean
    Select Case VBA.VarType(JSON_Value)
        Case VBA.vbEmpty
            JSON_IsUndefined = True
        Case VBA.vbObject
            Select Case VBA.typeName(JSON_Value)
                Case "Empty", "Nothing"
                    JSON_IsUndefined = True
            End Select
    End Select
End Function

'JSON_Encode
Private Function JSON_Encode(ByVal JSON_Text As Variant) As String
    Dim JSON_Index As Long
    Dim JSON_Char As String
    Dim JSON_AscCode As Long
    Dim JSON_Buffer As String
    Dim JSON_BufferPosition As Long
    Dim JSON_BufferLength As Long

    For JSON_Index = 1 To VBA.Len(JSON_Text)
        JSON_Char = VBA.Mid$(JSON_Text, JSON_Index, 1)
        JSON_AscCode = VBA.AscW(JSON_Char)


        If JSON_AscCode < 0 Then
            JSON_AscCode = JSON_AscCode + 65536
        End If

        Select Case JSON_AscCode
            Case 34
                JSON_Char = "\"""
            Case 92
                JSON_Char = "\\"
            Case 47
                If JsonOptions.EscapeSolidus Then
                    JSON_Char = "\/"
                End If
            Case 8
                JSON_Char = "\b"
            Case 12
                JSON_Char = "\f"
            Case 10
                JSON_Char = "\n"
            Case 13
                JSON_Char = "\r"
            Case 9
                JSON_Char = "\t"
            Case 0 To 31, 127 To 65535
                JSON_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(JSON_AscCode), 4)
            End Select

        JSON_BufferAppend JSON_Buffer, JSON_Char, JSON_BufferPosition, JSON_BufferLength
    Next JSON_Index

    JSON_Encode = JSON_BufferToString(JSON_Buffer, JSON_BufferPosition)
End Function

'JSON_Peek
Private Function JSON_Peek(JSON_String As String, ByVal JSON_Index As Long, Optional JSON_NumberOfCharacters As Long = 1) As String
    JSON_SkipSpaces JSON_String, JSON_Index
    JSON_Peek = VBA.Mid$(JSON_String, JSON_Index, JSON_NumberOfCharacters)
End Function

'JSON_SkipSpaces
Private Sub JSON_SkipSpaces(JSON_String As String, ByRef JSON_Index As Long)
    Do While JSON_Index > 0 And JSON_Index <= VBA.Len(JSON_String) And VBA.Mid$(JSON_String, JSON_Index, 1) = " "
        JSON_Index = JSON_Index + 1
    Loop
End Sub

'JSON_StringIsLargeNumber
Private Function JSON_StringIsLargeNumber(JSON_String As Variant) As Boolean
    Dim JSON_Length As Long
    Dim JSON_CharIndex As Long
    JSON_Length = VBA.Len(JSON_String)

    If JSON_Length >= 16 And JSON_Length <= 100 Then
        Dim JSON_CharCode As String

        JSON_StringIsLargeNumber = True

        For JSON_CharIndex = 1 To JSON_Length
            JSON_CharCode = VBA.Asc(VBA.Mid$(JSON_String, JSON_CharIndex, 1))
            Select Case JSON_CharCode
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                JSON_StringIsLargeNumber = False
                Exit Function
            End Select
        Next JSON_CharIndex
    End If
End Function

'JSON_ParseErrorMessage
Private Function JSON_ParseErrorMessage(JSON_String As String, ByRef JSON_Index As Long, ErrorMessage As String)
    Dim JSON_StartIndex As Long
    Dim JSON_StopIndex As Long

    JSON_StartIndex = JSON_Index - 10
    JSON_StopIndex = JSON_Index + 10
    If JSON_StartIndex <= 0 Then
        JSON_StartIndex = 1
    End If
    If JSON_StopIndex > VBA.Len(JSON_String) Then
        JSON_StopIndex = VBA.Len(JSON_String)
    End If

    JSON_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(JSON_String, JSON_StartIndex, JSON_StopIndex - JSON_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(JSON_Index - JSON_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

'JSON_BufferAppend
Private Sub JSON_BufferAppend(ByRef JSON_Buffer As String, ByRef JSON_Append As Variant, ByRef JSON_BufferPosition As Long, ByRef JSON_BufferLength As Long)
    Dim JSON_AppendLength As Long
    Dim JSON_LengthPlusPosition As Long

    JSON_AppendLength = VBA.Len(JSON_Append)
    JSON_LengthPlusPosition = JSON_AppendLength + JSON_BufferPosition

    If JSON_LengthPlusPosition > JSON_BufferLength Then

        Dim JSON_AddedLength As Long
        JSON_AddedLength = IIf(JSON_AppendLength > JSON_BufferLength, JSON_AppendLength, JSON_BufferLength)

        JSON_Buffer = JSON_Buffer & VBA.Space$(JSON_AddedLength)
        JSON_BufferLength = JSON_BufferLength + JSON_AddedLength
    End If

    Mid$(JSON_Buffer, JSON_BufferPosition + 1, JSON_AppendLength) = CStr(JSON_Append)
    JSON_BufferPosition = JSON_BufferPosition + JSON_AppendLength
End Sub

'JSON_BufferToString
Private Function JSON_BufferToString(ByRef JSON_Buffer As String, ByVal JSON_BufferPosition As Long) As String
    If JSON_BufferPosition > 0 Then
        JSON_BufferToString = VBA.Left$(JSON_Buffer, JSON_BufferPosition)
    End If
End Function




'解析 JSON 字符串为字典
Public Function Parse(ByVal JsonString As String) As Object
    Dim JSON_Index As Long
    JSON_Index = 1

    JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")

    JSON_SkipSpaces JsonString, JSON_Index
    Select Case VBA.Mid$(JsonString, JSON_Index, 1)
        Case "{"
            Set Parse = JSON_ParseObject(JsonString, JSON_Index)
        Case "["
            Set Parse = JSON_ParseArray(JsonString, JSON_Index)
        Case Else
            err.Raise 10001, "JSONConverter", JSON_ParseErrorMessage(JsonString, JSON_Index, "Expecting '{' or '['")
    End Select
End Function

'转换为 JSON 字符串
Public Function Stringify(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal JSON_CurrentIndentation As Long = 0) As String
    Dim JSON_Buffer As String
    Dim JSON_BufferPosition As Long
    Dim JSON_BufferLength As Long
    Dim JSON_Index As Long
    Dim JSON_LBound As Long
    Dim JSON_UBound As Long
    Dim JSON_IsFirstItem As Boolean
    Dim JSON_Index2D As Long
    Dim JSON_LBound2D As Long
    Dim JSON_UBound2D As Long
    Dim JSON_IsFirstItem2D As Boolean
    Dim JSON_Key As Variant
    Dim JSON_Value As Variant
    Dim JSON_DateStr As String
    Dim JSON_Converted As String
    Dim JSON_SkipItem As Boolean
    Dim JSON_PrettyPrint As Boolean
    Dim JSON_Indentation As String
    Dim JSON_InnerIndentation As String

    JSON_LBound = -1
    JSON_UBound = -1
    JSON_IsFirstItem = True
    JSON_LBound2D = -1
    JSON_UBound2D = -1
    JSON_IsFirstItem2D = True
    JSON_PrettyPrint = Not IsMissing(Whitespace)

    Select Case VBA.VarType(JsonValue)
        Case VBA.vbNull
            Stringify = "null"
        Case VBA.vbDate
            ' Date
            JSON_DateStr = DateBase.ConvertToIso(VBA.CDate(JsonValue))
            Stringify = """" & JSON_DateStr & """"
        Case VBA.vbString
            If Not Options.UseDoubleForLargeNumbers And JSON_StringIsLargeNumber(JsonValue) Then
                Stringify = JsonValue
            Else
                Stringify = """" & JSON_Encode(JsonValue) & """"
            End If
        Case VBA.vbBoolean
            If JsonValue Then
                Stringify = "true"
            Else
                Stringify = "false"
            End If

        Case VBA.vbArray To VBA.vbArray + VBA.vbByte

            If JSON_PrettyPrint Then
                If VBA.VarType(Whitespace) = VBA.vbString Then
                    JSON_Indentation = VBA.String$(JSON_CurrentIndentation + 1, Whitespace)
                    JSON_InnerIndentation = VBA.String$(JSON_CurrentIndentation + 2, Whitespace)
                Else
                    JSON_Indentation = VBA.Space$((JSON_CurrentIndentation + 1) * Whitespace)
                    JSON_InnerIndentation = VBA.Space$((JSON_CurrentIndentation + 2) * Whitespace)
                End If
            End If

            ' Array
            JSON_BufferAppend JSON_Buffer, "[", JSON_BufferPosition, JSON_BufferLength

            On Error Resume Next

            JSON_LBound = LBound(JsonValue, 1)
            JSON_UBound = UBound(JsonValue, 1)
            JSON_LBound2D = LBound(JsonValue, 2)
            JSON_UBound2D = UBound(JsonValue, 2)

            If JSON_LBound >= 0 And JSON_UBound >= 0 Then
                For JSON_Index = JSON_LBound To JSON_UBound
                    If JSON_IsFirstItem Then
                        JSON_IsFirstItem = False
                    Else
                        JSON_BufferAppend JSON_Buffer, ",", JSON_BufferPosition, JSON_BufferLength
                    End If

                    If JSON_LBound2D >= 0 And JSON_UBound2D >= 0 Then
                        If JSON_PrettyPrint Then
                            JSON_BufferAppend JSON_Buffer, vbNewLine, JSON_BufferPosition, JSON_BufferLength
                        End If
                        JSON_BufferAppend JSON_Buffer, JSON_Indentation & "[", JSON_BufferPosition, JSON_BufferLength

                        For JSON_Index2D = JSON_LBound2D To JSON_UBound2D
                            If JSON_IsFirstItem2D Then
                                JSON_IsFirstItem2D = False
                            Else
                                JSON_BufferAppend JSON_Buffer, ",", JSON_BufferPosition, JSON_BufferLength
                            End If

                            JSON_Converted = Stringify(JsonValue(JSON_Index, JSON_Index2D), Whitespace, JSON_CurrentIndentation + 2)

                            If JSON_Converted = "" Then
                                If JSON_IsUndefined(JsonValue(JSON_Index, JSON_Index2D)) Then
                                    JSON_Converted = "null"
                                End If
                            End If

                            If JSON_PrettyPrint Then
                                JSON_Converted = vbNewLine & JSON_InnerIndentation & JSON_Converted
                            End If

                            JSON_BufferAppend JSON_Buffer, JSON_Converted, JSON_BufferPosition, JSON_BufferLength
                        Next JSON_Index2D

                        If JSON_PrettyPrint Then
                            JSON_BufferAppend JSON_Buffer, vbNewLine, JSON_BufferPosition, JSON_BufferLength
                        End If

                        JSON_BufferAppend JSON_Buffer, JSON_Indentation & "]", JSON_BufferPosition, JSON_BufferLength
                        JSON_IsFirstItem2D = True
                    Else

                        JSON_Converted = Stringify(JsonValue(JSON_Index), Whitespace, JSON_CurrentIndentation + 1)

                        If JSON_Converted = "" Then
                            If JSON_IsUndefined(JsonValue(JSON_Index)) Then
                                JSON_Converted = "null"
                            End If
                        End If

                        If JSON_PrettyPrint Then
                            JSON_Converted = vbNewLine & JSON_Indentation & JSON_Converted
                        End If

                        JSON_BufferAppend JSON_Buffer, JSON_Converted, JSON_BufferPosition, JSON_BufferLength
                    End If
                Next JSON_Index
            End If

            On Error GoTo 0

            If JSON_PrettyPrint Then
                JSON_BufferAppend JSON_Buffer, vbNewLine, JSON_BufferPosition, JSON_BufferLength

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    JSON_Indentation = VBA.String$(JSON_CurrentIndentation, Whitespace)
                Else
                    JSON_Indentation = VBA.Space$(JSON_CurrentIndentation * Whitespace)
                End If
            End If

            JSON_BufferAppend JSON_Buffer, JSON_Indentation & "]", JSON_BufferPosition, JSON_BufferLength

            Stringify = JSON_BufferToString(JSON_Buffer, JSON_BufferPosition)

        ' 对象类型
        Case VBA.vbObject

            If JSON_PrettyPrint Then
                If VBA.VarType(Whitespace) = VBA.vbString Then
                    JSON_Indentation = VBA.String$(JSON_CurrentIndentation + 1, Whitespace)
                Else
                    JSON_Indentation = VBA.Space$((JSON_CurrentIndentation + 1) * Whitespace)
                End If
            End If

            ' 字典
            If VBA.typeName(JsonValue) = "Dictionary" Then
                JSON_BufferAppend JSON_Buffer, "{", JSON_BufferPosition, JSON_BufferLength
                For Each JSON_Key In JsonValue.Keys
                    JSON_Converted = Stringify(JsonValue(JSON_Key), Whitespace, JSON_CurrentIndentation + 1)
                    If JSON_Converted = "" Then
                        JSON_SkipItem = JSON_IsUndefined(JsonValue(JSON_Key))
                    Else
                        JSON_SkipItem = False
                    End If

                    If Not JSON_SkipItem Then
                        If JSON_IsFirstItem Then
                            JSON_IsFirstItem = False
                        Else
                            JSON_BufferAppend JSON_Buffer, ",", JSON_BufferPosition, JSON_BufferLength
                        End If

                        If JSON_PrettyPrint Then
                            JSON_Converted = vbNewLine & JSON_Indentation & """" & JSON_Key & """: " & JSON_Converted
                        Else
                            JSON_Converted = """" & JSON_Key & """:" & JSON_Converted
                        End If

                        JSON_BufferAppend JSON_Buffer, JSON_Converted, JSON_BufferPosition, JSON_BufferLength
                    End If
                Next JSON_Key

                If JSON_PrettyPrint Then
                    JSON_BufferAppend JSON_Buffer, vbNewLine, JSON_BufferPosition, JSON_BufferLength

                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        JSON_Indentation = VBA.String$(JSON_CurrentIndentation, Whitespace)
                    Else
                        JSON_Indentation = VBA.Space$(JSON_CurrentIndentation * Whitespace)
                    End If
                End If

                JSON_BufferAppend JSON_Buffer, JSON_Indentation & "}", JSON_BufferPosition, JSON_BufferLength

            ' 集合
            ElseIf VBA.typeName(JsonValue) = "Collection" Then
                JSON_BufferAppend JSON_Buffer, "[", JSON_BufferPosition, JSON_BufferLength
                For Each JSON_Value In JsonValue
                    If JSON_IsFirstItem Then
                        JSON_IsFirstItem = False
                    Else
                        JSON_BufferAppend JSON_Buffer, ",", JSON_BufferPosition, JSON_BufferLength
                    End If

                    JSON_Converted = Stringify(JSON_Value, Whitespace, JSON_CurrentIndentation + 1)

                    If JSON_Converted = "" Then
                        If JSON_IsUndefined(JSON_Value) Then
                            JSON_Converted = "null"
                        End If
                    End If

                    If JSON_PrettyPrint Then
                        JSON_Converted = vbNewLine & JSON_Indentation & JSON_Converted
                    End If

                    JSON_BufferAppend JSON_Buffer, JSON_Converted, JSON_BufferPosition, JSON_BufferLength
                Next JSON_Value

                If JSON_PrettyPrint Then
                    JSON_BufferAppend JSON_Buffer, vbNewLine, JSON_BufferPosition, JSON_BufferLength

                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        JSON_Indentation = VBA.String$(JSON_CurrentIndentation, Whitespace)
                    Else
                        JSON_Indentation = VBA.Space$(JSON_CurrentIndentation * Whitespace)
                    End If
                End If

                JSON_BufferAppend JSON_Buffer, JSON_Indentation & "]", JSON_BufferPosition, JSON_BufferLength
            End If

            Stringify = JSON_BufferToString(JSON_Buffer, JSON_BufferPosition)

        '数字类型
        Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
            Stringify = VBA.Replace(JsonValue, ",", ".")

        '其他类型
        Case Else
            On Error Resume Next
            Stringify = JsonValue
            On Error GoTo 0
    End Select
End Function
