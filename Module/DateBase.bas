Attribute VB_Name = "DateBase"
'@Lang VBA

Option Compare Database
Option Explicit


Private Type UTC_SYSTEMTIME
    UTC_Year As Integer
    UTC_Month As Integer
    UTC_DayOfWeek As Integer
    UTC_Day As Integer
    UTC_Hour As Integer
    UTC_Minute As Integer
    UTC_Second As Integer
    UTC_Milliseconds As Integer
End Type

Private Type UTC_TIME_ZONE_INFORMATION
    UTC_Bias As Long
    UTC_StandardName(0 To 31) As Integer
    UTC_StandardDate As UTC_SYSTEMTIME
    UTC_StandardBias As Long
    UTC_DaylightName(0 To 31) As Integer
    UTC_DaylightDate As UTC_SYSTEMTIME
    UTC_DaylightBias As Long
End Type


' 私有方法

'UTC_GetTimeZoneInformation
Private Declare PtrSafe Function UTC_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" (UTC_lpTimeZoneInformation As UTC_TIME_ZONE_INFORMATION) As Long

'UTC_SystemTimeToTzSpecificLocalTime
Private Declare PtrSafe Function UTC_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (UTC_lpTimeZoneInformation As UTC_TIME_ZONE_INFORMATION, UTC_lpUniversalTime As UTC_SYSTEMTIME, UTC_lpLocalTime As UTC_SYSTEMTIME) As Long

'UTC_TzSpecificLocalTimeToSystemTime
Private Declare PtrSafe Function UTC_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (UTC_lpTimeZoneInformation As UTC_TIME_ZONE_INFORMATION, UTC_lpLocalTime As UTC_SYSTEMTIME, UTC_lpUniversalTime As UTC_SYSTEMTIME) As Long


Private Function UTC_DateToSystemTime(UTC_Value As Date) As UTC_SYSTEMTIME
    UTC_DateToSystemTime.UTC_Year = VBA.Year(UTC_Value)
    UTC_DateToSystemTime.UTC_Month = VBA.Month(UTC_Value)
    UTC_DateToSystemTime.UTC_Day = VBA.Day(UTC_Value)
    UTC_DateToSystemTime.UTC_Hour = VBA.Hour(UTC_Value)
    UTC_DateToSystemTime.UTC_Minute = VBA.Minute(UTC_Value)
    UTC_DateToSystemTime.UTC_Second = VBA.Second(UTC_Value)
    UTC_DateToSystemTime.UTC_Milliseconds = 0
End Function

Private Function UTC_SystemTimeToDate(UTC_Value As UTC_SYSTEMTIME) As Date
    UTC_SystemTimeToDate = DateSerial(UTC_Value.UTC_Year, UTC_Value.UTC_Month, UTC_Value.UTC_Day) + _
        TimeSerial(UTC_Value.UTC_Hour, UTC_Value.UTC_Minute, UTC_Value.UTC_Second)
End Function



' 公共方法
Public Function ParseUtc(UTC_UtcDate As Date) As Date
    On Error GoTo UTC_ErrorHandling

    Dim UTC_TimeZoneInfo As UTC_TIME_ZONE_INFORMATION
    Dim UTC_LocalDate As UTC_SYSTEMTIME

    UTC_GetTimeZoneInformation UTC_TimeZoneInfo
    UTC_SystemTimeToTzSpecificLocalTime UTC_TimeZoneInfo, UTC_DateToSystemTime(UTC_UtcDate), UTC_LocalDate

    ParseUtc = UTC_SystemTimeToDate(UTC_LocalDate)

    Exit Function

UTC_ErrorHandling:
    err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & err.Number & " - " & err.Description
End Function


Public Function ConvertToUtc(UTC_LocalDate As Date) As Date
    On Error GoTo UTC_ErrorHandling

    Dim UTC_TimeZoneInfo As UTC_TIME_ZONE_INFORMATION
    Dim UTC_UtcDate As UTC_SYSTEMTIME

    UTC_GetTimeZoneInformation UTC_TimeZoneInfo
    UTC_TzSpecificLocalTimeToSystemTime UTC_TimeZoneInfo, UTC_DateToSystemTime(UTC_LocalDate), UTC_UtcDate

    ConvertToUtc = UTC_SystemTimeToDate(UTC_UtcDate)

    Exit Function

UTC_ErrorHandling:
    err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & err.Number & " - " & err.Description
End Function


Public Function ParseIso(UTC_IsoString As String) As Date
    On Error GoTo UTC_ErrorHandling

    Dim UTC_Parts() As String
    Dim UTC_DateParts() As String
    Dim UTC_TimeParts() As String
    Dim UTC_OffsetIndex As Long
    Dim UTC_HasOffset As Boolean
    Dim UTC_NegativeOffset As Boolean
    Dim UTC_OffsetParts() As String
    Dim UTC_Offset As Date

    UTC_Parts = VBA.Split(UTC_IsoString, "T")
    UTC_DateParts = VBA.Split(UTC_Parts(0), "-")
    ParseIso = VBA.DateSerial(VBA.CInt(UTC_DateParts(0)), VBA.CInt(UTC_DateParts(1)), VBA.CInt(UTC_DateParts(2)))

    If UBound(UTC_Parts) > 0 Then
        If VBA.InStr(UTC_Parts(1), "Z") Then
            UTC_TimeParts = VBA.Split(VBA.Replace(UTC_Parts(1), "Z", ""), ":")
        Else
            UTC_OffsetIndex = VBA.InStr(1, UTC_Parts(1), "+")
            If UTC_OffsetIndex = 0 Then
                UTC_NegativeOffset = True
                UTC_OffsetIndex = VBA.InStr(1, UTC_Parts(1), "-")
            End If

            If UTC_OffsetIndex > 0 Then
                UTC_HasOffset = True
                UTC_TimeParts = VBA.Split(VBA.Left$(UTC_Parts(1), UTC_OffsetIndex - 1), ":")
                UTC_OffsetParts = VBA.Split(VBA.Right$(UTC_Parts(1), Len(UTC_Parts(1)) - UTC_OffsetIndex), ":")

                Select Case UBound(UTC_OffsetParts)
                Case 0
                    UTC_Offset = TimeSerial(VBA.CInt(UTC_OffsetParts(0)), 0, 0)
                Case 1
                    UTC_Offset = TimeSerial(VBA.CInt(UTC_OffsetParts(0)), VBA.CInt(UTC_OffsetParts(1)), 0)
                Case 2
                    UTC_Offset = TimeSerial(VBA.CInt(UTC_OffsetParts(0)), VBA.CInt(UTC_OffsetParts(1)), Int(VBA.val(UTC_OffsetParts(2))))
                End Select

                If UTC_NegativeOffset Then: UTC_Offset = -UTC_Offset
            Else
                UTC_TimeParts = VBA.Split(UTC_Parts(1), ":")
            End If
        End If

        Select Case UBound(UTC_TimeParts)
        Case 0
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(UTC_TimeParts(0)), 0, 0)
        Case 1
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(UTC_TimeParts(0)), VBA.CInt(UTC_TimeParts(1)), 0)
        Case 2
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(UTC_TimeParts(0)), VBA.CInt(UTC_TimeParts(1)), Int(VBA.val(UTC_TimeParts(2))))
        End Select

        ParseIso = ParseUtc(ParseIso)

        If UTC_HasOffset Then
            ParseIso = ParseIso - UTC_Offset
        End If
    End If

    Exit Function

UTC_ErrorHandling:
    err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & UTC_IsoString & ": " & err.Number & " - " & err.Description
End Function


Public Function ConvertToIso(UTC_LocalDate As Date) As String
    On Error GoTo UTC_ErrorHandling

    ConvertToIso = VBA.Format$(ConvertToUtc(UTC_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")

    Exit Function

UTC_ErrorHandling:
    err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & err.Number & " - " & err.Description
End Function
