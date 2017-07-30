Attribute VB_Name = "模块1"
Public Enum EDutyType
    NONE = 0
    NORMAL = 2 ^ 0
    LATE_10_MIN = 2 ^ 1
    LATE_30_MIN = 2 ^ 2
    LATE_60_MIN = 2 ^ 3
    EARLY = 2 ^ 4
    HOLIDAY = 2 ^ 5
    ABSENCE = 2 ^ 6
    ABSENCE_SPE = 2 ^ 7
    ABSENCE_ILL = 2 ^ 8
End Enum

Public Function DutyTypeToString(dutyType As Integer) As String
    Dim str As String
    'Select Case dutyType
        If (EDutyType.NONE = dutyType) Then str = "未知"
        If EDutyType.NORMAL And dutyType Then
            str = "正常"
        Else
            Select Case True
                Case EDutyType.LATE_10_MIN And dutyType: str = "迟到10分钟"
                Case EDutyType.LATE_30_MIN And dutyType: str = "迟到10分钟"
                Case EDutyType.LATE_60_MIN And dutyType: str = "迟到1小时"
            End Select
        End If
        If EDutyType.HOLIDAY And dutyType Then str = str + "休息"
        If EDutyType.EARLY And dutyType Then str = str + "早退"
        If EDutyType.ABSENCE And dutyType Then str = str + "请假"
        If EDutyType.ABSENCE_SPE And dutyType Then str = str + "特休"
        If EDutyType.ABSENCE_ILL And dutyType Then str = str + "病假"
    'End Select
    DutyTypeToString = str
End Function

Public Function DateToDayMin(t As Date) As Integer
    'Dim min As Integer
    DateToDayMin = Hour(t) * 60 + Minute(t)
End Function

Public Function GetDefaultCheckInTime() As Date
    'Dim d As Date
    GetDefaultCheckInTime = #5/23/1999 9:30:00 AM#
End Function
Public Function GetDefaultCheckOutTime() As Date
    'Dim d As Date
    GetDefaultCheckOutTime = #5/23/1999 6:30:00 PM#
End Function

Public Function GetDutyState(checkInTime As Date, checkOutTime As Date) As Integer
    
    Dim reType As Integer
    reType = 1
    Dim dfcit As Date
    dfcit = GetDefaultCheckInTime()
    Dim dfcot As Date
    dfcot = GetDefaultCheckOutTime()
    dfm = DateToDayMin(checkInTime) - DateToDayMin(dfcit)
    Debug.Print ("checkInTime:" & checkInTime)
    Debug.Print ("checkOutTime:" & checkOutTime)
    Select Case True
        Case dfm >= 60: reType = EDutyType.LATE_60_MIN
        Case dfm >= 30: reType = EDutyType.LATE_30_MIN
        Case dfm >= 10:
            Debug.Print ("It is late then 10 minutes")
            reType = EDutyType.LATE_10_MIN
    End Select
    Debug.Print ("dfm:" & dfm)
    dfm = DateToDayMin(checkOutTime) - DateToDayMin(dfcot)
    If (dfm < 0) Then
        reType = reType + EDutyType.EARLY
    End If
    GetDutyState = reType
End Function

Public Sub CheckOneRowDuty(rowId As Integer, dayCount As Integer)
    Set ws = ActiveWorkbook.ActiveSheet
    For dayIt = 0 To dayCount
        Set curCell = ws.Cells(rowId, dayIt * 3 + 2)
        Dim checkInTime As Date
        checkInTime = CDate(curCell.Value)
        Set curCell = ws.Cells(rowId, dayIt * 3 + 3)
        Dim checkOutTime As Date
        checkOutTime = CDate(curCell.Value)
        Set curCell = ws.Cells(rowId, dayIt * 3 + 4)
        Dim dutyState As Integer
        dutyState = GetDutyState(checkInTime, checkOutTime)
        curCell.Value = DutyTypeToString(dutyState)
    Next dayIt
    CheckOneRowDuty = 0
End Sub

Sub test()
Attribute test.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏1 宏
'
    ' .Add.Name = "A New Sheet"
    'With ActiveWorkbook.Worksheets(1) 'Worksheets("A New Sheet")
    ' .Range("A5:A10").Formula = "=RAND()"
    'End With
    Dim rowId As Integer
    rowId = 2
    Dim dayCount As Integer
    dayCount = 1
    Set ws = ActiveWorkbook.ActiveSheet
    For dayIt = 0 To dayCount
        Set curCell = ws.Cells(rowId, dayIt * 3 + 2)
        Dim checkInTime As Date
        checkInTime = CDate(curCell.Value)
        Set curCell = ws.Cells(rowId, dayIt * 3 + 3)
        Dim checkOutTime As Date
        checkOutTime = CDate(curCell.Value)
        Set curCell = ws.Cells(rowId, dayIt * 3 + 4)
        Dim dutyState As Integer
        dutyState = GetDutyState(checkInTime, checkOutTime)
        curCell.Value = DutyTypeToString(dutyState)
    Next dayIt
    'CheckOneRowDuty = 0
   '
    
End Sub
