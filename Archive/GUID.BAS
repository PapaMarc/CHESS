Attribute VB_Name = "GUID"
'Declarations and Functions to support
'Microsoft Exchange conversation topic and threading

'Copyright 1994 Microsoft Corporation

Option Explicit

Type GUID
    Guid1 As Long
    Guid2 As Long
    Guid3 As Long
    Guid4 As Long
End Type

Type FILETIME
    lLowDate As Long
    lHighDate As Long
End Type

Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Declare Function CoCreateGuid Lib "COMPOBJ.DLL" (pGuid As GUID) As Long
Declare Function GetSystemTime Lib "mapi.DLL" (pSysTime As SYSTEMTIME) As Long
Declare Function SystemTimeToFileTime Lib "mapi.DLL" (pSysTime As SYSTEMTIME, pFileTime As FILETIME) As Long
Global Const S_OK = 0

Function sNewConversationIndex() As String
    'Exchange conversation is a unique 8-byte value
    
    Dim lResult As Long
    Dim lGuid As GUID
    Dim lTime As FILETIME
    Dim sTime As SYSTEMTIME
    Dim sTimeData
    Dim a As Integer
    
    Call GetSystemTime(sTime)

    lResult = SystemTimeToFileTime(sTime, lTime)
    lResult = CoCreateGuid(lGuid)
    If lResult = S_OK Then
        sTimeData = LongToHex(lTime.lHighDate) & LongToHex(lTime.lLowDate)
        sTimeData = Mid(sTimeData, 3, 10) ' cut off upper byte and lower 2 bytes
        sNewConversationIndex = Hex$(0) & Hex$(1) & sTimeData & LongToHex(lGuid.Guid1) & LongToHex(lGuid.Guid2) & LongToHex(lGuid.Guid3) & LongToHex(lGuid.Guid4)
    End If
End Function

Function sAddConversationIndex(CurrentIndex As String) As String
    'Exchange conversation is a unique 8-byte value
    
    Dim lResult As Long
    Dim lTime As FILETIME
    Dim sTimeStr As String
    Dim sDTimeStr As String
    Dim sThread As String
    
    Dim firstTime As Currency
    Dim totalTime As Currency
    Dim TimeAsCurrency As Currency
    Dim TempCurr As Currency
    Dim lToCurr As Currency
    Dim Cur32Bits As Currency
    Dim Cur23Bits As Currency
    Dim Cur18Bits As Currency
    Dim indxCount As Integer
    Dim sTime As SYSTEMTIME
    Dim dlTime As Long
    
    indxCount = 0
    
    lToCurr = 0.0001
    
    Cur32Bits = CCur(2 ^ 30) * 4
    
    Cur23Bits = CCur(2 ^ 23)
    
    Cur18Bits = CCur(2 ^ 18)
    
    sTimeStr = Mid(CurrentIndex, 3, 10) ' get the initial time
    
    firstTime = CCur(HexToLong(Left$(sTimeStr, 6)))
    
    firstTime = (firstTime * lToCurr) * (256 ^ 2)
    
    firstTime = firstTime + (CCur(HexToLong(Right$(sTimeStr, 4))) * lToCurr)

    firstTime = firstTime * (256 ^ 2)
    
    '
    ' get the delta time threads
    '
    sTimeStr = Right$(CurrentIndex, Len(CurrentIndex) - 44)
    totalTime = 0
    
    If Len(sTimeStr) >= 10 Then
        '
        ' loop throught all of the 5 byte index threads
        '
        
        Do While Len(sTimeStr) > 0
            sDTimeStr = Left(sTimeStr, 8) ' get the delta time
            '
            ' The uppermost bit is used to determine the units
            ' of time, not the sign of the number
            '

                TempCurr = CCur(HexToLong("0" & Left(sTimeStr, 2))) * lToCurr
                
                TempCurr = (TempCurr * 256) + (CCur("0" & HexToLong(Mid(sTimeStr, 3, 2))) * lToCurr)

                TempCurr = (TempCurr * 256) + (CCur("0" & HexToLong(Mid(sTimeStr, 5, 2))) * lToCurr)
                
                TempCurr = (TempCurr * 256) + (CCur("0" & HexToLong(Mid(sTimeStr, 7, 2))) * lToCurr)
                
            dlTime = HexToLong(sDTimeStr)
            
            If dlTime > 0 Then
                
                totalTime = totalTime + (TempCurr * Cur18Bits)
               
            Else
                If dlTime < -&H10000 Then  ' FIX for backwards time bug (Remove for final)
                    totalTime = totalTime + (TempCurr * Cur23Bits)
                End If
            End If
            
            sTimeStr = Mid$(sTimeStr, 11)
            
        Loop
        indxCount = HexToLong(Right(CurrentIndex, 1)) ' get the last index count
        indxCount = indxCount + 1
    End If
    '
    ' now calculate the delta time
    '
    lResult = GetSystemTime(sTime)

    lResult = SystemTimeToFileTime(sTime, lTime)
    
    'lResult = CoFileTimeNow(lTime)

    TimeAsCurrency = ((CCur(lTime.lHighDate And &HFFFFFF) * lToCurr) * Cur32Bits)
    
    If lTime.lLowDate < 0 Then
        TempCurr = (CCur(lTime.lLowDate And &H7FFF0000) * lToCurr) + CCur(&H7FFFFFFF * lToCurr) ' 2s complement
    Else
        TempCurr = (CCur(lTime.lLowDate And &HFFFF0000) * lToCurr)
    End If
    
    TimeAsCurrency = TimeAsCurrency + TempCurr

    TimeAsCurrency = TimeAsCurrency - firstTime
    
     
    totalTime = TimeAsCurrency - totalTime
    
    '
    ' check for final answer 31 bit overflow
    '
    
    If totalTime < 0 Then ' don't allow delta times < 0
        totalTime = Cur18Bits / 10000
    End If

    On Error GoTo OVERFlow
    
    If CCur(totalTime / Cur18Bits) > CCur(&H7FFFFFFF * lToCurr) Then
        totalTime = totalTime / Cur23Bits
        dlTime = CLng(totalTime * 10000)
    Else
        totalTime = totalTime / Cur18Bits
        dlTime = CLng(totalTime * 10000)
    End If
    
    sThread = LongToHex(dlTime) & Hex(Int(Rnd * 16)) & Hex(indxCount Mod 16)
    
    sAddConversationIndex = CurrentIndex & sThread
    
    Exit Function

OVERFlow:
    dlTime = &HFFFFFFFF ' max possible delta time
    Resume Next
     
End Function


Public Function LongToHex(lLong As Long) As String
    ' converts long to 8-digit hex string, left-padded as needed
    Dim Temp As String
    
    Temp = Hex$(lLong)
    LongToHex = Left$("00000000", 8 - Len(Temp)) & Temp
End Function

Public Function HexToLong(sHex As String) As Long
    ' converts 8-digit hex string to long
    
    HexToLong = CLng("&H" & sHex)
    
End Function



