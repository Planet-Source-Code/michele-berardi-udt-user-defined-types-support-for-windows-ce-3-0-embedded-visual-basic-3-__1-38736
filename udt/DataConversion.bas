Attribute VB_Name = "DataConversion"
Option Explicit

Public Function LongToBinaryString(ByVal LongIn As Long) As String
Dim i As Integer, aux As Long
    aux = LongIn: If LongIn < 0 Then aux = LongIn - &H80000000
    For i = 1 To 3
        LongToBinaryString = LongToBinaryString & ChrB(aux Mod 256)
        aux = aux \ 256
    Next i
    If LongIn < 0 Then
        LongToBinaryString = LongToBinaryString & ChrB(aux + &H80)
    Else
        LongToBinaryString = LongToBinaryString & ChrB(aux)
    End If
End Function

Public Function BinaryStringToLong(StringIn As String) As Long
Dim i As Integer, Neg As Boolean, aux As Integer
    If AscB(MidB(StringIn, 4)) And &H80 Then Neg = True
    For i = 4 To 1 Step -1
        aux = AscB(MidB(StringIn, i))
        If Neg Then aux = aux - &HFF
        BinaryStringToLong = BinaryStringToLong * 256 + aux
    Next i
    If Neg Then BinaryStringToLong = BinaryStringToLong - 1
End Function


Public Function IntegerToBinaryString(ByVal IntIn As Integer) As String
Dim aux As Integer
    aux = IntIn: If IntIn < 0 Then aux = IntIn - &H8000
    IntegerToBinaryString = IntegerToBinaryString & ChrB(aux Mod 256)
    aux = aux \ 256
    If IntIn < 0 Then
        IntegerToBinaryString = IntegerToBinaryString & ChrB(aux + &H80)
    Else
        IntegerToBinaryString = IntegerToBinaryString & ChrB(aux)
    End If
End Function

Public Function BinaryStringToInteger(StringIn As String) As Integer
    If AscB(MidB(StringIn, 2)) And &H80 Then
        BinaryStringToInteger = (AscB(MidB(StringIn, 2)) - &HFF) * 256 + (AscB(StringIn) - 256)
    Else
        BinaryStringToInteger = AscB(MidB(StringIn, 2)) * 256 + AscB(StringIn)
    End If
End Function


Public Function LongToBEBinaryString(ByVal LongIn As Long) As String
Dim i As Integer, aux As Long
    aux = LongIn: If LongIn < 0 Then aux = LongIn - &H80000000
    For i = 1 To 3
        LongToBEBinaryString = ChrB(aux Mod 256) & LongToBEBinaryString
        aux = aux \ 256
    Next i
    If LongIn < 0 Then
        LongToBEBinaryString = ChrB(aux + &H80) & LongToBEBinaryString
    Else
        LongToBEBinaryString = ChrB(aux) & LongToBEBinaryString
    End If
End Function

Public Function BEBinaryStringToLong(StringIn As String) As Long
Dim i As Integer, Neg As Boolean, aux As Integer
    If AscB(MidB(StringIn, 1)) And &H80 Then Neg = True
    For i = 1 To 4
        aux = AscB(MidB(StringIn, i))
        If Neg Then aux = aux - &HFF
        BEBinaryStringToLong = BEBinaryStringToLong * 256 + aux
    Next i
    If Neg Then BEBinaryStringToLong = BEBinaryStringToLong - 1
End Function


Public Function StringToFixLenStringW(StrIn As String, LenIn As Integer) As String
    StringToFixLenStringW = StrIn & String(LenIn - Len(StrIn), Chr(0))
End Function

