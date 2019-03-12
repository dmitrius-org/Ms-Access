Attribute VB_Name = "cmn_String"
Option Compare Database
Option Explicit

Public Function InStrCunt(ByVal AText As String, ByVal ASubText As String) As Integer
'InStrCunt - подсчет количества вхождений построки в строку

Dim K As Integer
Dim Kcount As Integer
  K = 1
  Kcount = 0

  Do
    K = InStr(K, AText, ASubText)
    If K = 0 Then
      Exit Do
    Else
      K = K + 1
      Kcount = Kcount + 1
    End If
  Loop
  
InStrCunt = Kcount
End Function

Public Function TrimL(ByVal AText As String) As String
'TrimL -удаление пробелов слева

Dim i As Integer
Dim arrTest() As String
Dim tmpText As String

  TrimL = ""
  arrTest = Split(AText, vbNewLine)

  For i = 0 To UBound(arrTest)
    tmpText = Trim(arrTest(i))
    
    If Len(tmpText) > 0 Then
      TrimL = TrimL & tmpText & vbNewLine
    End If
  Next
End Function

Public Function ConvNullToStr(Zm As Variant) As String
  If IsNull(Zm) Then ConvNullToStr = "" Else ConvNullToStr = Zm
End Function

Public Function ConvNullToTimeHM(Zm As Variant) As String
  If IsNull(Zm) Then ConvNullToTimeHM = "" Else ConvNullToTimeHM = Format(Zm, "hh:mm")
End Function

Public Function ConvNullToDate(Zm As Variant) As String
  If IsNull(Zm) Then ConvNullToDate = "" Else ConvNullToDate = Format(Zm, "dd.MM.yy")
End Function

Public Function ConvNullToCurrent(Zm As Variant) As String
  If IsNull(Zm) Then ConvNullToCurrent = "" Else ConvNullToCurrent = Format(Zm, "0.00")
End Function
