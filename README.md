<div align="center">

## \_A Smart String Comparison


</div>

### Description

This takes 2 strings and returns the percent alike that they are. (i.e. "test string number 1" is 86.48% similar to "teststring numb 2")

This function is very useful! You can use it in databases to match data that may have errors in it. Examples being people's names, company names, addresses, or anything else where you may encounter misspellings or inconsistencies in the data. Your feedback and/or votes are greatly appreciated! -- NEW - updated to use byte arrays instead of strings, 50-300% performance improvement!

An implementation of the , Ratcliff/Obershelp/Levenshtein method.
 
### More Info
 
mainstring and checkstring, the 2 strings to compare

This code recursively loops through the 2 strings, finding the largest common substring, then checking the remainder of the string.

how similar the 2 strings are (percent, as in .8)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Atul Brad Buono](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/atul-brad-buono.md)
**Level**          |Advanced
**User Rating**    |4.9 (168 globes from 34 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/atul-brad-buono-a-smart-string-comparison__1-9353/archive/master.zip)





### Source Code

```
Private b1() As Byte
Private b2() As Byte
Public Function Simil(String1 As String, String2 As String) As Double
 Dim l1 As Long
 Dim l2 As Long
 Dim l As Long
 Dim r As Double
 If UCase(String1) = UCase(String2) Then
  r = 1
 Else
  l1 = Len(String1)
  l2 = Len(String2)
  If l1 = 0 Or l2 = 0 Then
   r = 0
  Else
   ReDim b1(1 To l1): ReDim b2(1 To l2)
   For l = 1 To l1
    b1(l) = Asc(UCase(Mid(String1, l, 1)))
   Next
   For l = 1 To l2
    b2(l) = Asc(UCase(Mid(String2, l, 1)))
   Next
   r = SubSim(1, l1, 1, l2) / (l1 + l2) * 2
  End If
 End If
 Simil = r
 Erase b1
 Erase b2
End Function
Private Function SubSim(st1 As Long, end1 As Long, st2 As Long, end2 As Long) As Long
 Dim c1 As Long
 Dim c2 As Long
 Dim ns1 As Long
 Dim ns2 As Long
 Dim i As Long
 Dim max As Long
 If st1 > end1 Or st2 > end2 Or st1 <= 0 Or st2 <= 0 Then Exit Function
 For c1 = st1 To end1
  For c2 = st2 To end2
   i = 0
   Do Until b1(c1 + i) <> b2(c2 + i)
    i = i + 1
    If i > max Then
     ns1 = c1
     ns2 = c2
     max = i
    End If
    If c1 + i > end1 Or c2 + i > end2 Then Exit Do
   Loop
  Next
 Next
 max = max + SubSim(ns1 + max, end1, ns2 + max, end2)
 max = max + SubSim(st1, ns1 - 1, st2, ns2 - 1)
 SubSim = max
End Function
```

