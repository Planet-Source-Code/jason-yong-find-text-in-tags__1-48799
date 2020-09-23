<div align="center">

## Find text in tags


</div>

### Description

This is just a simple code to find the text in a tag. like, [tag] I'm a tag![/tag]
 
### More Info
 
Text = is the string to search

StartTag = is the first part of the tag to look for. like [tag]

EndTag = Is the end tag like [/tag]

'it will return different values depending on the ReturnCode

if ReturnCode = 1 will return the text

2 will return where the first tag starts

3 will return where the text in the tag starts

4 will return where the tag ends


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jason Yong](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jason-yong.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jason-yong-find-text-in-tags__1-48799/archive/master.zip)





### Source Code

```
'=================InTag Function====================
'=================================
'(C) Jason Yong 2003
'=================================
'This function will find what text is in a tag, like [b]this[/b]
'it will return different values depending on the ReturnCode
'ReturnCode = 1 will return the text
' 2 will return where the first tag starts
' 3 will return where the text in the tag starts
' 4 will return where the tag ends.
'===================================================
Function InTag(Text As String, StartTag As String, EndTag As String, ReturnCode As Integer) As String
Dim t As Long
Dim r As Long
 For t = 1 To Len(Text)
 If Mid(Text, t, Len(StartTag)) = StartTag Then
  For r = t To Len(Text)
  If Mid(Text, r, Len(EndTag)) = EndTag Then
   If ReturnCode = 1 Or ReturnCode = 0 Then InTag = Mid(Text, t + Len(StartTag), r - (t + Len(StartTag)))
   If ReturnCode = 2 Then InTag = Str(t)
   If ReturnCode = 3 Then InTag = Str(t) + Len(StartTag)
   If ReturnCode = 4 Then InTag = Str(r - (t + Len(StartTag) + Len(EndTag)))
  End If
  Next
 End If
 Next
End Function
```

