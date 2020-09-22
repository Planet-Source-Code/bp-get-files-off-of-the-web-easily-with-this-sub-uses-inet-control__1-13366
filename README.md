<div align="center">

## Get files off of the web easily with this sub \(uses Inet Control\)


</div>

### Description

This sub will allow you to downloaded files off of the web by making one call. The syntax will look sometime to this effect:

Call GetInternetFile(Inet1, "http://www.fakeURL.com/~yup/picture.jpg", "c:\temp")
 
### More Info
 
For this sub to work you need to have an Inet control placed onto your form. After that it's just a matter of calling it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[BP](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bp.md)
**Level**          |Intermediate
**User Rating**    |4.8 (43 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bp-get-files-off-of-the-web-easily-with-this-sub-uses-inet-control__1-13366/archive/master.zip)





### Source Code

```
Public Sub GetInternetFile(Inet1 As Inet, myURL As String, DestDIR As String)
' Written by: Blake Pell
Dim myData() As Byte
If Inet1.StillExecuting = True Then Exit Sub
myData() = Inet1.OpenURL(myURL, icByteArray)
For X = Len(myURL) To 1 Step -1
  If Left$(Right$(myURL, X), 1) = "/" Then RealFile$ = Right$(myURL, X - 1)
Next X
myFile$ = DestDIR + "\" + RealFile$
Open myFile$ For Binary Access Write As #1
  Put #1, , myData()
Close #1
End Sub
```

