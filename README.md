<div align="center">

## Store Binary files in a database


</div>

### Description

Here are two functions that will allow you to

read and write a Large Binary Object

(BLOB) to and from a database. This could be used

to store and retrieve images, documents, etc

inside the database it self. This is great for

those of use that have a lot of Binary files

around that we want to keep in a central place

that can be backed up and protected with the same

security that a database offers.

This code will work with *ANY* database that ADO can connect to.
 
### More Info
 
The code needs a referance to the ADO library.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Henri](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/henri.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/henri-store-binary-files-in-a-database__1-24907/archive/master.zip)





### Source Code

```
'***************************************************************
' Abstract: Writes a BLOB datafield to a file. If the Data Field is
'  big I would recommend that you set bUseStream = False.
'
' Input: strFullPath: Full path to the destination file
'  objField: Field object that contains the BLOB data.
'  bUseStream: (Optional) True = Use Stream methode, False = Use GetChunk
'  lngChunkSize: (Optional) Specifies the Chunk size to fetch with each GetChunk
'
' Output: True on success, False on failure
'***************************************************************
Public Function BLOBToFile(ByVal strFullPath As String, ByRef objField As ADODB.Field, Optional ByVal bUseStream As Boolean = True, Optional ByVal lngChunkSize As Long = 8192) As Boolean
On Error Resume Next
Dim objStream As ADODB.Stream
Dim intFreeFile As Integer
Dim lngBytesLeft As Long
Dim lngReadBytes As Long
Dim byBuffer() As Byte
 If bUseStream Then
 Set objStream = New ADODB.Stream
 With objStream
 .Type = adTypeBinary
 .Open
 .Write objField.Value
 .SaveToFile strFullPath, adSaveCreateOverWrite
 End With
 DoEvents
 Else
 If Dir(strFullPath) <> "" Then
 Kill strFullPath
 End If
 lngBytesLeft = objField.ActualSize
 intFreeFile = FreeFile
 Open strFullPath For Binary As #intFreeFile
 Do Until lngBytesLeft <= 0
 lngReadBytes = lngBytesLeft
 If lngReadBytes > lngChunkSize Then
 lngReadBytes = lngChunkSize
 End If
 byBuffer = objField.GetChunk(lngReadBytes)
 Put #intFreeFile, , byBuffer
 lngBytesLeft = lngBytesLeft - lngReadBytes
 DoEvents
 Loop
 Close #intFreeFile
 End If
 If Err.Number <> 0 Or Err.LastDllError <> 0 Then
 BLOBToFile = False
 Else
 BLOBToFile = True
 End If
End Function
'***************************************************************
' Abstract: Writes a binary file to a BLOB datafield. If the file
'  is big I would recommend that you set bUseStream = False.
'
' Input: strFullPath: Full path to the source file
'  objField: Field object that will contain the BLOB data.
'  bUseStream: (Optional) True = Use Stream methode, False = Use GetChunk
'  lngChunkSize: (Optional) Specifies the Chunk size to fetch with each GetChunk
'
' Output: True on success, False on failure
'***************************************************************
Public Function FileToBLOB(ByVal strFullPath As String, ByRef objField As ADODB.Field, Optional ByVal bUseStream As Boolean = True, Optional ByVal lngChunkSize As Long = 8192) As Boolean
On Error Resume Next
Dim objStream As ADODB.Stream
Dim intFreeFile As Integer
Dim lngBytesLeft As Long
Dim lngReadBytes As Long
Dim byBuffer() As Byte
Dim varChunk As Variant
 If bUseStream Then
 Set objStream = New ADODB.Stream
 With objStream
 .Type = adTypeBinary
 .Open
 .LoadFromFile strFullPath
 objField.Value = .Read(adReadAll)
 End With
 Else
 With objField
 '<<--If the field does not support Long Binary data'-->>
 '<<--then we cannot load the data into the field.-->>
 If (.Attributes And adFldLong) <> 0 Then
 intFreeFile = FreeFile
 Open strFullPath For Binary Access Read As #intFreeFile
 lngBytesLeft = LOF(intFreeFile)
 Do Until lngBytesLeft <= 0
  If lngBytesLeft > lngChunkSize Then
  lngReadBytes = lngChunkSize
  Else
  lngReadBytes = lngBytesLeft
  End If
  ReDim byBuffer(lngReadBytes)
  Get #intFreeFile, , byBuffer()
  objField.AppendChunk byBuffer()
  lngBytesLeft = lngBytesLeft - lngReadBytes
  DoEvents
 Loop
 Close #intFreeFile
 Else
 Err.Raise -10000, "FileToBLOB", "The Database Field does not support Long Binary Data."
 End If
 End With
 End If
 If Err.Number <> 0 Or Err.LastDllError <> 0 Then
 FileToBLOB = False
 Else
 FileToBLOB = True
 End If
End Function
```

