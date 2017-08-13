Const DestCharSet = "utf-8"
'Const DestCharSet = "ascii"
Dim FS
Set fs = CreateObject("Scripting.FileSystemObject")

ConvertFolder "f:\", "f:\1"

Function ConvertFolder(byval InputPath, OutputPath) 
  Dim InputFolder, File
  Set InputFolder = fs.GetFolder(InputPath)


  For Each File In InputFolder.Files
    If LCase(Right(File.Name,4)) = ".htm" Then
      Wscript.Echo File.Path
      'wscript.echo OutputPath & "\" & replace(file.path,":","")
      ConvertFile File.Path, OutputPath & "\" & file.Name, DestCharSet
    End If
  Next

  Dim FilesFolder
  For Each FilesFolder In InputFolder.SubFolders
    ConvertFolder FilesFolder.Path, OutputPath
  Next
End Function

Sub ConvertFile(SourceFileName, DestFileName, DestCharSet)
  'read the source file contents
  Dim FileContents
  Set FileContents = ReadOneFile(SourceFileName)

  'Convert to the destination charset
  Set FileContents = FileContents.CharSetConvert(DestCharSet)

  'Save to a destination file
  FileContents.SaveAs DestFileName 
End Sub 

Function ReadOneFile(FileName)
  Dim ByteArray
  Set ByteArray = CreateObject("ScriptUtils.ByteArray")

  'Read first two bytes from the file
  ByteArray.ReadFrom FileName,,2

  Select Case ByteArray.HexString
    'unicode big endian
    Case "FEFF": 
      ByteArray.CharSet = "unicodebig"
      'Read the file from 3rd byte to end.
      ByteArray.ReadFrom FileName,3
    'unicode little endian      
    Case "FFFE": 
      ByteArray.CharSet = "unicodelittle"
      'Read the file from 3rd byte to end.
      ByteArray.ReadFrom FileName,3
    Case Else: 
      'Read first three bytes from the file
      ByteArray.ReadFrom FileName,,3
      If ByteArray.HexString = "EFBBBF" Then 'unicode utf-8
        'read a file contents behind the BOM header 
        ByteArray.ReadFrom FileName,4
        ByteArray.CharSet = "utf-8"
      Else
        'read whole contents of the file in other cases
        ByteArray.ReadFrom FileName
        On Error Resume Next
        'try to detect charset from the data source'
        ByteArray.CharSet = DetectCharSet(ByteArray.String)
        'Set some default charset (default is OEM)
        'if err<>0 then ByteArray.CharSet = "windows-1250"
      End If 
  End Select
  Set ReadOneFile = ByteArray
End Function

'The Function detects charset from the source string data.
Function DetectCharSet(Data)
  On Error Resume Next
  Dim charset
  'the charset tag usually look like
  '<meta http-equiv="Content-Type" content="text/html; charset=windows-1250">
  charset = Split(Data, "charset=", 2, vbTextCompare)(1)
  If Len(charset)>0 Then
    charset = Split(charset, """", 2, vbTextCompare)(0)
  End If
  DetectCharSet = charset 
End Function