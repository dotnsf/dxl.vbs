' dxl_import.vbs
' Execute: c:\Windows\SysWOW64\CScript //nologo dxl_import.vbs c:\Users\username\fullpath\xxxxx.xml
Option Explicit

Dim objWsh
Dim objArgs
Dim objDbPath
Dim objNotesSession
Dim objNotesDbDirectory
Dim objNotesDb
Dim stream
Dim dxlImporter
Dim dxl

Dim inputFilePath
Dim outputFilePath
Dim fso
Dim file

Dim tmpArr, tmpStr, tmpInt
Dim tmpDArr
Dim i, c, encoding

Set objArgs = Wscript.Arguments
If objArgs.Count = 0 Then
  Wscript.Echo "Please specify xml file path as command line parameter."
Else
  c = 0
  encoding = "SHIFT_JIS"
  
  For i = 0 To objArgs.Count - 2
    If StartsWith( LCase( objArgs(i) ), "-encoding=" ) = 1 Then
      encoding = Mid( objArgs(i), 10 )
    End If
    c = c + 1
  Next

  'local xml filepath
  inputFilePath = objArgs(0)
  
  'normalize filepath
  tmpArr = Split( inputFilePath, "/" )
  inputFilePath = Join( tmpArr, "\" )
  
  'output DB filepath
  tmpDArr = Split( inputFilePath, "." )
  Redim Preserve tmpDArr(UBound(tmpDArr)-1)
  outputFilePath = Join( tmpDArr, "." ) & ".nsf"
  'Wscript.Echo outputFilePath

  Set objNotesSession = Wscript.CreateObject( "Lotus.NotesSession" )
  Call objNotesSession.Initialize

  Set stream = objNotesSession.CreateStream
  If Not stream.Open( inputFilePath, encoding ) Then
    Wscript.Echo "Please specify valid local XML file path as command line parameter."
  Else
    If stream.Bytes = 0 Then
      Wscript.Echo "File did not exist or was empty.", , inputFilePath
    Else
      Set fso = CreateObject( "Scripting.FileSystemObject" )

      If fso.FileExists( outputFilePath ) = True Then
        fso.DeleteFile( outputFilePath )
      End If
    
      Set objNotesDbDirectory = objNotesSession.GetDbDirectory( "" )
      Set objNotesDb = objNotesDbDirectory.CreateDatabase( outputFilePath, True )
      'Call objNotesDb.Create( "", outputFilePath, True )
      'Set dxlImporter = objNotesSession.CreateDXLImporter( stream, objNotesDb )
      Set dxlImporter = objNotesSession.CreateDXLImporter
      dxlImporter.ReplaceDBProperties = True
      dxlImporter.ReplicaRequiredForReplaceOrUpdate = True 'False
      dxlImporter.ACLImportOption = 5
      dxlImporter.DesignImportOption = 2
      'Call dxlImporter.Process
      Call dxlImporter.Import( stream, objNotesDb )

      Wscript.Echo outputFilePath
    End If
  End If

  Set objNotesSession = Nothing
End If


Public Function StartsWith(target_str, search_str)
  StartsWith = 0
  If Len(search_str) > Len(target_str) Then
    Exit Function
  End If
  
  If Left(target_str, Len(search_str)) = search_str Then
    StartsWith = 1
  End If
End Function

Public Function EndsWith(target_str, search_str)
  EndsWith = 0
  If Len(search_str) > Len(target_str) Then
    Exit Function
  End If
  
  If Right(target_str, Len(search_str)) = search_str Then
    EndsWith = 1
  End If
End Function


