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

Set objArgs = Wscript.Arguments
If objArgs.Count = 0 Then
  Wscript.Echo "Please specify xml file path as command line parameter."
Else
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
  If Not stream.Open( inputFilePath, "Shift_JIS" ) Then
    Wscript.Echo "Please specify valid local XML file path as command line parameter."
  Else
    If stream.Bytes = 0 Then
      Wscript.Echo "File did not exist or was empty.", , inputFilePath
    Else
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
