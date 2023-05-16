' dxl_export.vbs
' Execute: c:\Windows\SysWOW64\CScript //nologo dxl_export.vbs [-encoding=SHIFT_JIS] path/xxxxx.nsf
Option Explicit

Dim objWsh
Dim objArgs
Dim objDbPath
Dim objNotesSession
Dim objNotesDb
Dim nc
Dim dxlExporter
Dim dxl

Dim outputFileFolder
Dim outputFilePath
Dim fso
Dim file

Dim tmpArr, tmpStr, tmpInt
Dim tmpDArr
Dim i, c, encoding

Set objArgs = Wscript.Arguments
If objArgs.Count = 0 Then
  Wscript.Echo "Please specify local database path as command line parameter."
Else
  c = 0
  encoding = "SHIFT_JIS"
  
  For i = 0 To objArgs.Count - 2
    If StartsWith( LCase( objArgs(i) ), "-encoding=" ) = 1 Then
      encoding = Mid( objArgs(i), 10 )
    End If
    c = c + 1
  Next
  
  'local database filepath
  objDbPath = objArgs(c)
  
  'normalize DB filepath
  tmpArr = Split( objDbPath, "\" )
  objDbPath = Join( tmpArr, "/" )
  
  'output xml filepath
  tmpArr = Split( objDbPath, "/" )
  outputFileFolder = Join( tmpArr, "_" )
  outputFilePath = tmpArr(UBound(tmpArr))
  tmpDArr = Split( outputFilePath, "." )
  Redim Preserve tmpDArr(UBound(tmpDArr)-1)
  outputFilePath = outputFileFolder & "\" & Join( tmpDArr, "." ) & ".xml"
  'Wscript.Echo outputFilePath

  Set fso = CreateObject( "Scripting.FileSystemObject" )

  If fso.FolderExists( outputFileFolder ) = False Then
    fso.CreateFolder( outputFileFolder )
  End If

  Set objNotesSession = Wscript.CreateObject( "Lotus.NotesSession" )
  Call objNotesSession.Initialize

  'Wscript.Echo objDbPath
  Set objNotesDb = objNotesSession.GetDatabase( "", objDbPath )
  'Wscript.Echo objNotesDb.Title
  
  Set nc = objNotesDb.CreateNoteCollection( False )
  nc.SelectACL = True                  'ACL
  nc.SelectActions = True              'Action
  nc.SelectAgents = True               'Agents
  nc.SelectDatabaseScript = True       'DatabaseScript
  nc.SelectDataConnections = True 'False     'DataConnections
  nc.SelectDocuments = False           'Documents
  nc.SelectFolders = True              'Folders
  nc.SelectForms = True                'Forms
  nc.SelectFrameSets = True            'Framesets
  nc.SelectHelpAbout = True 'False           'HelpAbout
  nc.SelectHelpIndex = True 'False           'HelpIndex
  nc.SelectHelpUsing = True 'False           'HelpUsing
  nc.SelectIcon = True 'False                'HelpIcon
  nc.SelectImageResources = True 'False      'ImageResources
  nc.SelectJavaResources = True        'JavaResources
  nc.SelectMiscCodeElements = True 'False    'MiscCodeElements
  nc.SelectMiscFormatElements = True 'False  'MiscFormatElements
  nc.SelectMiscIndexElements = True 'False   'MiscIndexElements
  nc.SelectNavigators = True           'Navigators
  nc.SelectOutlines = True             'Outlines
  nc.SelectPages = True                'Pages
  nc.SelectProfiles = True 'False            'Profiles
  nc.SelectReplicationFormulas = True  'ReplicationFormulas
  nc.SelectScriptLibraries = True      'ScriptLibraries
  nc.SelectSharedFields = True         'SharedFields
  nc.SelectStyleSheetResources = True 'False 'StyleSheetResources
  nc.SelectSubforms = True             'Subforms
  nc.SelectViews = True                'Views

  Call nc.BuildCollection
  
  Set dxlExporter = objNotesSession.CreateDXLExporter

  dxl = dxlExporter.Export( nc )
  
  'forced to Shift-JIS
  tmpArr = Split( dxl, "xml version='1.0'" )
  dxl = Join( tmpArr, "xml version='1.0' encoding='" & encoding & "'" )
  
  'Wscript.Echo dxl
  Set file = fso.CreateTextFile( outputFilePath, True, False )
  file.Write( dxl )
  Wscript.Echo outputFilePath

  file.Close
  Set fso = Nothing
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


