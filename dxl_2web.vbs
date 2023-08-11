' dxl_2web.vbs
' Execute: c:\Windows\SysWOW64\CScript //nologo dxl_2web.vbs path/xxxxx.nsf
Option Explicit

Dim objWsh
Dim objArgs
Dim objDbPath
Dim objNotesSession
Dim objNotesDb
Dim objNotesView
Dim objNotesViewEntry
Dim objNotesViewEntryCollection
Dim objNotesDoc
Dim nc
Dim dxlExporter
Dim dxl, xml

Dim objXML
Dim nodeList, obj, unid, name, nodeList0, obj0
Dim docUnids, formNames
Dim vcc, cv

Dim outputFileFolder, outputXMLFolder
Dim outputFilePath, outputXMLPath
Dim fso
Dim file

Dim tmpArr, tmpStr, tmpInt, tmpBool, tmpV
Dim tmpDArr

Dim i, exportDoc

exportDoc = True
objDbPath = ""

Set objArgs = Wscript.Arguments

For i = 0 To objArgs.Count - 1
  If objArgs(i) = "-nodocs" Then
    exportDoc = False
  Else
    'local database filepath
    objDbPath = objArgs(i)
  End If
Next


If objDbPath = "" Then
  Wscript.Echo "Please specify local database path as command line parameter."
Else
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
  nc.SelectACL = False                  'ACL
  nc.SelectActions = False              'Action
  nc.SelectAgents = False               'Agents
  nc.SelectDatabaseScript = False       'DatabaseScript
  nc.SelectDataConnections = False      'DataConnections
  nc.SelectDocuments = exportDoc        'Documents
  nc.SelectFolders = True               'Folders
  nc.SelectForms = True                 'Forms
  nc.SelectFrameSets = False            'Framesets
  nc.SelectHelpAbout = True             'HelpAbout
  nc.SelectHelpIndex = True             'HelpIndex
  nc.SelectHelpUsing = True             'HelpUsing
  nc.SelectIcon = True                  'Icon
  nc.SelectImageResources = True        'ImageResources
  nc.SelectJavaResources = False        'JavaResources
  nc.SelectMiscCodeElements = False     'MiscCodeElements
  nc.SelectMiscFormatElements = False   'MiscFormatElements
  nc.SelectMiscIndexElements = False    'MiscIndexElements
  nc.SelectNavigators = False           'Navigators
  nc.SelectOutlines = False             'Outlines
  nc.SelectPages = False                'Pages
  nc.SelectProfiles = False             'Profiles
  nc.SelectReplicationFormulas = False  'ReplicationFormulas
  nc.SelectScriptLibraries = False      'ScriptLibraries
  nc.SelectSharedFields = True          'SharedFields
  nc.SelectStyleSheetResources = False  'StyleSheetResources
  nc.SelectSubforms = True              'Subforms
  nc.SelectViews = True                 'Views

  Call nc.BuildCollection
  
  Set dxlExporter = objNotesSession.CreateDXLExporter
  dxlExporter.OutputDOCTYPE = False

  dxl = dxlExporter.Export( nc )
  
  'forced to Shift-JIS
  tmpArr = Split( dxl, "xml version='1.0'" )
  dxl = Join( tmpArr, "xml version='1.0' encoding='SHIFT_JIS'" )
  
  'XML
  Set objXML = WScript.CreateObject( "MSXML2.DOMDocument" )
  tmpBool = objXML.loadXML( dxl )
  If tmpBool = True Then
    
    'Forms
    outputXMLFolder = outputFileFolder & "\forms" 
    If fso.FolderExists( outputXMLFolder ) = False Then
      fso.CreateFolder( outputXMLFolder )
    End If
    
    formNames = ""
    Set nodeList = objXML.DocumentElement.selectNodes( "/database/form" )
    For Each obj In nodeList
      'Wscript.Echo obj.nodeName '"form"
      'Wscript.Echo obj.xml

      xml = obj.xml
      tmpArr = Split( xml, " xmlns=""http://www.lotus.com/dxl""" )
      xml = Join( tmpArr, "" )

      unid = GetUNID( obj )
      
      outputXMLPath = outputXMLFolder & "\" & unid & ".xml"
      Set file = fso.CreateTextFile( outputXMLPath, True, False )
      file.Write( "<?xml version='1.0' encoding='SHIFT_JIS'?>" & xml )
      file.Close
      
      Dim formName
      formName = GetName( obj )
      formNames = formNames & formName & "," & unid & vbCrLf
      formName = GetAlias( obj )
      If formName <> "" Then
        tmpArr = Split( formName, "|" )
        For i = LBound(tmpArr) To UBound(tmpArr)
          tmpStr = Trim(tmpArr(i))
          formNames = formNames & tmpStr & "," & unid & vbCrLf
        Next
      End If
    Next
      
    outputXMLPath = outputXMLFolder & "\formnames.csv"
    Set file = fso.CreateTextFile( outputXMLPath, True, False )
    file.Write( formNames )
    file.Close

    'Documents
    outputXMLFolder = outputFileFolder & "\documents" 
    If fso.FolderExists( outputXMLFolder ) = False Then
      fso.CreateFolder( outputXMLFolder )
    End If
    
    Set nodeList = objXML.DocumentElement.selectNodes( "/database/document" )
    For Each obj In nodeList
      'Wscript.Echo obj.nodeName '"document"
      'Wscript.Echo obj.xml
      
      '"xmlns" 属性を削除する
      xml = obj.xml
      tmpArr = Split( xml, " xmlns=""http://www.lotus.com/dxl""" )
      xml = Join( tmpArr, "" )
      
      unid = GetUNID( obj )
      
      outputXMLPath = outputXMLFolder & "\" & unid & ".xml"
      Set file = fso.CreateTextFile( outputXMLPath, True, False )
      file.Write( "<?xml version='1.0' encoding='SHIFT_JIS'?>" & xml )
      file.Close
    Next
    
    'Views
    outputXMLFolder = outputFileFolder & "\views" 
    If fso.FolderExists( outputXMLFolder ) = False Then
      fso.CreateFolder( outputXMLFolder )
    End If
    
    Set nodeList = objXML.DocumentElement.selectNodes( "/database/view" )
    For Each obj In nodeList
      'Wscript.Echo obj.nodeName '"view"
      'Wscript.Echo obj.xml

      xml = obj.xml
      tmpArr = Split( xml, " xmlns=""http://www.lotus.com/dxl""" )
      xml = Join( tmpArr, "" )

      unid = GetUNID( obj )
      
      outputXMLPath = outputXMLFolder & "\" & unid & ".xml"
      Set file = fso.CreateTextFile( outputXMLPath, True, False )
      file.Write( "<?xml version='1.0' encoding='SHIFT_JIS'?>" & xml )
      file.Close
      
      docUnids = ""
      name = GetName( obj )
      Set objNotesView = objNotesDb.getView( name )
      Set objNotesViewEntryCollection = objNotesView.AllEntries
      vcc = 1
      
      Set objNotesDoc = objNotesView.GetFirstDocument()
      While Not ( objNotesDoc Is Nothing )
        If docUnids = "" Then
          docUnids = objNotesDoc.UniversalID
        Else
          docUnids = docUnids & vbCrLf & objNotesDoc.UniversalID
        End If
        
        Set objNotesViewEntry = objNotesViewEntryCollection.getNthEntry( vcc )
        cv = objNotesViewEntry.ColumnValues
        docUnids = docUnids & "," & Join( cv, "," )
        
        Set objNotesDoc = objNotesView.GetNextDocument( objNotesDoc )
        vcc = vcc + 1
      Wend
      
      docUnids = docUnids & vbCrLf
      outputXMLPath = outputXMLFolder & "\" & unid & ".csv"
      Set file = fso.CreateTextFile( outputXMLPath, True, False )
      file.Write( docUnids )
      file.Close
    Next
    
    'Folders
    outputXMLFolder = outputFileFolder & "\folders" 
    If fso.FolderExists( outputXMLFolder ) = False Then
      fso.CreateFolder( outputXMLFolder )
    End If
    
    Set nodeList = objXML.DocumentElement.selectNodes( "/database/folder" )
    For Each obj In nodeList
      'Wscript.Echo obj.nodeName '"folder"
      'Wscript.Echo obj.xml

      xml = obj.xml
      tmpArr = Split( xml, " xmlns=""http://www.lotus.com/dxl""" )
      xml = Join( tmpArr, "" )

      unid = GetUNID( obj )
      
      outputXMLPath = outputXMLFolder & "\" & unid & ".xml"
      Set file = fso.CreateTextFile( outputXMLPath, True, False )
      file.Write( "<?xml version='1.0' encoding='SHIFT_JIS'?>" & xml )
      file.Close
      
      docUnids = ""
      name = GetName( obj )
      Set objNotesView = objNotesDb.getView( name )
      Set objNotesViewEntryCollection = objNotesView.AllEntries
      vcc = 1
      
      Set objNotesDoc = objNotesView.GetFirstDocument()
      While Not ( objNotesDoc Is Nothing )
        If docUnids = "" Then
          docUnids = objNotesDoc.UniversalID
        Else
          docUnids = docUnids & vbCrLf & objNotesDoc.UniversalID
        End If
        
        Set objNotesViewEntry = objNotesViewEntryCollection.getNthEntry( vcc )
        cv = objNotesViewEntry.ColumnValues
        docUnids = docUnids & "," & Join( cv, "," )
        
        Set objNotesDoc = objNotesView.GetNextDocument( objNotesDoc )
        vcc = vcc + 1
      Wend
      
      docUnids = docUnids & vbCrLf
      outputXMLPath = outputXMLFolder & "\" & unid & ".csv"
      Set file = fso.CreateTextFile( outputXMLPath, True, False )
      file.Write( docUnids )
      file.Close
    Next
    
    'Sharedfields
    outputXMLFolder = outputFileFolder & "\sharedfields" 
    If fso.FolderExists( outputXMLFolder ) = False Then
      fso.CreateFolder( outputXMLFolder )
    End If
    
    Set nodeList = objXML.DocumentElement.selectNodes( "/database/sharedfield" )
    For Each obj In nodeList
      'Wscript.Echo obj.nodeName '"sharedfield"
      'Wscript.Echo obj.xml

      xml = obj.xml
      tmpArr = Split( xml, " xmlns=""http://www.lotus.com/dxl""" )
      xml = Join( tmpArr, "" )

      unid = GetUNID( obj )
      
      outputXMLPath = outputXMLFolder & "\" & unid & ".xml"
      Set file = fso.CreateTextFile( outputXMLPath, True, False )
      file.Write( "<?xml version='1.0' encoding='SHIFT_JIS'?>" & xml )
      file.Close
    Next
    
    'Subforms
    outputXMLFolder = outputFileFolder & "\subforms" 
    If fso.FolderExists( outputXMLFolder ) = False Then
      fso.CreateFolder( outputXMLFolder )
    End If
    
    Set nodeList = objXML.DocumentElement.selectNodes( "/database/subform" )
    For Each obj In nodeList
      'Wscript.Echo obj.nodeName '"subform"
      'Wscript.Echo obj.xml

      xml = obj.xml
      tmpArr = Split( xml, " xmlns=""http://www.lotus.com/dxl""" )
      xml = Join( tmpArr, "" )

      unid = GetUNID( obj )
      
      outputXMLPath = outputXMLFolder & "\" & unid & ".xml"
      Set file = fso.CreateTextFile( outputXMLPath, True, False )
      file.Write( "<?xml version='1.0' encoding='SHIFT_JIS'?>" & xml )
      file.Close
    Next
    
    'Imageresources
    outputXMLFolder = outputFileFolder & "\imageresources" 
    If fso.FolderExists( outputXMLFolder ) = False Then
      fso.CreateFolder( outputXMLFolder )
    End If
    
    Set nodeList = objXML.DocumentElement.selectNodes( "/database/imageresource" )
    For Each obj In nodeList
      'Wscript.Echo obj.nodeName '"imageresource"
      'Wscript.Echo obj.xml

      xml = obj.xml
      tmpArr = Split( xml, " xmlns=""http://www.lotus.com/dxl""" )
      xml = Join( tmpArr, "" )

      unid = GetUNID( obj )
      
      outputXMLPath = outputXMLFolder & "\" & unid & ".xml"
      Set file = fso.CreateTextFile( outputXMLPath, True, False )
      file.Write( "<?xml version='1.0' encoding='SHIFT_JIS'?>" & xml )
      file.Close
    Next

    'Icons
    outputXMLFolder = outputFileFolder & "\icons" 
    If fso.FolderExists( outputXMLFolder ) = False Then
      fso.CreateFolder( outputXMLFolder )
    End If
    
    Set nodeList = objXML.DocumentElement.selectNodes( "/database/note" )
    For Each obj In nodeList
      'Wscript.Echo obj.nodeName '"note"
      'Wscript.Echo obj.xml
      Set nodeList0 = obj.selectNodes( "@class" )
      For Each obj0 In nodeList0
        If obj0.text = "icon" Then

          xml = obj.xml
          tmpArr = Split( xml, " xmlns=""http://www.lotus.com/dxl""" )
          xml = Join( tmpArr, "" )

          unid = GetUNID( obj )
      
          outputXMLPath = outputXMLFolder & "\" & unid & ".xml"
          Set file = fso.CreateTextFile( outputXMLPath, True, False )
          file.Write( "<?xml version='1.0' encoding='SHIFT_JIS'?>" & xml )
          file.Close
        End If
      Next
    Next
  Else
    WScript.Echo objXML.ParseError.errorCode
    WScript.Echo objXML.ParseError.reason
  End If
  Set objXML = Nothing
  
  
  'Wscript.Echo dxl
  'Set file = fso.CreateTextFile( outputFilePath, True, False )
  'file.Write( dxl )
  'Wscript.Echo outputFilePath
  'file.Close
  
  Set fso = Nothing
  Set objNotesSession = Nothing
End If

Function GetUNID( o )
  Dim uid
  Dim nodeList, obj
  
  uid = ""
  Set nodeList = o.selectNodes( "noteinfo/@unid" )
  For Each obj in nodeList
    uid = obj.text
  Next
  
  GetUNID = uid
End Function

Function GetName( o )
  Dim n
  Dim nodeList, obj
  
  n = ""
  Set nodeList = o.selectNodes( "@name" )
  For Each obj in nodeList
    n = obj.text
  Next
  
  GetName = n
End Function

Function GetAlias( o )
  Dim a
  Dim nodeList, obj
  
  a = ""
  Set nodeList = o.selectNodes( "@alias" )
  For Each obj in nodeList
    a = obj.text
  Next
  
  GetAlias = a
End Function

Function GetForm( o )
  Dim f
  Dim nodeList, obj
  
  n = ""
  Set nodeList = o.selectNodes( "@form" )
  For Each obj in nodeList
    f = obj.text
  Next
  
  GetForm = f
End Function

