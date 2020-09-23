'
'
'  How To: Populate SQL Database from XML Files
'  (C) 2002 Berardi Michele
'
'


'
' Parser start declarations
'
' global variables
'

Dim g_oSwitches,TotParametri
Set g_oSwitches = CreateObject("Scripting.Dictionary")
g_oSwitches.CompareMode = vbTextcompare
'
' Parser end declarations
'



'
' Usefull declaration for DB connection
'
 
Dim MM_EURIS_STRING, MM_EURIS_MYSERVER_STRING '
 
Dim MM_editConnection '
Dim MM_editCmd, MM_editQuery '
Dim MM_editTable, MM_tableValues, MM_dbValues, MM_dbValues_MYDATABASE_DBNAME '
Dim MM_Last_DBNAME_N_COD_TKN
'
' generic variables
'

Dim FsysObj,FsysObj_FOpen,FileCloning,hdlFileLog
Dim Path,SearchDir,TargetDir,ConvertFile,CloneFile,BanRes,BatOutput,PrgTitle
Dim CURRENT_FILE, CURRENT_FILE_SESSION_LIMIT, START_FROM_FILE ' contatore file corrente

'
' XML nodes definition , i think you can use some tips for generalizing this one
' for example using 2 comma separated values strings one for the xml fileds and
' one for the respective sql filds , then parsing them... simple....
' ...the purpose of this source is to explain how use some xml objects.
'

Dim NODE_notice
'
Dim NODE_notice_metadata
Dim NODE_notice_metadata_category,NODE_notice_metadata_ambit,NODE_notice_metadata_priority
'
Dim NODE_notice_title
Dim NODE_notice_text
Dim NODE_notice_text_p ' subfileds of text...
Dim NODE_notice_firm
Dim NODE_notice_date
Dim NODE_notice_hour
'
Dim NODE_notice_metadata_priority_value
Dim NODE_notice_title_value
Dim NODE_notice_text_value
Dim NODE_notice_firm_value
Dim NODE_notice_date_value
Dim NODE_notice_hour_value
'

'
' the name of xml nodes - start definition
'
NODE_notice = "notice"
NODE_notice_metadata = "metadata"
NODE_notice_metadata_category = "category"
NODE_notice_metadata_ambit = "ambit"

NODE_notice_metadata_priority = "priority"
NODE_notice_title = "title"
NODE_notice_text = "text"
NODE_notice_text_p = "p"
NODE_notice_firm = "firm"
NODE_notice_date = "data"
NODE_notice_hour = "hour"
'
' the name of xml nodes - end definition
'

Dim xDoc ' As MSXML2.DOMDocument

Set xDoc = CreateObject("MSXML2.DomDocument.3.0") 'Set xDoc = New MSXML2.DOMDocument


'
' internal variables - start
'
         PrgTitle = "XML to SQL"
             Path = "C:\WINDOWS\Desktop\Xml\Pol\"
        SearchDir = ""
        TargetDir = ""
      ConvertFile = "convert____null.xml"
        CloneFile = "clone______null.xml"
          BanRes = ""
       BatOutput = "XMLtoSQL.log"
 START_FROM_FILE = 0
'
'internal variables - end
'  


'
' program start here..
'



'
'
'




Call ParseArgs

list = "Lista Argomenti Passati: " & vbNewline
For Each Item In g_oSwitches
	list = list & "  " & Item & "=" & g_oSwitches(Item) & vbNewline
Next

'If TotParametri < 1 Then
'   	list = "Nessun Argomento Passato"
'       wscript.echo list
'End If



'parameter validation

AssignandValidateArgs

'wscript.echo vbNewline & "Controllo Eccezioni:"

If ucase(g_oSwitches("Path")) <> "" Then
	Wscript.echo "  String passed to switch: Path "
	Path = g_oSwitches("Path")
Else
	Wscript.echo "  Please specify a path for the switch: Path "
End If

If ucase(g_oSwitches("DBConnect")) <> "" Then

MM_EURIS_MYSERVER_STRING = g_oSwitches("DBConnect")

	Wscript.echo "  DBConnect Passed"
Else
	Wscript.echo "  DBConnect not passed"
End If

If g_oSwitches("VectVals") <> "" Then
	multivaluelist = vbNewline & vbNewline & "sub arguments of VecVals:" & vbNewline
	multivaluevalues=split(g_oSwitches("VectVals"),";")
	For x = 0 To UBound(multivaluevalues)
		multivaluelist = multivaluelist & "  " & multivaluevalues(x) & vbNewline
	Next
	wscript.echo multivaluelist
End If




'
'
'

'If g_oSwitches.count < 1 Then
Call GetInput()
'end If




Sub GetInput()

	While Question<>6 
	    Path=InputBox("(0) Specify the Path where are the XML files:",PrgTitle,Path)
       SearchDir=InputBox("(1) Sub folder of path (0) where are the XML files archived by  Year:",PrgTitle,SearchDir)
       TargetDir=InputBox("(2) Sub folder of path (1) where are the XML files archived by Month:",PrgTitle,TargetDir)
     ConvertFile=InputBox("model file used for the XML>>SQL conversion:",PrgTitle,ConvertFile)
       CloneFile=InputBox("XML/SQL template:",PrgTitle,CloneFile)
          BanRes=InputBox("Unused Feature (On/Off):",PrgTitle,BanRes)
       BatOutput=InputBox("Name of the generated Batch:",PrgTitle,BatOutput)
 START_FROM_FILE=InputBox("Last resume Index:",PrgTitle,START_FROM_FILE)

START_FROM_FILE=CInt(START_FROM_FILE)

       Question=  MsgBox("The parameters are corrects?",vbYesNoCancel + vbInformation,PrgTitle)
    
		If Question = vbCancel Then
    			WScript.Quit
    		End If
	Wend

Operation()

end sub

Sub Operation()

	Set FsysObj = CreateObject("Scripting.FileSystemObject")
	set FileCloning = FsysObj.OpenTextFile(BatOutput,2,true)

	If Path="" then
		Call Get_Drivers(Path,SearchDir,TargetDir)
	Else
		Call PopolatoreXML_to_SQL(Path,ConvertFile,CloneFile,BanRes)
	End if

                                                               FileCloning.close

	Question =  MsgBox("Would you like to do more conversions?",vbOKCancel + vbInformation,PrgTitle)

 If Question=1 Then
	Call  GetInput()
 End If

End Sub

Sub Get_Drivers(Path,SearchDir,TargetDir)

Dim SysDrivers,SingleDrive
Set SysDrivers = FsysObj.Drives

	For each SingleDrive in SysDrivers

		If SingleDrive.DriveType = 2 or SingleDrive.DriveType = 3 Then

			Call Get_Folder(SingleDrive.path & "/",SearchDir,TargetDir)
		end if
	Next

end Sub

Sub Get_Folder(CurrentDrive,SearchDir,TargetDir)
dim Folder,SubFolder,SingleFolder
set Folder = FsysObj.GetFolder(CurrentDrive)  
set SubFolder = Folder.SubFolders

	for each SingleFolder in SubFolder

	MsgBox SingleFolder.name

		if SingleFolder.name = SearchDir then
		' sub folders scan using some recursion....
			Call PopolatoreXML_to_SQL(SingleFolder.path&"\"&TargetDir&"\",ConvertFile,CloneFile,BanRes)
		end if

	Call Get_Folder(SingleFolder.path,SearchDir,TargetDir)

	next  

End Sub

Sub PopolatoreXML_to_SQL(Folder,ConvertFile,CloneFile,BanRes)

Dim FoldScan,FoldFiles,FileName
set FoldScan = FsysObj.GetFolder(Folder)
set FoldFiles = FoldScan.Files

FileExtension=FsysObj.GetExtensionName(CloneFile)
FileExtension=lcase(FileExtension)

'
' REMMATO
'
	If CloneFile<>"" then
'If (FileExtension="gif") and (ConvertFile<>"") then
'Conversion ="/b "+ConvertFile+"+"
'else
'Conversion =""
'End if
'
		Else

			If BanRes="Hight" then
				CloneFile=Folder+"xmlpopdefcfgOFF.xml"
			Else
				CloneFile=Folder+"xmlpopdefcfgON.xml"
			End if

		End if


'
' XML extraction start here....
'

CURRENT_FILE = 1

	for each FileName in FoldFiles

	FileExtension=FsysObj.GetExtensionName(FileName.path)
	FileExtension=lcase(FileExtension)

'
' the only xml that we cannot use are the xml used as configuration schemes.
'

		If CURRENT_FILE > START_FROM_FILE Then
			If (FileExtension="xml") and ( (FileName.name<>"xmlpopdefcfgON.xml") and (FileName.name<>"xmlpopdefcfgOFF.xml")) then
			'FileCloning.write ">>>" & Conversion&CloneFile & " " & Folder & FileName.name & chr(13) & chr(10)
			FileCloning.write "-------------------------------------"  & chr(13) & chr(10)
			FileCloning.write " >>> START EXPORTING FILE XML N° " & CURRENT_FILE & " <<< "  & chr(13) & chr(10)
			FileCloning.write " >>> " & Folder & FileName.name & chr(13) & chr(10)
			LoadDocument(FileName.Path)
			FileCloning.write " >>> END   EXPORTING FILE XML N°  " & CURRENT_FILE & " <<< "  & chr(13) & chr(10)
			FileCloning.write "-------------------------------------"  & chr(13) & chr(10)

'
'if you would like to delete the parsed file unrem here...
'
' FsysObj.DeleteFile(Folder & FileName.name)
'
'
'

			End If
		Else
		'
		' skipping file.....
		'
		End If

	CURRENT_FILE = CURRENT_FILE + 1



	Next

End Sub



Sub LoadDocument(Percorso_File_XML)

	xDoc.validateOnParse = False

	If xDoc.Load(Percorso_File_XML) Then


 ' document loaded in the collection

 ' slq population
 
       MM_EURIS_STRING = "Provider=SQLOLEDB;dsn=EURIS;uid=MYUSERNAME;pwd=MYPASSWORD;" '
MM_EURIS_MYSERVER_STRING = "Provider=SQLOLEDB;Server=MYSERVER;database=MYDATABASE;uid=MYUSERNAME;pwd=MYPASSWORD;" '
 

'
' setup of the db connection
'
 
MM_editConnection = MM_EURIS_MYSERVER_STRING ' settiamola una sola volta è ridontante per ogni insert / select

' sql population

'
' population in: "MYDATABASE.DBNAME" 
'
' fake insert for a subsequent xml population...
'
MM_editTable = "MYDATABASE.DBNAME" '
MM_tableValues = "      N_COD_PRT,      C_DES_TIT,      C_DES_TXT,      C_COD_FIR,       D_DAT_NTZ,       C_ORA_NTZ" '
MM_dbValues = "0,'C_DES_TIT_value','C_DES_TXT_value','BM/AS/YourCompany','19-OCT-1974','00:01'" '
MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"
'
Call QueryInToDbTable(MM_editConnection, MM_editQuery)
'
' population in: "MYDATABASE.DBNAME" end section...
'

'
'recordset extraction: MYDATABASE.DBNAME begin section...
'
MM_editTable = "MYDATABASE.DBNAME"
MM_Last_DBNAME_N_COD_TKN = DBSelect(MM_editConnection, "SELECT max(" & MM_editTable & ".N_COD_TKN" & ")" & " FROM " & MM_editTable)

'
'recordset extraction: MYDATABASE.DBNAME end section...
'


'
'calling xml extraction cicle -- begin...
'
    DisplayNode xDoc.childNodes, 0
'
' calling xml extraction cicle -- end..
'

'
' 
' MM_dbValues_MYDATABASE_DBNAME generated by XML extraction cicle... 
'
'


'
' update string-...
'

 MM_dbValues_MYDATABASE_DBNAME = " N_COD_PRT = " & NODE_notice_metadata_priority_value
 MM_dbValues_MYDATABASE_DBNAME = MM_dbValues_MYDATABASE_DBNAME & " , " & " C_DES_TIT = '" & Replace(NODE_notice_title_value,"'","''") & "' "
 MM_dbValues_MYDATABASE_DBNAME = MM_dbValues_MYDATABASE_DBNAME & " , " & " C_DES_TXT = '" & Replace(NODE_notice_text_value,"'","''") & "' "
 MM_dbValues_MYDATABASE_DBNAME = MM_dbValues_MYDATABASE_DBNAME & " , " & " C_COD_FIR = '" & Replace(NODE_notice_firm_value,"'","''") & "' "
 MM_dbValues_MYDATABASE_DBNAME = MM_dbValues_MYDATABASE_DBNAME & " , " & " D_DAT_NTZ = '" & NODE_notice_date_value & "' "
 MM_dbValues_MYDATABASE_DBNAME = MM_dbValues_MYDATABASE_DBNAME & " , " & " C_ORA_NTZ = '" & NODE_notice_hour_value & "' "

'
' -fine- costruzione della stringa di update
'

'
' update in: "MYDATABASE.DBNAME" 
'
MM_editTable = "MYDATABASE.DBNAME" '
MM_dbValues = MM_dbValues_MYDATABASE_DBNAME
MM_editQuery = "update " & MM_editTable &  " set " & MM_dbValues & " where " & MM_editTable& ".N_COD_TKN = " & MM_Last_DBNAME_N_COD_TKN

FileCloning.write "- - - - - - - - - - - - - - - - - - -"  & chr(13) & chr(10)
FileCloning.write ">>> SQL: " & MM_editQuery & chr(13) & chr(10)
FileCloning.write "- - - - - - - - - - - - - - - - - - -"  & chr(13) & chr(10)

'
Call QueryInToDbTable(MM_editConnection, MM_editQuery)

'
NODE_notice_text_value = ""
'
'
' -fine- update in: "MYDATABASE.DBNAME"
'

Else
    ' Il documento non è stato caricato.
    ' Consultare l'elenco precedente per informazioni sull'errore.
    Dim strErrText
'    Dim xPE As xDoc.IXMLDOMParseError
    ' Ottenere l'oggetto ParseError
    Set xPE = xDoc.parseError
    With xPE
        strErrText = "Your XML Document failed to load" & _
        "due the following error." & vbCrLf & _
        "Error #: " & .errorCode & ": " & xPE.reason & _
        "Line #: " & .Line & vbCrLf & _
        "Line Position: " & .linepos & vbCrLf & _
        "Position In File: " & .filepos & vbCrLf & _
        "Source Text: " & .srcText & vbCrLf & _
        "Document URL: " & .Url
    End With
    
    MsgBox strErrText, vbExclamation
    Set xPE = Nothing
    
End If
End Sub

Sub DisplayNode(ByRef Nodes , ByVal Indent) ' Sub DisplayNode(ByRef Nodes As MSXML2.IXMLDOMNodeList, ByVal Indent As Integer)

    Dim xNode '    set xNode = xDoc.IXMLDOMNode

    Indent = Indent + 2 ' ??????????????

' 
' -inizio- XML document node recursion...
'
    For Each xNode In Nodes
      
'MsgBox xNode.parentNode.nodeName & " 1N: " & xNode.nodeValue     
'MsgBox xNode.parentNode.nodeName & " 1V= " & xNode.nodeValue  
                      
      ' If xNode.nodeType = NODE_notice_title Then
      ' msgbox Indent & xNode.parentNode.nodeName & ":" & xNode.nodeValue     
      ' msgbox xNode.parentNode.nodeName & " = " & xNode.nodeValue  
     
'
' -inizio- UPDATE COSTRUCTION PASSED TO THIS VARIABILE: MM_dbValues_MYDATABASE_DBNAME
'
If xNode.parentNode.nodeName = NODE_notice_metadata_priority Then
NODE_notice_metadata_priority_value = xNode.nodeValue
	Else If xNode.parentNode.nodeName = NODE_notice_title Then
	NODE_notice_title_value = xNode.nodeValue
'Else If xNode.nodeName = NODE_notice_text Then
'NODE_notice_text_value = xNode.nodeValue
		Else If xNode.parentNode.nodeName = NODE_notice_text_p Then
		NODE_notice_text_value = NODE_notice_text_value & vbCrLf & xNode.nodeValue
			Else If xNode.parentNode.nodeName = NODE_notice_firm Then
			NODE_notice_firm_value = xNode.nodeValue
				Else If xNode.parentNode.nodeName = NODE_notice_date Then
				NODE_notice_date_value = MeseToMonth(xNode.nodeValue)
					Else If xNode.parentNode.nodeName = NODE_notice_hour Then
					NODE_notice_hour_value = xNode.nodeValue
					End If
				End If
'End If
			End If
		End If 
	End If
End If
'
' -fine- COSTRUISCE L'UPDATE DA PASSARE A SQL IN QUESTA VARIABILE: MM_dbValues_MYDATABASE_DBNAME
'








'
' -inizio- inserimento in : "MYDATABASE.TADNCTG" - tabella contenente le voci di categoria di appartenenza -
'
If xNode.parentNode.nodeName = NODE_notice_metadata_category Then
'
MM_editTable = "MYDATABASE.TADNCTG" '
MM_tableValues = "                N_COD_TKN,      C_DES_CTG"
MM_dbValues = MM_Last_DBNAME_N_COD_TKN & ",'" & xNode.nodeValue & "'" 
MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

'MsgBox xNode.parentNode.nodeName & " 2N: " & xNode.nodeValue     
'MsgBox xNode.parentNode.nodeName & " 2V= " & xNode.nodeValue  
      
'
Call QueryInToDbTable(MM_editConnection, MM_editQuery)
'
End If 
'
' -fine- ciclo di inserimento in: "MYDATABASE.TADNCTG" - tabella contenente le voci di categoria di appartenenza -
'




'
'-inizio- inserimento in : "MYDATABASE.TADNAMBRGN" - tabella contenente le voci relative all' ambito
'
If xNode.parentNode.nodeName = NODE_notice_metadata_ambit Then
'
MM_editTable = "MYDATABASE.TADNAMBRGN" '
MM_tableValues = "                N_COD_TKN,      C_DES_AMB_RGN" '
MM_dbValues = MM_Last_DBNAME_N_COD_TKN & ",'" & xNode.nodeValue & "'"
MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

'MsgBox xNode.parentNode.nodeName & " 3N: " & xNode.nodeValue     
'MsgBox xNode.parentNode.nodeName & " 3V= " & xNode.nodeValue  

'
Call QueryInToDbTable(MM_editConnection, MM_editQuery)
'
End If
'
' -fine- inserimento in: "MYDATABASE.TADNAMBRGN" - tabella contenente le voci relative all' ambito
'






'
' -inizio- more nodes?
'
        If xNode.hasChildNodes Then
            DisplayNode xNode.childNodes, Indent
        End If
'
' -fine- more nodes?
'


    Next

'
' -fine- ciclo di scansione recursiva dei nodi del documento XML
'

End Sub





'
' QueryInToDbTable -inizio-
'
    Sub QueryInToDbTable(MM_editConnection, MM_editQuery)
 
Set MM_editCmd = CreateObject("ADODB.Command") '     CreateObject("SQLDMO.SQLServer")
 
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
 

Set MM_editCmd = Nothing
 
    End Sub
'
' QueryInToDbTable -fine-
'
 

Function DBSelect(MM_EURIS_STRING, Sql)
  Dim rsTemp
  DBSelect = ""
  Set rsTemp = CreateObject("ADODB.Recordset")
  rsTemp.ActiveConnection = MM_EURIS_STRING
  rsTemp.Source = Sql
  rsTemp.CursorType = 0
  rsTemp.CursorLocation = 2
  rsTemp.LockType = 1
  rsTemp.Open
  If Not rsTemp.EOF Then
     DBSelect = rsTemp.Fields(0)
  End If
End Function



Function MeseToMonth ( cData )
   Dim Mese
    Mese = Mid(cData, 4, 3)
    Select Case Mese
        Case "GEN"
            Mese = "JAN"
        'Case "FEB"
        'CASE "MAR"
        'CASE "APR"
        Case "MAG"
            Mese = "MAY"
        Case "GIU"
            Mese = "JUN"
        Case "LUG"
           Mese = "JUL"
        Case "AGO"
           Mese = "AUG"
        Case "SET"
           Mese = "SEP"
        Case "OTT"
            Mese = "OCT"
        'CASE "NOV"
        Case "DIC"
            Mese = "DEC"
    End Select

MeseToMonth = Left(CData,2) + "-" + Mese +"-" + Right(cData,2)

End Function


Function LogOpen(sNomFileLog)
    Set hdlFileLog = FsysObj.OpenTextFile(sNomFileLog, 2, True)
    hdlFileLog.Write "<html><head><title>XML to SQL (C) 2002 Berardi Michele</title></head><body bgcolor=""#C0C0C0""><p align=""center""><font color=""#0000FF""><big><big><big><strong> XML - Importer </strong></big></big></big></font></p>"
    hdlFileLog.Write "<p align=""center""><font color=""#FF0000""><big><big><strong>Log Errori</strong></big></big></font></p>"
    hdlFileLog.Write "<table border=""1"" width=""100%"" bordercolor=""#000000"">"
    hdlFileLog.Write "<tr>"
    hdlFileLog.Write "<td width=""15%"" bgcolor=""#808080""><font color=""#FFFFFF""><strong>Data</strong></font></td>"
    hdlFileLog.Write "<td width=""5%"" bgcolor=""#808080""><font color=""#FFFFFF""><strong>Pgm/Script</strong></font></td>"
    hdlFileLog.Write "<td width=""5%"" bgcolor=""#808080""><font color=""#FFFFFF""><strong>Tipo Errore</strong></font></td>"
    hdlFileLog.Write "<td width=""5%"" bgcolor=""#808080""><font color=""#FFFFFF""><strong>Nr. Errore</strong></font></td>"
    hdlFileLog.Write "<td width=""30%"" bgcolor=""#808080""><font color=""#FFFFFF""><strong>Descrizione</strong></font></td>"
    hdlFileLog.Write "<td width=""30%"" bgcolor=""#808080""><font color=""#FFFFFF""><strong>Origine</strong></font></td>"
    hdlFileLog.Write "<td width=""10%"" bgcolor=""#808080""><font color=""#FFFFFF""><strong>Routine</strong></font></td>"
    hdlFileLog.Write "</tr>"
End Function

Sub LogWrite(TipErrore, NumErrore, Description, ErrorSource, Routine)
  hdlFileLog.Write "<tr>"
  hdlFileLog.Write "<td width=""15%"" bgcolor=""#E8E8E8"" height=""20"">" & Now() & "</td>"                             ' Data
  hdlFileLog.Write "<td width=""5%"" bgcolor=""#E8E8E8"" height=""20""><small>XML to Sql </small></td>"  ' Pgm
  hdlFileLog.Write "<td width=""5%"" bgcolor=""#E8E8E8"" height=""20""><small>" & TipErrore & "</small></td>" ' Tipo Errore
  hdlFileLog.Write "<td width=""5%"" bgcolor=""#E8E8E8"" height=""20""><small>" & NumErrore & "</small></td>" ' Numero Errore
  hdlFileLog.Write "<td width=""30%"" bgcolor=""#E8E8E8"" height=""20""><small>" & Description & "</small></td>" ' Descrizione
  hdlFileLog.Write "<td width=""30%"" bgcolor=""#E8E8E8"" height=""20""><small>" & ErrorSource & "</small></td>" ' Origine
  hdlFileLog.Write "<td width=""10%"" bgcolor=""#E8E8E8"" height=""20""><small><font color=""#000000"">" & Routine & "</font></small></td>" ' Routine
  hdlFileLog.Write "</tr>"
End Sub

Sub LogClose()
    hdlFileLog.Write "</table>"
    hdlFileLog.Close
End Sub


Sub ParseArgs

On Error Resume Next
	Dim pair, list, sArg, Item
	For Each sArg In Wscript.Arguments

		pair = Split(sArg, "==", 2)

		'if value is specified multiple times, last one overwrite precedents...
		If g_oSwitches.Exists(Trim(pair(0))) Then
			g_oSwitches.Remove(Trim(pair(0)))
		End If

		If UBound(pair) >= 1 Then
			g_oSwitches.add Trim(pair(0)), Trim(pair(1))
		Else
			g_oSwitches.add Trim(pair(0)),""
		End If
	Next
End Sub


Sub AssignAndValidateArgs
	If (g_oSwitches.count < 1) Or (g_oSwitches.Exists("help")) or (g_oSwitches.Exists("/help")) or (g_oSwitches.Exists("?")) or (g_oSwitches.Exists("/?")) then
		wscript.echo "procedure uses:" &_
		vbNewLine & "--""--"" Example:" &_
		vbNewLine & "try this:" &_
		vbNewline & "  cscript XMLtoSQL.vbs" &_
		vbNewline & "  cscript XMLtoSQL.vbs /help" &_
		vbNewline & "  cscript XMLtoSQL.vbs Path== C:\XML Log== ImportSQLinXML.log DBConnect== Provider=SQLOLEDB;Server=MYSERVER;database=MYDATABASE;uid=MYUSERNAME;pwd=MYPASSWORD;"
'		Wscript.Quit(1)

TotParametri = g_oSwitches.count

	End If
End Sub



