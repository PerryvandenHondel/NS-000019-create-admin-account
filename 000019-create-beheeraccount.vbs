'
'	CreateBeheerAccount
'
'
'	Version	Description
'	=======	================================================================================
'	12		Use a database as source instead with an Excel sheet
'	11		New version; works from REC domain
'	10		Place the function of a beheerder in title field/Place ARGUS=xxxxx in description
'	09		UUpdate to use Class Splunk_06
'	08		Updates to use class Splunk_04
'	07		Added Splunk logging
'	06		?
'	05		New version that build command files (.CMD) with DSADD commands in it.
'	04		?
'	03		Make use full for multiple domains
'	02		Add Argus field for export to log file
'	01		Initial build of script.
'
''	CLASSES, FUNCTIONS AND SUBS:
''		Class ClassTextFile
''		Function AdGetRootDse
''		Function ConfigReadSettingSection
''		Function DsutilsGetDnFromSam
''		Function EncloseWithDQ
''		Function FixFirstname
''		Function FixLastName
''		Function FixMiddleName
''		Function GenerateAccountName
''		Function GeneratePassword
''		Function GetScriptNameMinVersion
''		Function GetScriptPath
''		Function GetTempFileName
''		Function RunCommand
''		Sub	ScriptInit
''		Sub CreateNewAccount
''		Sub DeleteFile
''		Sub DsmodGroupAddMember
''		Sub MakeSameAs
''		Sub RecordCreate
''		Sub RecordPrepare
''		Sub RecordInform		
''		Sub ScriptDone
''		Sub ScriptRun
''		Sub SetNotDelegatedFlag



Option Explicit



Const	TBL_NWA	= 				"new_account"
Const	FLD_NWA_ID =			"new_account_id"
Const	FLD_NWA_FNAME = 		"first_name"
Const	FLD_NWA_MNAME = 		"middle_name"
Const	FLD_NWA_LNAME = 		"last_name"
Const	FLD_NWA_TITLE = 		"title"
Const	FLD_NWA_SUPP_ID = 		"supplier_id"
Const	FLD_NWA_MOBILE = 		"mobile"
Const	FLD_NWA_DOMAIN = 		"domain_id"
Const	FLD_NWA_USERNAME = 		"user_name"
Const	FLD_NWA_USERNAME_SAME = "user_name_same"
Const	FLD_NWA_MAIL = 			"mail"
Const	FLD_NWA_PASSWORD = 		"password"
Const	FLD_NWA_REF = 			"reference_number"
Const	FLD_NWA_REQ_ID = 		"requestor_id"
Const	FLD_NWA_STATUS = 		"status_id"
Const	FLD_NWA_RCD = 			"rcd"
Const	FLD_NWA_RLU = 			"rlu"

Const	TBL_DMN = 				"new_account_domain"
Const	FLD_DMN_ID = 			"domain_id"
Const	FLD_DMN_UPN = 			"upn"
Const	FLD_DMN_NT = 			"domain_nt"
Const	FLD_DMN_OU = 			"org_unit"
Const	FLD_DMN_USE_SUPPLIER_OU = "use_supplier_ou"

Const	MAX_ACCOUNT_LEN		=	20

Const	ADS_NAME_INITYPE_GC =	3
Const	ADS_NAME_TYPE_NT4	=	3
Const	ADS_NAME_TYPE_1779	=	1

Const	FOR_READING			=	1
Const	FOR_WRITING			=	2
Const	FOR_APPENDING		=	8



Dim		gobjRootDse
Dim		gstrRootDse
Dim		gstrPathExcel
Dim		gobjExcel
Dim		gobjSheet
Dim		gobjFso
Dim		gobjFile
Dim		gstrDomainDns
Dim		gobjShell
Dim		tfNewAccount	'' Text File New Account; write all actions in this file.
Dim		db				'' Global object to connect to the database



Sub RecordSetStatus(ByVal intRecordId, ByVal intNewStatusId)
	''
	'' Update the status ID with a new ID, 
	'' in case of errors.
	''
	Dim 		qu
	
	qu = "UPDATE " & TBL_NWA & " "
	qu = qu & "SET "
	qu = qu & FLD_NWA_STATUS & "=" & intNewStatusId & ","
	qu = qu & FLD_NWA_RLU & "=" & db.FixDtm(Now()) & " "
	qu = qu & "WHERE " & FLD_NWA_ID & "=" & intRecordId & ";"
			
	db.ExecQuery(qu)
End Sub



Function GetDomainValues(ByVal strDomainId, ByVal strFieldName)
	Dim		qs
	Dim		rs
	Dim		r	'' Function result
	
	qs = "SELECT " & strFieldName & " "
	qs = qs & "FROM " & TBL_DMN & " "
	qs = qs & "WHERE " & FLD_DMN_ID & "=" & db.FixStr(strDomainId) & ";"
	
	Call db.GetRecordSet(rs, qs)
	If rs.Eof = True Then
		r = ""
	Else
		rs.MoveFirst
		r = rs(strFieldName).Value
	End If
	GetDomainValues = r
End Function '' of Function GetDomainValues



Function EncloseWithDQ(ByVal s)
	''
	''	Returns an enclosed string s with double quotes around it.
	''	Check for exising quotes before adding adding.
	''
	''	s > "s"
	''
	
	If Left(s, 1) <> Chr(34) Then
		s = Chr(34) & s
	End If
	
	If Right(s, 1) <> Chr(34) Then
		s = s & Chr(34)
	End If

	EncloseWithDQ = s
End Function '' of Function EncloseWithDQ



Sub SetNotDelegatedFlag(ByVal strDomainId, ByVal strUserName)
	''
	''	Set the not delegated flag in the UserAccountControl
	''	NOT_DELEGATED - When this flag is set, the security context of the user is not delegated to a service even if the service account is set as trusted for Kerberos delegation.
	''
	''	Source:	https://support.microsoft.com/en-us/kb/305144
	''
	Dim		c
	
	'adfind.exe -b "DC=prod,DC=ns,DC=nl" -f "sAMAccountName=BEH_WMIScanProject" userAccountControl -adcsv | admod userAccountControl::{{.:SET:1048576}}	
	c = "adfind.exe -b " & EncloseWithDQ(strDomainId) & " -f " & EncloseWithDQ("sAMAccountName=" & strUserName) & " userAccountControl -adcsv | admod.exe userAccountControl::{{.:SET:1048576}}"
	'c = "admod.exe -b " & EncloseWithDQ(strDn) & " userAccountControl::{{.:SET:1048576}}"
	WScript.Echo c
	'Call RunCommand(c)
End Sub '' of Sub SetNotDelegatedFlag



Sub RecordPrepare(ByVal intStatusId)
	''
	'' Prepare all records to fill in the gaps,
	'' with username, password etc.
	''
	Const	NEXT_STATUS = 100
	
	Dim		qs			'' Query Select
	Dim		qu			'' Query Update
	Dim		rs			'' RecordSet
	
	Dim		strUserName
	Dim		strFirstName
	Dim		strMiddleName
	Dim		strLastName
	Dim		intRecordId
	Dim		strSupplierId
	Dim		strPassword

	'' Build a select query to get all records with status 0 (all new records)
	qs = "SELECT * "
	qs = qs & "FROM " & TBL_NWA & " "
	qs = qs & "WHERE " & FLD_NWA_STATUS & "=" & intStatusId & ";"
	
	Call db.GetRecordSet(rs, qs)
	If rs.Eof = True Then
		WScript.Echo "NO RECORDS FOUND FOR STATUS " & intStatusId
		WScript.Echo "- " & qs
	Else
		WScript.Echo "RECORDS FOUND  FOR STATUS " & intStatusId
		WScript.Echo "- " & qs
		rs.MoveFirst
		While Not rs.EOF
			intRecordId = rs(FLD_NWA_ID).Value
			
			strFirstName = rs(FLD_NWA_FNAME).Value
			strMiddleName = rs(FLD_NWA_MNAME).Value
			strLastName = rs(FLD_NWA_LNAME).Value
			strSupplierId = rs(FLD_NWA_SUPP_ID).Value
			strUserName = GenerateAccountName(strSupplierId, strFirstName, strMiddleName, strLastName)
			
			strPassword = GeneratePassword()
			WScript.Echo "strPassword=" & strPassword
			
			'' The variables are filled with valid values.
			'' Update the table.
			qu = "UPDATE " & TBL_NWA & " "
			qu = qu & "SET "
			qu = qu & FLD_NWA_USERNAME & "=" & db.FixStr(strUserName) & ","
			qu = qu & FLD_NWA_PASSWORD & "=" & db.FixStr(strPassword) & ","
			qu = qu & FLD_NWA_STATUS & "=" & NEXT_STATUS & ","
			qu = qu & FLD_NWA_RLU & "=" & db.FixDtm(Now()) & " "
			qu = qu & "WHERE " & FLD_NWA_ID & "=" & intRecordId & ";"
			
			db.ExecQuery(qu)
			''WScript.Echo qu
			rs.MoveNext '' Next record
		Wend
		Set rs = Nothing
	End If
End Sub '' of Sub PrepareRecords



Sub RecordCreate(ByVal intStatusId)
	''
	'' Create new accounts when status is intStatusId
	''
	Const	NEXT_STATUS = 200
	
	Dim		c
	Dim		chrUseSupplierOu
	Dim		dn
	Dim		intRecordId
	Dim		qs			'' Query Select
	Dim		qu			'' Query Update
	Dim		rs			'' RecordSet
	Dim		strDescription
	Dim		strDomainId
	Dim		strFirstName
	Dim		strLastName
	Dim		strMiddleName
	Dim		strMobile
	Dim		strOrgUnit
	Dim		strPassword
	Dim		strReference
	Dim		strRequestor
	Dim		strRootDse
	Dim		strSupplierId
	Dim		strTitle
	Dim		strUserDn
	Dim		strUserName
	Dim		strUserNameSame
	Dim		strUserNameSameDn
	Dim		strUserUpn
	Dim		d			'' DN of Default groups.

	'' Build a select query to get all records with status 0 (all new records)
	qs = "SELECT * "
	qs = qs & "FROM " & TBL_NWA & " "
	qs = qs & "WHERE " & FLD_NWA_STATUS & "=" & intStatusId & ";"
	
	Call db.GetRecordSet(rs, qs)
	If rs.Eof = True Then
		WScript.Echo "NO RECORDS FOUND FOR STATUS " & intStatusId
		WScript.Echo "- " & qs
	Else
		WScript.Echo "RECORDS FOUND FOR STATUS " & intStatusId
		WScript.Echo "- " & qs
		rs.MoveFirst
		While Not rs.EOF
			intRecordId = rs(FLD_NWA_ID).Value
			WScript.Echo String(80, "-")
			WScript.Echo "RECORD ID="& intRecordId
			
			strUserName = rs(FLD_NWA_USERNAME).Value
			
			WScript.Echo "CREATE NEW ACCOUNT " & strUserName
			
			strRootDse = rs(FLD_NWA_DOMAIN).Value
			
			dn = DsutilsGetDnFromSam(strRootDse, "user", strUserName)
			
			If Len(dn) > 0 Then
				'' The result of DsutilsGetDnFromSam contains a DN path, so the account exist
				WScript.Echo "WARNING: Account " & strUserName & " already exists in " & strRootDse
				
				'' Account already exist in the domain.
				Call RecordSetStatus(intRecordId, intStatusId + 1)
			Else
				'' The account does not exist, create it.
				
				strDomainId = rs(FLD_NWA_DOMAIN).Value

				strOrgUnit = GetDomainValues(strDomainId, FLD_DMN_OU)
				
				strSupplierId = rs(FLD_NWA_SUPP_ID).Value
				
				chrUseSupplierOu = GetDomainValues(strDomainId, FLD_DMN_USE_SUPPLIER_OU)
				
				strUserUpn = LCase(strUserName & "@" & GetDomainValues(strDomainId, FLD_DMN_UPN))
				WScript.Echo "strUserUpn=" & strUserUpn
				
				WScript.Echo strOrgUnit
				WScript.Echo chrUseSupplierOu
				strUserDn = "CN=" & strUserName & ","
				If UCase(chrUseSupplierOu) = "Y" Then
					'' We are placing the user accounts under a supplier specific OU.
					strUserDn = strUserDn & "OU=" & rs(FLD_NWA_SUPP_ID).Value & ","
				End If
				strUserDn = strUserDn & strOrgUnit & "," & strDomainId
				WScript.Echo "User DN=" & strUserDn
				
				strReference = rs(FLD_NWA_REF).Value
				strRequestor = rs(FLD_NWA_REQ_ID).Value
				strDescription = "CALL=" & strReference & " REQUEST_BY=" & strRequestor
				WScript.Echo strDescription
				
				strPassword = rs(FLD_NWA_PASSWORD).Value
				
				strFirstName = rs(FLD_NWA_FNAME).Value
				strMiddleName = rs(FLD_NWA_MNAME).Value
				strLastName = rs(FLD_NWA_LNAME).Value
				strTitle = rs(FLD_NWA_TITLE).Value
				strMobile = rs(FLD_NWA_MOBILE).Value
				
				'	Build the command line to create a new account using DSADD.EXE
				c = "dsadd user " & EncloseWithDQ(strUserDn) & " "
				c = c & "-samid " & EncloseWithDQ(strUserName) & " "
				c = c & "-pwd " & EncloseWithDQ(strPassword) & " "
				c = c & "-fn " & EncloseWithDQ(Trim(strFirstName & " " & strMiddleName)) & " "
				c = c & "-ln " & EncloseWithDQ(strLastName) & " "
				c = c & "-title " & EncloseWithDQ(strTitle) & " "
				c = c & "-desc " & EncloseWithDQ(strDescription) & " "
				c = c & "-display " & EncloseWithDQ(strUserName) & " "
				c = c & "-upn " & EncloseWithDQ(strUserUpn) & " "
				
				If Len(strMobile) > 0 Then
					c = c & "-mobile " & EncloseWithDQ(strMobile) & " "
				End If
				c = c & "-company " & EncloseWithDQ(strSupplierId) & " "
				c = c & "-mustchpwd yes"
				
				'' Execute the DSADD command to create the account
				WScript.Echo c
				gobjShell.Run "cmd /c " & c, 0, True
				
				'' Sleap for 2 seconds before continuing
				WScript.Sleep 2000
				
				'' Add account to default groups
				If strDomainId = "DC=prod,DC=ns,DC=nl" Then
					'' Add account to the SmartXS groups, only for PROD.
					
					'' Add default group ROL_SmartXS_Autorisatie
					d = DsutilsGetDnFromSam(strDomainId, "group", "RP_SmartXS_Autorisatie") ' Was ROL_SmartXS_Autorisatie
					Call DsmodGroupAddMember(d, strUserDn)

					'' Add default group for BONS Autorisatie
					d = DsutilsGetDnFromSam(strDomainId, "group", "RP_BONS_Autorisatie") ' 2014-12-29 PVDH Kan de groep niet vinden in PROD.NS.NL
					Call DsmodGroupAddMember(d, strUserDn)
				End If
				
				'' Set the not delegated flag on the user account by poking UserAccountControl
				Call SetNotDelegatedFlag(strDomainId, strUserName)			

				'' Make the account similar as a reference account
				strUserNameSame = rs(FLD_NWA_USERNAME_SAME).Value
				If Len(strUserNameSame) > 0 Then
					'' A reference account to copy the groups from has been provided
					
					'' Get the DN of the source user
					strUserNameSameDn = DsutilsGetDnFromSam(strDomainId, "user", strUserNameSame)
					
					'' Now copy the groups from strUserNameSameDn to strUserDn
					Call MakeSameAs(strUserNameSameDn, strUserDn)
				Else
					WScript.Echo "INFO: No source account was provided so the account has no ROLE groups at this moment"
				End If
			End If
			
			Call RecordSetStatus(intRecordId, NEXT_STATUS)
			rs.MoveNext '' Next record
		Wend
		Set rs = Nothing
	End If
End Sub '' of Sub RecordCreate



Sub RecordInform(ByVal intStatusId)
	''
	''	Process al records that have status intStatusId
	''
	''
	
	Const	NEXT_STATUS = 900
	
	Dim		qs			'' Query Select
	Dim		qu			'' Query Update
	Dim		rs			'' RecordSet
	Dim		intRecordId
	

	qs = "SELECT * "
	qs = qs & "FROM " & TBL_NWA & " "
	qs = qs & "WHERE " & FLD_NWA_STATUS & "=" & intStatusId & ";"
	
	Call db.GetRecordSet(rs, qs)
	If rs.Eof = True Then
		WScript.Echo "NO RECORDS FOUND FOR STATUS " & intStatusId
		WScript.Echo "- " & qs
	Else
		WScript.Echo "RECORDS FOUND FOR STATUS " & intStatusId
		WScript.Echo "- " & qs
		rs.MoveFirst
		While Not rs.EOF
			intRecordId = rs(FLD_NWA_ID).Value
			
			
		
		
			Call RecordSetStatus(intRecordId, NEXT_STATUS)
			rs.MoveNext '' Next record
		Wend
		Set rs = Nothing
	End If
End Sub '' of Sub RecordInfom


Sub	ScriptInit()
	Dim	strRootLog
	
	' Get the current RootDSE location.
	'gstrRootDse = AdGetRootDse
	
	
	'gstrRootDse = "DC=acceptatie,DC=ns,DC=nl"
	'WScript.Echo "gstrRootDse=" & gstrRootDse
	
	' 2015-02-19 PVDH: Use file 000019-input.xls
	'gstrPathExcel = GetScriptPath & "000019-input.xlsx"
	'gstrPathExcel = GetScriptPath & "input.xlsx"
	
	'WScript.Echo "Using Excel file: " & gstrPathExcel
	
	'Set gobjExcel = CreateObject("Excel.Application")
	'Set gobjSheet = gobjExcel.Workbooks.Open(gstrPathExcel)
	
	Set gobjFso = CreateObject("Scripting.FileSystemObject")
	
	Set gobjShell = CreateObject("WScript.Shell")
	
	'strRootLog = "\\nsd0dt00140.prod.ns.nl\logs$"
	
	'' Connect to the class object SplunkLog
	'Set gobjSplunkLog = New SplunkLog_06
	
	'' Setup class with the correct information
	'gobjSplunkLog.SetLogPath(gobjSplunkLog.GetServerShare() & "\[SCRIPTNAME]\[DOMAINNETBIOS]\[COMPUTERNAME]\[YYYY-MM-DD]\[HH-MM-SS].log")
	' gobjSplunkLog.OpenFile()
	
	Set db = New ClassMySQL
	db.DbOpen("DSN_ADBEHEER")
	
	WScript.Echo "Is the database open: " & db.Status
End Sub


Sub ScriptRun()
	Const	COL_FIRST	=	1
	Const	COL_MIDDLE	=	2
	Const	COL_LAST	=	3 
	Const	COl_COMP	=	4
	Const	COL_TITLE	=	5
	Const	COL_MOBILE	=	6
	Const	COL_SAME	=	7
	Const	COL_DOMAIN	=	8
	Const	COL_ARGUS	=	9
	Const	COL_REQUEST_BY = 11

	Dim		intRow
	Dim		strAccountName
	Dim		strFirstName
	Dim		strMiddleName
	Dim		strLastName
	Dim		strSupplierId
	Dim		strDescription
	Dim		strSameAs
	Dim		strTitle			''	 Funtietitel van de beheerder. 
	Dim		strArgusCall
	Dim		strRootDseXls
	Dim		strCurrentDomain
	Dim		strCommandLine
	Dim		strMobile
	Dim		strDomainDns
	Dim		strOuBeheer
	Dim 	blnUseCompanyOu
	Dim		objFileAccount
	Dim		strPassword
	Dim		strPreWin2000
	Dim		strRequestedBy
	

	
	WScript.Echo 
	
	''On Error Resume Next
	''gobjFso.DeleteFile GetScriptPath & "RUNON*.CMD", True
	
	'' intRow = 10		' First row of the information in Excel
	
	' Loop until a empty line is found.
	'' Do Until gobjExcel.Cells(intRow, 1).Value = ""
		''strFirstName = Trim(gobjExcel.Cells(intRow, COL_FIRST).Value)
		''strMiddleName = Trim(gobjExcel.Cells(intRow, COL_MIDDLE).Value)
		''strLastName = Trim(gobjExcel.Cells(intRow, COL_LAST).Value)
		'strCompany = Trim(gobjExcel.Cells(intRow, COl_COMP).Value)
		'strTitle = Trim(gobjExcel.Cells(intRow, COL_TITLE).Value)
		'strSameAs = Trim(gobjExcel.Cells(intRow, COL_SAME).Value)
		'strRootDseXls = Trim(gobjExcel.Cells(intRow, COL_DOMAIN).Value)
		'strArgusCall = Trim(gobjExcel.Cells(intRow, COL_ARGUS).Value)
		' strMobile = Trim(gobjExcel.Cells(intRow, COL_MOBILE).Value)
		'strRequestedBy = Trim(gobjExcel.Cells(intRow, COL_REQUEST_BY).Value)
		'strAccountName = GenerateAccountName(strCompany, strFirstName, strMiddleName, strLastName)
		
		'strDomainDns = ConfigReadSettingSection(strRootDseXls, "Upn")
		'strOuBeheer = ConfigReadSettingSection(strRootDseXls, "BeheerOu")
		'blnUseCompanyOu = CBool(ConfigReadSettingSection(strRootDseXls, "UseCompanyOu"))
		'strPreWin2000 = ConfigReadSettingSection(strRootDseXls, "PreWin2000")
		
		'strDescription = "CALL=" & strArgusCall
		
		'WScript.Echo strAccountName & vbTab & strDomainDns
		
		'Set gobjFile = gobjFso.OpenTextFile("RUNON_" & strDomainDns & ".cmd", FOR_APPENDING, True)
		
		' Set objFileAccount = gobjFso.OpenTextFile(strArgusCall & " " & strPreWin2000 & " " & strAccountName & ".txt", FOR_WRITING, True)
		
		' dsadd user "CN=KPN_Joop.deCock,OU=KPN,OU=Beheer,DC=test,DC=ns,DC=nl" 
		'	-samid KPN_Joop.deCock 
		'	-pwd "JINX@whos28" 
		'	-fn "Joop de"
		'	-ln "Cock" 
		'	-desc "CALL=546464"
		'	-title "Technisch Beheerder"
		'	-display "KPN_Joop.deCock" 
		'	-upn "KPN_Joop.deCock@test.ns.nl" 
		'	-mustchpwd yes
		
		'Call CreateNewAccount(strFirstName, strMiddleName, strLastName, strCompany, strTitle, strMobile, strSameAs, strRootDseXls, strArgusCall, strRequestedBy)
		
		
		'gobjFile.WriteLine
		'gobjFile.WriteLine "echo Creating " & strAccountName
		'strCommandLine = "dsadd.exe user "
		'strCommandLine = strCommandLine & Chr(34) & "CN=" & strAccountName & ","
		'If blnUseCompanyOu = True Then
		'	strCommandLine = strCommandLine & "OU=" & strCompany & "," 
		'End If
		'strCommandLine = strCommandLine & strOuBeheer & "," & strRootDseXls & Chr(34) & " "
		'strCommandLine = strCommandLine & "-samid " & Chr(34) & strAccountName & Chr(34) & " "
		'strCommandLine = strCommandLine & "-pwd " & Chr(34) & strPassword & Chr(34) & " "
		'strCommandLine = strCommandLine & "-fn " & Chr(34) & Trim(strFirstName & " " & strMiddleName) & Chr(34) & " "
		'strCommandLine = strCommandLine & "-ln " & Chr(34) & strLastName & Chr(34) & " "
		'strCommandLine = strCommandLine & "-title " & Chr(34) & strTitle & Chr(34) & " "
		'strCommandLine = strCommandLine & "-desc " & Chr(34) & strDescription & Chr(34) & " "
		'strCommandLine = strCommandLine & "-display " & Chr(34) & strAccountName & Chr(34) & " "
		'strCommandLine = strCommandLine & "-upn " & Chr(34) & strAccountName & "@" & strDomainDns & Chr(34) & " "
		'strCommandLine = strCommandLine & "-mustchpwd yes"
		'WScript.Echo strCommandLine
		'gobjFile.WriteLine strCommandLine
		
		'objFileAccount.WriteLine "Het nieuwe beheeraccount is aangemaakt"
		'objFileAccount.WriteLine
		'objFileAccount.WriteLine "Account:          " & strPreWin2000 & "\" & strAccountName
		'objFileAccount.WriteLine "Initial password: " & strPassword
		'objFileAccount.WriteLine
		
		'gobjSplunkLog.AddKey "domain_netbios", strPreWin2000, vbString
		' gobjSplunkLog.AddKey "executed_by", gobjSplunkLog.GetCurrentUserName(), vbString
		'gobjSplunkLog.AddKey "user_name", strAccountName, vbString
		'gobjSplunkLog.AddKey "full_name", Trim(strFirstName & " " & strMiddleName) & " " & strLastName, vbString
		
		
		'If Len(strSameAs) > 0 Then
		'	objFileAccount.WriteLine "De domain groupen voor het account zijn gelijk gemaakt aan account " & strSameAs
		'	strCommandLine = "cscript.exe //nologo AdMakeUserGroupsSameAs.vbs " & Chr(34) & strSameAs & Chr(34) & " " & Chr(34) & strAccountName & chr(34)
		'	WScript.Echo strCommandLine
		'	gobjFile.WriteLine strCommandLine
			
			'gobjSplunkLog.AddKey "same_as", strSameAs, vbString
		'End If
		'gobjSplunkLog.WriteLog
		
		'objFileAccount.Close
		'Set objFileAccount = Nothing
		
		'gobjFile.Close
		'Set gobjFile = Nothing
		'intRow = intRow + 1
	'Loop
	
	Call RecordPrepare(0)
	Call RecordCreate(100)
	Call RecordInform(200)
End Sub ' ScriptRun


Sub ScriptDone()
	'' close the class SplunkLog
	'Set gobjSplunkLog = Nothing

	'' Close the database object
	db.DbClose
	Set db = Nothing
	
	Set gobjShell = Nothing
	
	Set gobjFso = Nothing

	'Set gobjSheet = Nothing
	'gobjExcel.Quit
	'Set gobjExcel = Nothing
End Sub



Sub ScriptTest()
	Dim		s
	Dim		r
	Dim		x
	
	WScript.Echo
	WScript.Echo "TESTING!!"
	
	'r = "DC=prod,DC=ns,DC=nl"
	's = "Perry.vandenHondel"
	'WScript.Echo "DN for " & s & " = " & DsutilsGetDnFromSam(r, "user", s)
	
	's = "Rudolf.TheRedNosedReendeer"
	'WScript.Echo "DN for " & s & " = " &DsutilsGetDnFromSam(r, "user", s)
	
	' s = "BEH_Perry.vdHondel"
	'WScript.Echo "DN for " & s & " = " & DsutilsGetDnFromSam(r, "user", s)
	
	'WScript.Echo GetDomainValues("DC=prod,DC=ns,DC=nl", FLD_DMN_UPN)
	
	Call SetNotDelegatedFlag("DC=test,DC=ns,DC=nl", "NSA_Pieter.Post")
	
	'For x = 1 To 1500
	'	WScript.Echo x & ":" & vbTab & GeneratePassword()
	'Next
	
	
End Sub ' ScriptTest



Function RunCommand(sCommandLine)
	''
	'//	RunCommand(sCommandLine)
	'//
	'//	Run a DOS command and wait until execution is finished before the cript can commence further.
	'//
	'//	Input
	'//		sCommandLine	Contains the complete command line to execute 
	'//
	Dim oShell
	Dim sCommand
	Dim	nReturn

	Set oShell = WScript.CreateObject("WScript.Shell")
	sCommand = "CMD /c " & sCommandLine
	' 0 = Console hidden, 1 = Console visible, 6 = In tool bar only
	'LogWrite "RunCommand(): " & sCommandLine
	nReturn = oShell.Run(sCommand, 6, True)
	Set oShell = Nothing
	RunCommand = nReturn 
End Function '' RunCommand



Sub CreateNewAccount(ByVal strFirstName, ByVal strMiddleName, ByVal strLastName, ByVal strCompany, ByVal strTitle, ByVal strMobile, ByVal strSameAs, ByVal strRootDseXls, ByVal strArgusCall, ByVal strRequestedBy)
	'
	'	Create a new account
	'
	Dim		strPassword
	Dim		strAccountName
	Dim		strUpn
	Dim		strUpnFileName
	Dim		strDomainDns
	Dim		strSameAsDn
	Dim		strAccountDn
	Dim		blnUseCompanyOu
	Dim		strOuBeheer
	Dim		strDescription
	Dim		c
	Dim		d
	
	
	WScript.Echo 
	WScript.Echo Left("CreateNewAccount()" & String(80, "-"), 80)
	
	strAccountName = GenerateAccountName(strCompany, strFirstName, strMiddleName, strLastName)
	'	Create a lower case account name
	strAccountName = LCase(strAccountName)
	strDomainDns = ConfigReadSettingSection(strRootDseXls, "Upn")
	strUpn = strAccountName & "@" & strDomainDns
	strUpnFileName = strArgusCall & "_" & Replace(strUpn, "@", "-at-") & ".txt" 
	
	Set tfNewAccount = New ClassTextFile
	tfNewAccount.SetMode(FOR_WRITING)
	tfNewAccount.SetPath(strUpnFileName)
	tfNewAccount.OpenFile()

	WScript.Echo "Write to text file: " & strUpnFileName
	
	If Len(DsutilsGetDnFromSam(strRootDseXls, "user", strAccountName)) > 0 Then
		'	Generated account exists in the target AD domain. Skip this entry
		WScript.Echo "WARNING: Account " & strAccountName & " already exists in " & strRootDseXls
		tfNewAccount.WriteLineToFile("WARNING: Account " & strAccountName & " already exists in AD domain " & strRootDseXls)
		Exit Sub
	End If
	
	strPassword = GeneratePassword() 
	
	tfNewAccount.WriteLineToFile("CALL")
	tfNewAccount.WriteLineToFile(vbTab & "Call reference:     " & strArgusCall)
	tfNewAccount.WriteLineToFile("")
	tfNewAccount.WriteLineToFile("ACCOUNT")
	tfNewAccount.WriteLineToFile(vbTab & "Account:            " & strUpn)
	tfNewAccount.WriteLineToFile(vbTab & "Initial password:   " & strPassword)
	tfNewAccount.WriteLineToFile("")
	
	If Len(strSameAs) > 0 Then
		tfNewAccount.WriteLineToFile("REFERENCE ACCOUNT")
		tfNewAccount.WriteLineToFile(vbTab & "Clone groups from: " & strSameAs)
		
		strSameAsDn = DsutilsGetDnFromSam(strRootDseXls, "user", strSameAs)
		If Len(strSameAsDn) = 0 Then
			tfNewAccount.WriteLineToFile(vbTab & "WARNING: Could not find the reference account of " & strSameAs & " to perform a group clone action")
		End If
	End If
	
	'	Build the DN for the new account to create
	blnUseCompanyOu = CBool(ConfigReadSettingSection(strRootDseXls, "UseCompanyOu"))
	strOuBeheer = ConfigReadSettingSection(strRootDseXls, "BeheerOu")
	strAccountDn = "CN=" & strAccountName & ","
	If blnUseCompanyOu = True Then
		strAccountDn = strAccountDn & "OU=" & strCompany & "," 
	End If
	
	strAccountDn = strAccountDn & strOuBeheer & "," & strRootDseXls
	strDescription = "CALL=" & strArgusCall & " REQUEST_BY=" & strRequestedBy
	
	If Left(strMobile, 1) <> "+" Then
		tfNewAccount.WriteLineToFile(vbTab & "WARNING: No valid mobile number is specified for this new administrative account using SmartXS, please supply a valid valid mobile number")
		strMobile = "+31600000000"
	End If
	
	WScript.Echo vbTab & "First name                   : " & strFirstName
	WScript.Echo vbTab & "Middle name                  : " & strMiddleName
	WScript.Echo vbTab & "Last name                    : " & strLastName
	WScript.Echo vbTab & "Company                      : " & strCompany
	WScript.Echo vbTab & "Title                        : " & strTitle
	WScript.Echo vbTab & "Mobile                       : " & strMobile
	WScript.Echo vbTab & "Reference account            : " & strSameAs
	WScript.Echo vbTab & "Root DSE                     : " & strRootDseXls
	WScript.Echo vbTab & "Reference call id            : " & strArgusCall
	WScript.Echo vbTab & "----"
	WScript.Echo vbTab & "Account name                 : " & strAccountName
	WScript.Echo vbTab & "UPN                          : " & strUpn
	WScript.Echo vbTab & "Initial password             : " & strPassword
	WScript.Echo vbTab & "New account DN               : " & strAccountDn
	WScript.Echo vbTab & "Reference account DN         : " & strSameAsDn
	WScript.Echo vbTab & "Description                  : " & strDescription
	
	'	Build the command line to create a new account using DSADD.EXE
	c = "dsadd user " & Chr(34) & strAccountDn & Chr(34) & " "
	c = c & "-samid " & Chr(34) & strAccountName & Chr(34) & " "
	c = c & "-pwd " & Chr(34) & strPassword & Chr(34) & " "
	c = c & "-fn " & Chr(34) & Trim(strFirstName & " " & strMiddleName) & Chr(34) & " "
	c = c & "-ln " & Chr(34) & strLastName & Chr(34) & " "
	c = c & "-title " & Chr(34) & strTitle & Chr(34) & " "
	c = c & "-desc " & Chr(34) & strDescription & Chr(34) & " "
	c = c & "-display " & Chr(34) & strAccountName & Chr(34) & " "
	c = c & "-upn " & Chr(34) & strUpn & Chr(34) & " "
	c = c & "-mobile " & Chr(34) & strMobile & Chr(34) & " "
	c = c & "-company " & Chr(34) & strCompany & Chr(34) & " "
	c = c & "-mustchpwd yes"
	
	WScript.Echo vbTab & c
	gobjShell.Run "cmd /c " & c, 0, True
	
	WScript.Echo vbTab & "ADD TO DEFAULT GROUPS FOR SmartXS (Works only under PROD.NS.NL)"
	
	'' Add default group ROL_SmartXS_Autorisatie
	d = DsutilsGetDnFromSam(strRootDseXls, "group", "RP_SmartXS_Autorisatie") ' Was ROL_SmartXS_Autorisatie
	Call DsmodGroupAddMember(d, strAccountDn)

	'' Add default group for BONS Autorisatie
	d = DsutilsGetDnFromSam(strRootDseXls, "group", "RP_BONS_Autorisatie") ' 2014-12-29 PVDH Kan de groep niet vinden in PROD.NS.NL
	Call DsmodGroupAddMember(d, strAccountDn)

	'' Add to default beheer group, 2015-02-18 PVDH
	WScript.Echo vbTab & "ADD ACCOUNT TO DEFAULT ACCOUNT TYPE GROUP (AT-*)"
	d = DsutilsGetDnFromSam(strRootDseXls, "group", "AT-PERSOON-BEHEER-" & strCompany)
	Call DsmodGroupAddMember(d, strAccountDn)

	WScript.Echo vbTab & "SETTING THE ACCOUNT UAC FLAG NOT_DELEGATED (Wait 10 seconds...)"
	WScript.Sleep 10000
	'adfind.exe -b "DC=prod,DC=ns,DC=nl" -f "sAMAccountName=BEH_WMIScanProject" userAccountControl -adcsv | admod userAccountControl::{{.:SET:1048576}}	
	c = "adfind.exe -b " & Chr(34) & strRootDseXls & Chr(34) & " -f " & Chr(34) & "sAMAccountName=" & strAccountName & Chr(34) & " userAccountControl -adcsv | admod.exe userAccountControl::{{.:SET:1048576}}"
	'WScript.Echo c
	Call RunCommand(c)
	
	'	Clone the groups of the reference accounts
	If Len(strSameAsDn) > 0 Then
		WScript.Echo vbTab & "ADD TO REFERENCE ACCOUNT GROUPS"	
		WScript.Echo vbTab & "INFO: Reference account specified, clone groups from " & strSameAsDn
		Call MakeSameAs(strSameAsDn, strAccountDn)
	Else
		WScript.Echo vbTab & "INFO: No reference account specified, no group cloning done."
	End If
	
	tfNewAccount.CloseFile()
	Set tfNewAccount = Nothing
End Sub ' CreateNewAccount



Sub MakeSameAs(ByVal s, ByVal d)
	'
	'	Copy all groups from s to d
	'	
	'	s		Source DN
	'	d		Destionation DN
	'
		
	Dim		c
	Dim		strPathTemp
	Dim		f
	Dim		ts
	Dim		l
	Dim		i
	
	WScript.Echo "MakeSameAs()"

	'	Get a temp file name to store the group DN's
	strPathTemp = GetTempFileName()
	
	i = 0
	
	c = "dsget.exe user " & Chr(34) & s & Chr(34) & " -memberof >" & strPathTemp
	WScript.Echo vbTab & c
	gobjShell.Run "cmd /c " & c, 0, True
	
	On Error Resume Next
	Set f = gobjFso.GetFile(strPathTemp)
	If Err.Number = 0 Then
		Set ts = f.OpenAsTextStream(FOR_READING)
		Do While ts.AtEndOfStream <> True
			l = ts.ReadLine
			If InStr(l, "CN=") > 0 Then
				c = "dsmod.exe group " & l & " -addmbr " & d
				WScript.Echo c
				i = i + 1
				gobjShell.Run "cmd /c " & c, 0, True
			End If
		Loop
		f.Close
		f.Delete
		Set f = Nothing
		
		tfNewAccount.WriteLineToFile(vbTab & "Account added to " & i & " groups")
		
	Else
		WScript.Echo "WARNING: Could not open " & strPathTemp & " (" & Err.Number & ")"
	End If 
End Sub ' MakeSameAs



Function GetTempFileName
	'//////////////////////////////////////////////////////////////////////////////
	'//
	'//	GetTempFileName
	'//
	'//	Input:
	'//		None
	'//
	'//	Output:
	'//		A string with a temporary file name, e.g. E:\temp\filename.tmp
	'//
	Dim	oTempFolder
	Dim	sTempFile
	Dim	oFso

	Const	WINDOWS_FOLDER		=	0
	Const	SYSTEM_FOLDER		=	1
	Const	TEMPORARY_FOLDER	=	2

	Set oFso = CreateObject("Scripting.FileSystemObject")
	Set oTempFolder = oFSO.GetSpecialFolder(TEMPORARY_FOLDER)
	sTempFile = oFSO.GetTempName

	GetTempFileName = oTempFolder & "\" & sTempFile

	Set oTempFolder = Nothing
	Set oFso = Nothing
End Function '' GetTempFileName



Sub DsmodGroupAddMember(ByVal g, ByVal m)
	'
	'	dsmod group "CN=US INFO,OU=Distribution Lists,DC=microsoft,DC=com" -addmbr "CN=John Smith,CN=Users,DC=microsoft,DC=com"
	'
	Dim		c
	
	c = "dsmod.exe group " & Chr(34) & g & Chr(34) & " -addmbr " & Chr(34) & m & Chr(34)
	'WScript.Echo "DsmodGroupAddMember(): " & c
	gobjShell.Run "cmd /c " & c, 0, True
End Sub ' DsmodGroupAddMember



Sub DeleteFile(sPath)
	''
	''	DeleteFile()
	''	
	''	Delete a file specified as "d:\folder\filename.ext"
	''
	''	sPath	The name of the file to delete.
	''
   	Dim oFSO
   	
   	Set oFSO = CreateObject("Scripting.FileSystemObject")
   	If oFSO.FileExists(sPath) Then
   		oFSO.DeleteFile sPath, True
   	End If
   	Set oFSO = Nothing
End Sub '' DeleteFile



Function GenerateAccountName(ByVal strCompany, ByVal strFirstName, ByVal strMiddleName, ByVal strLastName)
	'
	'	Generate a new SAMAccountName for a user.
	'
	'		strFirstName			Richard
	'		strMiddleName			van
	'		strLastName				Beukenstein
	'		strCompany				HP
	'
	Dim	strAccountName
	Dim	intLen
	
	'WScript.Echo 
	'WScript.Echo "GenerateAccountName: " & strCompany & vbTab & "FN:" & strFirstName & vbTab & "MN:" & strMiddleName & vbTab & "LN:" & strLastName

	strFirstName = FixFirstName(strFirstName)
	
	' Remove all middle names such as: van den, van, de, 't , van 't
	strMiddleName = FixMiddleName(strMiddleName)
	
	strLastName = FixLastName(strLastName)
	
	' Build the Account name.
	' V06: Select between company if blank!
	If Len(strCompany) = 0 Then
		strAccountName = strFirstName & "." & strMiddleName & strLastName
	Else
		strAccountName = strCompany & "_" & strFirstName & "." & strMiddleName & strLastName
	End If
		
	'WScript.Echo "GenerateAccountName(): " & strAccountName
	
	'WScript.Echo strAccountName
	' Get the length of the generated account name.
	intLen = Len(strAccountName)
	'WScript.Echo "Len(" & strAccountName & ")=" & intLen
	If intLen > MAX_ACCOUNT_LEN Then
		' The generated username is to long, shorten it by using only
		' the initial of the user's first name.
		
		strAccountName = strCompany & "_" & Left(strFirstName, 1) & "." & strMiddleName & strLastName
		
	End If
	' Only take the MAX_ACCOUNT_LEN of chars of the newly generated account.
	
	' Strip all spaces from the account name.
	strAccountName = Replace(strAccountName, " ", "")
	
	' Return the generated account name, but only the first 20 chars.
	GenerateAccountName = Left(strAccountName, MAX_ACCOUNT_LEN)
End Function ' GenerateAccountName



Function FixMiddleName(strMiddleName)
	Dim	arrPrefix
	Dim	strPrefix
	Dim	x
	
	WScript.Echo "FixMiddleName(" & strMiddleName & ")"
	If Len(strMiddleName) > 0 Then
	
		' Convert the string in lower case.
		strMiddleName = LCase(strMiddleName)
	
		' Remove all quote's (') in the middle name.
		strMiddleName = Replace(strMiddleName, "'", "")
	
		' Vul de string met velden van de array.
		strPrefix = "van den|vd"
		strPrefix = strPrefix & "|van der|vd"
		strPrefix = strPrefix & "|van den|vd"
		strPrefix = strPrefix & "|van de|vd"
		strPrefix = strPrefix & "|van t|vt"
		strPrefix = strPrefix & "|ten|t"
		strPrefix = strPrefix & "|van|v"
		strPrefix = strPrefix & "|den|d" 
		strPrefix = strPrefix & "|de|d" 	
		strPrefix = strPrefix & "|in t|it"
		strPrefix = strPrefix & "|t|t"
		strPrefix = strPrefix & "|la|l"
		strPrefix = strPrefix & "|le|l"
	
		' Convert the Prefix string to an array.
		arrPrefix = Split(strPrefix, "|")
	
		For x = 0 To UBound(arrPrefix) Step 2
			'WScript.Echo x & vbTab & arrPrefix(x) & vbTab & arrPrefix(x + 1)
			If InStr(1, strMiddleName, arrPrefix(x), vbTextCompare) > 0 Then
				strMiddleName = Replace(strMiddleName, arrPrefix(x), arrPrefix(x + 1))
				Exit For
			End If
		Next

		'WScript.Echo "FixMiddleName(): " & strMiddleName
		FixMiddleName = strMiddleName
	Else	
		FixMiddleName = ""
	End If
End Function ' FixMiddleName



Function FixLastName(strName)
	' Remove all quote's (') in the last name.
	strName = Replace(strName, " ", "")
	
	' Remove all '-' in the last name.
	strName = Replace(strName, "-", "")
	
	FixLastname = strName
End Function ' FixLastname()



Function FixFirstName(strName)
	' Remove all quote's (') in the last name.
	strName = Replace(strName, " ", "")
	
	' Remove all '-' in the last name.
	strName = Replace(strName, "-", "")
	
	FixFirstname = strName
End Function ' FixLastname()



Function GeneratePassword()
	''
	''	Generate a new password based on Scrabble 4-letter words.
	''
	''		
	''	Returns: 
	''		JAME-axil-05 (example) (uppercase-lowercase-day)
	''		123456789012
	''
	Dim	strReturn
	Dim	strWord
	Dim	arrWord
	Dim	x
	Dim	intMax
	Dim	intMin
	Dim	strWord1
	Dim	strWord2
	Dim		r
	
	' Set the number randomizer
	Randomize
	
	' Use four letter Scrabble words for password parts
	strWord = "APEX;AXAL;AXED;AXEL;AXES;AXIL;AXIS;AXLE;AXON;CALX;COAX;COXA;CRUX;DEXY;DOUX;DOXY;EAUX;EXAM;EXEC;" & _
		"EXES;EXIT;EXON;EXPO;FALX;FAUX;FIXT;FLAX;FLEX;FLUX;FOXY;HOAX;IBEX;ILEX;IXIA;JEUX;JINX;LUXE;LYNX;MAXI;MINX;" & _
		"MIXT;MOXA;NEXT;NIXE;NIXY;OYNX;ORYX;OXEN;OXES;OXID;OXIM;PIXY;PREX;ROUX;SEXT;SEXY;TAXA;TAXI;TEXT;VEXT;WAXY;" & _
		"XYST;AJAR;AJEE;DJIN;DOJO;FUJI;HADJ;HAJI;HAJJ;JABS;JACK;JADE;JAGG;JAGS;JAIL;JAKE;JAMB;JAMS;JANE;JAPE;JARL;" & _
		"JARS;JATO;JAUK;JAUP;JAVA;JAWS;JAYS;JAZZ;JEAN;JEED;JEEP;JEER;JEES;JEEZ;JEFE;JEHU;JELL;JEON;JERK;JESS;JEST;" & _
		"JETE;JETS;JEUX;JIAO;JIBB;JIBE;JIBS;JIFF;JIGS;JILL;JILT;JIMP;JINK;JINN;JINS;JINX;JISM;JIVE;JOBS;JOCK;" & _
		"JOES;JOEY;JOGS;JOHN;JOIN;JOKE;JOKY;JOLE;JOLT;JOSH;JOSS;JOTA;JOTS;JOUK;JOWL;JOWS;JOYS;JUBA;JUBE;JUDO;JUGA;" & _
		"JUGS;JUJU;JUKE;JUMP;JUNK;JUPE;JURA;JURY;JUST;JUTE;JUTS;MOJO;PUJA;RAJA;SOJA;RAIL;NAIL;ZARA;ZAZA;YULO"
		
	' Fill the array with the words.
	arrWord = Split(strWord, ";")

	' Set the low and high bounderies
	intMin = 0
	intMax = UBound(arrWord)
	
	' Retrieve a word from the array based on the randomized posisition for Word 1 and Word 2.
	strWord1 = arrWord(Int((intMax - intMin + 1) * Rnd + intMin))
	strWord2 = arrWord(Int((intMax - intMin + 1) * Rnd + intMin))
	
	' Return the generated password. Format XXXX@xxxx00
	r = UCase(strWord1) & "-" & LCase(strWord2) & "-" & Right("0" & Day(Date), 2)
	If Len(r) <> 12 Then
		'' When the password is not 12 chars long, generate a new one.
		r = GeneratePassword()
	End If
	GeneratePassword = r
End Function  '' of Function GeneratePassword



Function DsutilsGetDnFromSam(ByVal strRoot, ByVal strType, ByVal strSamAccount)
	'
	'	Get the DN from an sAMAccountName using DSQUERY.EXE
	'
	'	strRoot			"DC=test,DC=ns,DC=nl"
	'	strType			user, group
	'	strSamAccount	"Perry.vandenHondel"
	'
	'	Returns
	'		DN				Found the DN path of the strSamAccount
	'		Empty string	Not found
	'
	
	Dim		r			'	Result
	Dim		c			'	Command Line
	Dim		f			'	File Object
	Dim		ts			' 	TextStream
	Dim		l			' 	Line
	Dim		i
	Dim		x
	Dim		objShell
	Dim		objExec
	Dim		strPath
	
	Set objShell = CreateObject("WScript.Shell")

	c = "dsquery.exe " & strType & " " & strRoot & " -samid " & strSamAccount 
	'WScript.Echo c
	Set objExec = objShell.Exec(c)
	r = Replace(objExec.StdOut.ReadLine, """", "") ' And Remove " around the string
	'WScript.Echo "r=[" & r & "]"
	Set objShell = Nothing
	
	DsutilsGetDnFromSam = r
End Function ' DsutilsGetDnFromSam



Function GetScriptPath()
	'==
	'==	Returns the path where the script is located.
	'==
	'==	Output:
	'==		A string with the path where the script is run from.
	'==
	'==		drive:\folder\folder\
	'==
	Dim sScriptPath
	Dim sScriptName

	sScriptPath = WScript.ScriptFullName
	sScriptName = WScript.ScriptName

	GetScriptPath = Left(sScriptPath, Len(sScriptPath) - Len(sScriptName))
End Function '' GetScriptPath



Function AdGetRootDse()
    '==
    '==     Returns the RootDSE of the current domain.
    '==
    '==     Returns:
    '==          DC=gwnet,DC=nl
    '==
    Dim     objRootDse
    
	Set objRootDse = GetObject("LDAP://RootDse")
    AdGetRootDse = objRootDse.get("defaultNamingContext")
    Set objRootDse = Nothing
End Function '== AdGetRootDse



Function ConfigReadSettingSection(sSection, sSetting)
    ''
    ''     Verbeterde versie 2009-04-02
    ''
    ''     Reads a setting from a .conf file and returns the value.
    ''
    ''     Name the file.csv the same as the script but with a csv extension.
    ''    
    ''     ----- FILE_NAME_AS_SCRIPT.VBS ---
    ''
    ''     Function_ConfigReadSettingSection.vbs -- Config file.
    ''
    ''    
    ''     [Section1]
    ''     Name=Whatever is the biatch
    ''     Name=Perry
    ''    
    ''     [Section2]
    ''     Name=Adrian
    ''
    ''     [Section3]
    ''     Name=Jill
    ''     ------------------
    ''
    ''     Example looping for more entries:
    ''     Dim     x
    ''     Dim     bAgain
    ''     Dim     sLogEntry
    ''     bAgain = True
    ''     x = 1
    ''     Do
    ''          sLogEntry = ConfigReadSetting("LogEntry" & x)
    ''    
    ''          If IsEmpty(sLogEntry) Then
    ''               bAgain = False
    ''          Else
    ''          WScript.Echo x & ": [" & sLogEntry & "]"
    ''          End If
    ''          x = x + 1
    ''     Loop Until bAgain = False
    ''
    ''     Remark: Convert strings to Integers for numbers
    ''          n = Int(ConfigReadSetting("", "Number"))
	'
	'	Uses:
	'		GetScriptNameMinVersion()
	'
    '
    Const     FOR_READING = 1
    Const     SEPERATOR = "="
    
    Dim     sLine
    Dim     oFso
    Dim     oFile
    Dim     sPath
    Dim     sReturn
    Dim     bInSection
    Dim     sSectionSelect
    
	'WScript.Echo "sSection="&sSection
	'WScript.Echo "sSetting="&sSetting
	
    sReturn = ""
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
     
	
	' WScript.Echo "ScriptPath="& WScript.ScriptFullName
	 
	' Build the path for the configuration file. x:\folder\script.vbs >> x:\folder\script.config
	sPath = Replace(WScript.ScriptFullName, ".vbs", ".config")
	Set oFile = oFso.OpenTextFile(sPath, FOR_READING)
    
     sSectionSelect = "[" & sSection & "]"
    
    'WScript.Echo "ConfigReadSetting(): sSection="&sSection
    
    If sSectionSelect = "[]" Then
        'WScript.Echo "No section specified"
         
         Do While oFile.AtEndOfStream <> True
            sLine = oFile.ReadLine
              
            '' If in the line is a seperator char or does not start with a quote, it's a config line.
            If (InStr(sLine, SEPERATOR) > 0) And (Left(sLine, 1) <> "'") Then
                '' Check if the value of sSetting is in the line.
                If InStr(sLine, sSetting) > 0 Then
                    sReturn = Right(sLine, Len(sLine) - InStr(sLine, SEPERATOR))
                     Exit Do
                End If
            End If
        Loop
    Else
		'WScript.Echo "Section " & sSectionSelect & " specified"
                   
        bInSection = False
         
        Do While oFile.AtEndOfStream <> True
            sLine = oFile.ReadLine
         
			If InStr(sLine, sSectionSelect) > 0 Then
                'WScript.Echo "Found Section " & sSectionSelect
                bInSection = True
            End If
         
            If bInSection = True Then
                '' If in the line is a seperator char or does not start with a quote, it's a config line.
                If (InStr(sLine, SEPERATOR) > 0) And (Left(sLine, 1) <> "'") Then
                    '' Check if the value of sSetting is in the line.
                    If InStr(sLine, sSetting) > 0 Then
                        sReturn = Right(sLine, Len(sLine) - InStr(sLine, SEPERATOR))
                        Exit Do
                    End If
                End If
            End If
        Loop
    End If
    
    oFile.Close
    Set oFile = Nothing
    
    Set oFso = Nothing

    ConfigReadSettingSection = sReturn
End Function '' ConfigReadSettingSection



Function GetScriptNameMinVersion
    '
    '	GetScriptName
	'	
	'	Return the name of the current script with or without a version number.
	'	
	'	Variables:
	'		blnMinusVersion		True returns the script name minus the version: ScriptName-02 > ScriptName
	'
	'	Returns:
    '		GetScriptName(True)		ScriptName
	'		GetScriptName(False)	ScriptName-03
	'
    Dim	strScriptName
	Dim	strReturn

	strScriptName = WScript.ScriptName						'' Get the current script name
	strScriptName = Replace(strScriptName, ".vbs", "")		'' Remove the file extension
	
	If InStr(strScriptName, "-") > 0 Then
		strReturn = left(strScriptName, InStrRev(strScriptName, "-") - 1)
	Else
		strReturn = Left(strScriptName, InStrRev(strScriptName, ".") - 1)
	End If
	
     GetScriptNameMinVersion = strReturn
End Function '' GetScriptName







Class ClassTextFile
	'
	'	General class to handle file operations for text files .
	'
	'	Parent class for
	'		ClassTextFileTsv
	'		ClassTextFileSplunk
	'
	'	Class Subs and Functions:
	'		Private Sub Class_Initialize				Class initializer sub, set all default values
	'		Private Sub Class_Terminate					Class terminator, releases all variables, etc..
	'		Public Sub SetPath(ByVal strPathNew)		Sets the path to the file.
	'		Public Function GetPath						Returns the path of the file C:\folder\file.ext
	'		Public Sub SetMode							Set the mode of access for the file (READ, WRITE, APPEND)
	'		Public Sub OpenFile							Open the file
	'		Public Sub CloseFile						Closes the file
	'		Public Sub WriteToFile(ByVal strLine)		Write a line to the file
	'		Public Function ReadFromFile()				Read a line from the  file
	'		Public Sub DeleteFile						Delete the file
	'		Function IsEndOfFile						Boolean returns the end of the file reached
	'		Public Function CurrentLine()				Returns the current line number
	'

	Private		objFso	
	Private		objFile
	Private		strPath
	Private		blnIsOpen
	Private		intMode				'	Modus of file activity, READING=1, WRITING=2, APPENDING=8
	Private		intLineCount
	
	Private Sub Class_Initialize
		'
		'	Class initializer, open objects, set default variable values.
		'
		Set objFso = CreateObject("Scripting.FileSystemObject")
		blnIsOpen = False
		intLineCount = 0
	End Sub '' Class_Initialize

	Private Sub Class_Terminate
		'
		'	Class terminator, closes objects
		'
		Call CloseFile()
		
		'	Terminate the object to the text file
		Set objFile = Nothing
		
		'	Terminate the object to the File System Object
		Set objFso = Nothing
	End Sub '' Class Terminate
	
	Public Sub SetPath(ByVal strPathNew)
		'
		'	Sets the path to the file.
		'	Assumes all folders exist before calling this function
		'
		'If objFso.FileExists(strPathNew) = False Then
		'	WScript.Echo "ClassTextFile.SetPath() ERROR: Path is not found!"
		' End If
		strPath = strPathNew
	End Sub
	
	Public Function GetPath
		'
		'	Returns the current path. c:\folder\file.ext
		'
		GetPath = strPath
	End Function
	
	Public Sub SetMode(intModeNew)
		'
		'	Set the mode of file opening
		'		1:	READ
		'		2:	WRITE
		'		8:	APPEND
		'
		intMode = intModeNew
	End Sub
	
	Public Function GetMode()
		'
		'	Return the mode of file opening.
		'		1:	READ
		'		2:	WRITE
		'		8:	APPEND
		'		
		GetMode = intMode
	End Function
	
	Public Sub OpenFile
		'
		'	Open the file strPath
		'
		On Error Resume Next
		Set objFile = objFso.OpenTextFile(strPath, intMode, True)
		If Err.Number = 0 Then
			blnIsOpen = True
		Else
			WScript.Echo("ClassTextFile/OpenFile ERROR: Could not open textfile: " & strPath)
		End If
	End Sub
	
	Public Sub CloseFile()
		'	
		'	Closes the current opened file
		'	
		If blnIsOpen = False Then
			objFile.Close
		End If
	End Sub
	
	Public Sub WriteLineToFile(ByVal strLine)
		'
		'	Write the contents of strLine to the text file.
		'
		If blnIsOpen = True Then
			objFile.WriteLine(strLine)
		Else
			WScript.Echo "ClassTextFile/WriteLineToFile WARNING: Tried to write to a closed file: " & strPath
		End If
	End Sub
	
	Public Function ReadLineFromFile()
		'
		'	Read a line from the text file.
		'	
		'	Returns a string
		'	Returns a empty string when nothing could be read.
		'
		If blnIsOpen = True Then
			'	Increase the line counter +1
			intLineCount = intLineCount + 1
			
			'	Read a line from the text file.
			ReadLineFromFile = objFile.ReadLine
		Else
			ReadLineFromFile = ""
			WScript.Echo "ClassTextFile/ReadLineFromFile WARNING: Tried to read from a closed file: " & strPath
		End If
	End Function
	
	Public Function CurrentLine()
		'
		'	Return the current line number.
		'
		CurrentLine = intLineCount
	End Function
	
	Public Sub DeleteFile()
		'
		'	Delete the file
		'	Close it if it is open.
		'
		If blnIsOpen = True Then
			'	Close the file if it's open
			Call CloseFile()
			If objFso.FileExists(strPath) Then
				'	Delete the file, always!
				Call objFso.DeleteFile(strPath, True)
			End If
		End If
   	End Sub
	
	Function IsEndOfFile()
		'
		'	Return the AtEndOfStreamStatus
		'	
		'	True	End of stream reached
		'	False	No reached yet
		'
		IsEndOfFile = objFile.AtEndOfStream
	End Function

End Class	'	ClassTextFile





Class ClassMySQL
	'
	'	Versie 02
	'
	'	Class_Initialize()		Private Sub			Class initializer
	'	Class_Terminate()		Private Sub			Class terminator
	'	DbOpen					Public Sub			Open the database en setup database object
	'	DbOpenDsn				Public Sub			Open the database using a Data Source Name (DSN)
	'	DbClose					Public Sub			Close the database en close all open handles, kill database object
	'	FixStr(strText)			Public Function		Fix a string to a format to be used in a query
	'	FixBool(blnValue)		Public Function 	Fix a VBScript True or False to INTEGER 0 or 1 to be used in MySQL
	'	FixDtm(strText)			Public Function		Fix a date time string to a proper format e.g. YYY-MM-DD HH:MM:SS
	'	ExecQuery(strQuery)		Public Function		Run a query
	'	UniqueRecordId			Public Function		Generate a new unique record id in the format: 13306761B8CC03AF (16 chars long)
	'	NumberAlign(intNumber, intLen)
	'							Private Function 	Align a number on the length (NumberAlign(2, 3) > 2 becomes 002)
	'	GetRecCount				Public Function		Return the number of records in a table.
	'	ClearTable(strTable)	Public Sub			Delete all rows from a table.
	'

	
	Private strDatabase
	Private strServer
	Private strUser
	Private strPassword
	Private objDb
	Private blnIsOpen
	Private strUniqueRecId

	
	Private Sub Class_Initialize()
		strDatabase = ""
		strServer = ""
		strUser = ""
		strPassword = ""
		blnIsOpen = False
	End Sub

	
	Private Sub Class_Terminate()
		blnIsOpen = False
	End Sub
	
	
	Public Sub DbOpen(ByVal strDataSourceName)
		Dim	strConnector
		
		''strDatabase = strDatabaseNew
		''strServer = strServerNew
		''strUser = strUserNew
		''strPassword = strPasswordNew
		''strUniqueRecId = UniqueRecordId()
		
		''strConnector = "|={MySQL ODBC 5.1 Driver};" & _
		''"Server=" & strServer & ";" & _
		''"Database=" & strDatabase & ";" & _
		''"User=" & strUser & ";" &_
	 	''"Password=" & strPassword & ";" & _
		''"Option=3;"
		
		Set objDb = CreateObject("ADODB.Connection")
		On Error Resume Next
		objDb.Open "DSN=" & strDataSourceName
		If Err.Number <> 0 Then
			WScript.Echo "ERROR opening DSN " & strDataSourceName & " (" & Err.Description & ")"
			WScript.Quit(Err.Number)
			''Call DbClose()
		End If
		
		blnIsOpen = True
	End Sub
	
	
	Public Sub DbOpenDsn(ByVal strDsn)
		Set objDb = CreateObject("ADODB.Connection")
		On Error Resume Next
		objDb.Open strDsn
		If Err.Number <> 0 Then
			WScript.Echo "ERROR opening " & strDsn & " (" & Err.Description & ")"
			Call DbClose()
		End If
		
		blnIsOpen = True
	End Sub
	
	
	Public Sub DbClose
		If blnIsOpen = True Then
			objDb.Close
			Set objDb = Nothing
		End If
	End Sub
	
	
	Public Function Status()
		Status = blnIsOpen
		'WScript.Echo "Current status of the database " & UCase(strDatabase) & " on server " & strServer & ": " & blnIsOpen
	End Function

	
	Public Function FixInt(ByVal intNumber)
		If IsNull(intNumber) Then
			FixInt = 0
		Else
			FixInt = intNumber
		End If
	End Function
	
	
	Public Function FixStr(ByVal strText)
		'
		'	Returns a string suitable for a SQL query, Attach quote's around it and replace
		'	single quote (') by double quote's ('')
		'
		'	Variables:
		'		strText
		'	
		'	Returns:
		'		A string suitable for using in a SQL query
		' 
		Dim	strBuffer
	
		If IsNull(strText) Or Len(strText) = 0 Then
			strBuffer = "Null"
		Else
			strBuffer = strText
			strBuffer = Replace(strBuffer, "'", "''")
			strBuffer = Replace(strBuffer, "\", "\\")
			strBuffer = "'" & strBuffer & "'"
		End If
		FixStr = strBuffer
	End Function ' FixStr
	

	Public Function FixBool(ByVal blnValue)
		'
		'	Fix the boolean value (true or False) from VBScript (language specific) to its SQL equavaliant (0 or 1)
		'	MySQL doesn't have a boolean field. Use INTEGER instead. This function converts the data
		
		Dim	 blnReturn
		If blnValue = True Then
			intReturn = 1
		Else
			intReturn  = 0
		End If
		FixBool = intReturn
	End Function
	
	
	Public Function FixDtm(ByVal strText)
		'
		'	Fix a date string to the correct layout.
		'	E.g. 2012-12-12 08:34:45
		'
		Dim	strBuffer
	
		If IsNull(strText) Or Len(strText) = 0 Then
			strBuffer = "Null"
		Else
			strBuffer = "'" & ProperDateTime(strText) & "'" 
		End If
		FixDtm = strBuffer
	End Function ' FixDtm


	Private Function ProperDateTime(ByVal dDateTime)
		'
		'	Convert a system formatted date time to a proper format
		'	Returns the current date time in proper format when no date time
		'	is specified.
		'
		'	15-5-2009 4:51:57  ==>  2009-05-15 04:51:57
		'
		If Len(dDateTime) = 0 Then
			dDateTime = Now()
		End If
	
		ProperDateTime = NumberAlign(Year(dDateTime), 4) & "-" & _
			NumberAlign(Month(dDateTime), 2) & "-" & _
			NumberAlign(Day(dDateTime), 2) & " " & _
			NumberAlign(Hour(dDateTime), 2) & ":" & _
			NumberAlign(Minute(dDateTime), 2) & ":" & _
			NumberAlign(Second(dDateTime), 2)
	End Function ' ProperDateTime

	
	Private Function NumberAlign(ByVal intNumber, ByVal intLen)
		'	
		'	Returns a number aligned with zeros to a defined length
		'
		'	NumberAlign(1234, 6) returns '001234'
		'
		NumberAlign = Right(String(intLen, "0") & intNumber, intLen)
	End Function '  NumberAlign
	
	
	Public Sub GetRecordSet(ByRef objRs, ByVal strQuery)
		''	Usage:
		''	
		''	Dim	strQuery
		''	Dim	objRs
		'' 
		'' 	strQuery = "SELECT " & FLD_TMP_SID_ID & " "
		''	strQuery = strQuery & "FROM " & TBL_TMP & " "
		''	strQuery = strQuery & "WHERE " & FLD_TMP_DOMAIN_ID & "=" & intDomainId & ";"
		''
		''	Call objMySql.GetRecordSet(objRs, strQuery)
		''	If objRs.Eof = True Then
		''		WScript.Echo "NO RECORDS FOUND " & strQuery
		''	Else
		''		WScript.Echo "RECORDS FOUND " & strQuery
		''		objRs.MoveFirst
		''		While Not objRs.EOF
		''			WScript.Echo objRs(FLD_TMP_SID_ID).Value
		''			objRs.MoveNext                                        'next foudn object 
		''	Wend
		''	Set objRs = Nothing
		''
		On Error Resume Next
		Set objRs = objDb.Execute(strQuery)
		If Err.Number <> 0 Then
			WScript.Echo "GetRecordSet: ERROR (0x" & Hex(Err.Number) & ") " & Err.Description
			WScript.Echo vbTab & strQuery
			Exit Sub
		End If
	End Sub
	
	
	Public Function GetNextId(ByVal sTable, ByVal sField)
		'
		'	Return the maximum id of a field 
		'
		'	Example table values:
		'
		'		sng_song_id
		'		-----------
		'		         32
		'                33
		'                34
		'
		'	Command:
		'		SqlGetMaxId("song_sng", "sng_song_id")
		'
		'	Returns:
		'				35
		'
		'	Variables:
		'		sTable		Table to check
		'		sField		Id field to check of maximum number
		Dim	sQuery
		Dim	oRs
		Dim	nReturn
	
		sQuery = "SELECT MAX(" & sField & ") AS MaxId FROM " & sTable & ";"
		'	WScript.Echo sQuery
		Set oRs = objDb.Execute(sQuery)
		nReturn = oRs("MaxId")
		If IsNull(nReturn) = True Then
			' WScript.Echo "NO VALUE"
			nReturn = 1
		Else	
			' WScript.Echo "VALUE " & nReturn & " FOUND"
			nReturn = nReturn + 1
		End If
		GetNextId = nReturn
	End Function ' GetNextId
	
	
	Public Function ExecQuery(ByVal strQuery)
		'
		'	Executes a SQL query.
		'	When errors occur, display the query and error message.
		'
		'	Returns values:
		'		0 when succesfull.
		'		other number as error code when unsuccesfull.
		'
		Dim	intReturn
	
		intReturn = 0
	
		On Error Resume Next
		objDb.Execute(strQuery)
		If Err.Number <> 0 Then
			WScript.Echo
			WScript.Echo "---[ SqlExec() ]----------------------"
			WScript.Echo "** ERROR:       0x" & Hex(Err.Number)
			WScript.Echo "** Description: " & Err.Description
			WScript.Echo "** Query:"
			WScript.Echo 
			WScript.Echo strQuery
			WScript.Echo 
			WScript.Echo "--------------------------------------"
			intReturn = Err.Number
		Else
			intReturn = 0
		End If
		ExecQuery = nReturn
	End Function ' Exec
	
	
	Public Function UniqueRecordId()
		'
		'	Returns a date-time code based on the previous value sDateTimeCode
		'
		'	sDatePrev = "20060522-125432-0002"
		'	sDatePrev = GetDateTimeCode(sDatePrev)	'' returns "20060522-125432-0003"
		'
		'	Init:
		'		Dim	gstrUniqueRecId		
		'		gstrUniqueRecId = UniqueRecordId()
		'
		'	Usage:
		'		Dim	x
		' 		For x = 1 To 10000
		'			WScript.Echo NumberAlign(x, 5) & ": " & UniqueRecordId()
		' 		Next
		'
	
		Dim	sDate
		Dim	sCodeCur
		Dim	sCodeNew
		Dim	sDateNew
		Dim	sDateCur
		Dim	nCodeCur
		Dim	intDate
		Dim intTime
		Dim	dNow
		Dim	strReturn
	
		' Get the new date time in a string
		dNow = Now()
	
		' Format it to YYYYMMDD-HHMMSS
		sDateNew = Year(dNow) & _
			NumberAlign(Month(dNow), 2) & _
			NumberAlign(Day(dNow), 2) & _
			"-" & _
			NumberAlign(Hour(dNow), 2) & _
			NumberAlign(Minute(dNow), 2) & _
			NumberAlign(Second(dNow), 2)

		' Extract the DateTime and Code 
		sCodeCur = Right(strUniqueRecId, 4)		' Code
		sDateCur = Left(strUniqueRecId, 15)		' DateTime 

		If sDateCur = sDateNew Then
			nCodeCur = Int(sCodeCur)
			nCodeCur = nCodeCur + 1
		Else
			nCodeCur = 0
		End If
	
		' Build the new UniqueRecordId
		strUniqueRecId = sDateNew & "-" & NumberAlign(nCodeCur, 4)
		' Return the code
	
		intDate = Year(dNow) & NumberAlign(Month(dNow), 2) & NumberAlign(Day(dNow), 2)
		intTime = NumberAlign(Hour(dNow), 2) & NumberAlign(Minute(dNow), 2) & NumberAlign(Second(dNow), 2)
	
		'WScript.Echo intDate
		'WScript.Echo intTime
		'Wscript.Echo nCodeCur
	
		'strReturn = "U-" & UCase(NumberAlign(Hex(intDate), 7) & "-" & NumberAlign(Hex(intTime), 5) & "-" & NumberAlign(Hex(nCodeCur), 4))
		strReturn = UCase(NumberAlign(Hex(intDate), 7) & NumberAlign(Hex(intTime), 5) & NumberAlign(Hex(nCodeCur), 4))
		'WScript.Echo strReturn
		UniqueRecordId = strReturn
	End Function ' UniqueRecordId

	
	Public Function GetRecCount(ByVal strTable)
		'
		'	Return the number of records in a table.
		'
		Dim	strQuery
		Dim	objRs
		Dim	intReturn
	
		intReturn = 0
	
		strQuery = "SELECT COUNT(*) AS rec_count "
		strQuery = strQuery & "FROM "
		strQuery = strQuery & strTable
		strQuery = strQuery & ";"
		
		Set objRs = objDb.Execute(strQuery)
		
		If objRs.Eof = False Then
			objRs.MoveFirst
			intReturn = objRs("rec_count")
		End If
		GetRecCount = intReturn
	End Function ' GetRecCount

	
	Public Sub ClearTable(ByVal strTable)
		'
		'	Delete all rows from a table strTable.
		'
		Dim	strQuery
		
		strQuery = "TRUNCATE TABLE " & strTable
		Call ExecQuery(strQuery)
	End Sub
	
	
	Public Sub DelRecord(ByVal strTable, ByVal strField, ByVal strValue)
		Dim	strQuery
		
		strQuery = "DELETE FROM " & strTable & " " & _
			"WHERE " & strField & "="
		
		If IsNumeric(strValue) = False Then
			strQuery = strQuery & FixStr(strValue)
		Else
			strQuery = strQuery & strValue
		End If
		
		strQuery = strQuery & ";"
		
		Call ExecQuery(strQuery)
	End Sub
End Class ' of Class ClassMySQL



Call ScriptInit()
'Call ScriptTest()
Call ScriptRun()
Call ScriptDone()
WScript.Quit()



'' End of script