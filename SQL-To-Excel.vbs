' ------------------------------------------- Const Variables (change if necessary) ------------------------------------------

const My_SQL_Server = ""
const My_SQL_DB = ""

const New_Disputes_Cols = "Vendor, Platform, Dispute_Category, STC_Claim_Number, Record_Type, BAN, Bill_Date, Billed_Amt, Claimed_Amt, Dispute_Reason, USI, USOC, Billed_Phrase_Code, Causing_SO, PON, CLLI, Usage_Rate, MOU, Jurisdiction, Short_Paid, Batch"
const Cred_Received_Cols = "Previous_Display_Status, INDEX, Claim_number, Bill_Date, Invoice_Date, Credit_Amount, Batch"
const Closed_Escalate_Cols = "Action Close Or Escalate, INDEX, STC_CLAIM_NUMBER, DISPUTE_CATEGORY, ACCOUNT_NUMBER, BILL_DATE, DISPUTE_AMOUNT, Norm_Close_Reason, Batch"
const Partial_Paid_to_Paid_Cols = "Previous_Display_Status, INDEX, Claim_Number, Bill_Date, Invoice_Date, Credit_Amount, Batch"
const Unwinnable_Cols = "DISPLAY_STATUS, INDEX, STC_CLAIM_NUMBER, DISPUTE_CATEGORY, ACCOUNT_NUMBER, BILL_DATE, Dispute_Amount, Norm_Close_Reason, BATCH"
const Cred_Corr_Cols = "Previous_Display_Status, INDEX, Claim_number, Bill_Date, Invoice_Date, Credit_Amount, Batch"

const Staging_TBL = ""

const Cred_Received_SQL_Path = "Batching SQL\Credits_Received.sql"
const Closed_Escalate_SQL_Path = "Batching SQL\Closed_Escalate.sql"
const Partial_Paid_to_Paid_SQL_Path = "Batching SQL\Partial_Paid_to_Paid.sql"
const Unwinnable_SQL_Path = "Batching SQL\Unwinnables.sql"
const Cred_Corr_SQL_Path = "Batching SQL\Credit_Corrections.sql"
const STC_Batch_Folder_Dir = ""

' ------------------------------------------- Main Code below ------------------------------------------

Public WshShell, oFSO, Log_Path, Filepath, networkInfo, myquery, Batch

dim Temp, myresults

Set WshShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set networkInfo = CreateObject("WScript.NetWork")

select case Weekday(Now())
	case 1
		Temp = DateAdd("d", -2, Now())
	case 2
		Temp = DateAdd("d", -3, Now())
	case 3
		Temp = DateAdd("d", -4, Now())
	case 4
		Temp = DateAdd("d", -5, Now())
	case 5
		Temp = DateAdd("d", -6, Now())
	case 6
		Temp = DateAdd("d", -7, Now())
	case 7
		Temp = DateAdd("d", -8, Now())
end select

if len(cstr(month(Temp)))=2 and len(cstr(day(Temp)))=2 then
	batch = cstr(year(Temp)) & cstr(month(Temp)) & cstr(day(Temp))
end if
if len(cstr(month(Temp)))=2 and not len(cstr(day(Temp)))=2 then
	batch = cstr(year(Temp)) & cstr(month(Temp)) & "0" & cstr(day(Temp))
end if
if not len(cstr(month(Temp)))=2 and len(cstr(day(Temp)))=2 then
	batch = cstr(year(Temp)) & "0" & cstr(month(Temp)) & cstr(day(Temp))
end if
if not len(cstr(month(Temp)))=2 and not len(cstr(day(Temp)))=2 then
	batch = cstr(year(Temp)) & "0" & cstr(month(Temp)) & "0" & cstr(day(Temp))
end if

' Filepath = oFSO.GetParentFolderName(WScript.ScriptFullName) & "\GRT_" & batch & "_disputes.xlsx"
Filepath = STC_Batch_Folder_Dir & "\GRT_" & batch & "_disputes.xlsx"

Log_Path = replace(WScript.ScriptFullName,".vbs","") & "_Error_Log.txt"

myquery = "select top 1 ID from " & Staging_TBL & " where isnull(status,'')='' and batch=" & batch
Query_ODS myquery, myresults, iserr
if isarray(myresults) then
	write_log Now() & " * Error * " & networkInfo.UserName & " * Plz email CDA to change Status to 'Delivered to STC'"
else
	myquery = "select " & New_Disputes_Cols & " from " & Staging_TBL & " where isnull(status,'')='Delivered to STC' and batch=" & batch
	Query_ODS myquery, myresults, iserr

	iserr = Create_Batch(myresults)

	if len(iserr) < 1  then
		msgbox("Batch has been created. Plz find it at " & oFSO.GetFileName(filepath))
	end if
end if

Set WshShell = nothing
Set oFSO = nothing
Set networkInfo = nothing

private function Create_Batch(mydata)
	On Error Resume Next

	dim myExcelWorker, oWorkBook, hcols, sheet, mylist, iserr, counter
	Set myExcelWorker = CreateObject("Excel.Application")
	strSaveDefaultPath = myExcelWorker.DefaultFilePath

	if oFSO.fileexists(Filepath) then
		oFSO.DeleteFile Filepath, True
	end if

	myExcelWorker.DisplayAlerts = False
	myExcelWorker.AskToUpdateLinks = False
	myExcelWorker.AlertBeforeOverwriting = False
	myExcelWorker.FeatureInstall = msoFeatureInstallNone

	myExcelWorker.DefaultFilePath = oFSO.GetParentFolderName(Filepath)

	Set oWorkBook = myExcelWorker.Workbooks.Add()

	If Err.Number <> 0 Then
		Create_Batch = err.description
		write_log Now() & " * Error * " & networkInfo.UserName & " * Create Temp Workbook (" & Err.Description & ")"
            
		myExcelWorker.DefaultFilePath = strSaveDefaultPath
		myExcelWorker.Quit
		Set myExcelWorker = Nothing
		Set oWorkBook = Nothing
		Exit Function
	End If

	if isarray(mydata) then
		counter = true
		hcols = split(replace(New_Disputes_Cols,", ",","),",")

		oworkbook.sheets(1).range("1:1048576").NumberFormat = "@"
		oworkbook.sheets(1).range("A1").resize(1,ubound(hcols) + 1) = hcols
		oworkbook.sheets(1).range("A2").resize(ubound(mydata, 1), ubound(mydata, 2)) = mydata
	end if

	oworkbook.sheets(1).Name = "New Disputes"

	iserr = SQL_Scripts(oFSO.GetParentFolderName(WScript.ScriptFullName) & "\" & Cred_Received_SQL_Path, mylist)
	if len(iserr) > 0 then
		Create_Batch = iserr
		myExcelWorker.DefaultFilePath = strSaveDefaultPath
		myExcelWorker.Quit
		Set myExcelWorker = Nothing
		Set oWorkBook = Nothing
		Exit Function
	end if

	if isarray(mylist) then
		counter = true
		set sheet = oworkbook.sheets.add

		hcols = split(replace(Cred_Received_Cols,", ",","),",")

		sheet.range("1:1048576").NumberFormat = "@"
		sheet.range("A1").resize(1,ubound(hcols) + 1) = hcols
		sheet.range("A2").resize(ubound(mylist, 1), ubound(mylist, 2)) = mylist

		sheet.Name = "Credits Received"
	end if

	iserr = SQL_Scripts(oFSO.GetParentFolderName(WScript.ScriptFullName) & "\" & Closed_Escalate_SQL_Path, mylist)
	if len(iserr) > 0 then
		Create_Batch = iserr
		myExcelWorker.DefaultFilePath = strSaveDefaultPath
		myExcelWorker.Quit
		Set myExcelWorker = Nothing
		Set oWorkBook = Nothing
		Exit Function
	end if

	if isarray(mylist) then
		counter = true
		set sheet = oworkbook.sheets.add

		hcols = split(replace(Closed_Escalate_Cols,", ",","),",")

		sheet.range("1:1048576").NumberFormat = "@"
		sheet.range("A1").resize(1,ubound(hcols) + 1) = hcols
		sheet.range("A2").resize(ubound(mylist, 1), ubound(mylist, 2)) = mylist

		sheet.Name = "Close or Escalate"
	end if

'	iserr = SQL_Scripts(oFSO.GetParentFolderName(WScript.ScriptFullName) & "\" & Partial_Paid_to_Paid_SQL_Path, mylist)
'	if len(iserr) > 0 then
'		Create_Batch = iserr
'		myExcelWorker.DefaultFilePath = strSaveDefaultPath
'		myExcelWorker.Quit
'		Set myExcelWorker = Nothing
'		Set oWorkBook = Nothing
'		Exit Function
'	end if
'
'	if isarray(mylist) then
'		counter = true
'		set sheet = oworkbook.sheets.add
'
'		hcols = split(replace(Partial_Paid_to_Paid_Cols,", ",","),",")
'
'		sheet.range("1:1048576").NumberFormat = "@"
'		sheet.range("A1").resize(1,ubound(hcols) + 1) = hcols
'		sheet.range("A2").resize(ubound(mylist, 1), ubound(mylist, 2)) = mylist
'
'		sheet.Name = "Partial Paid to Paid"
'	end if

	iserr = SQL_Scripts(oFSO.GetParentFolderName(WScript.ScriptFullName) & "\" & Unwinnable_SQL_Path, mylist)
	if len(iserr) > 0 then
		Create_Batch = iserr
		myExcelWorker.DefaultFilePath = strSaveDefaultPath
		myExcelWorker.Quit
		Set myExcelWorker = Nothing
		Set oWorkBook = Nothing
		Exit Function
	end if

	if isarray(mylist) then
		counter = true
		set sheet = oworkbook.sheets.add

		hcols = split(replace(Unwinnable_Cols,", ",","),",")

		sheet.range("1:1048576").NumberFormat = "@"
		sheet.range("A1").resize(1,ubound(hcols) + 1) = hcols
		sheet.range("A2").resize(ubound(mylist, 1), ubound(mylist, 2)) = mylist

		sheet.Name = "Unwinnable Claims"
	end if

	iserr = SQL_Scripts(oFSO.GetParentFolderName(WScript.ScriptFullName) & "\" & Cred_Corr_SQL_Path, mylist)
	if len(iserr) > 0 then
		Create_Batch = iserr
		myExcelWorker.DefaultFilePath = strSaveDefaultPath
		myExcelWorker.Quit
		Set myExcelWorker = Nothing
		Set oWorkBook = Nothing
		Exit Function
	end if

	if isarray(mylist) then
		counter = true
		set sheet = oworkbook.sheets.add

		hcols = split(replace(Cred_Corr_Cols,", ",","),",")

		sheet.range("1:1048576").NumberFormat = "@"
		sheet.range("A1").resize(1,ubound(hcols) + 1) = hcols
		sheet.range("A2").resize(ubound(mylist, 1), ubound(mylist, 2)) = mylist

		sheet.Name = "Credit Corrections"
	end if
	if not counter then
		Create_Batch = "No Batch"
		oworkbook.close
		Set myExcelWorker = Nothing
		Set oWorkBook = Nothing
		write_log Now() & " * Error * " & networkInfo.UserName & " * There is no batch this week"
		exit function
	end if

	oWorkBook.Saveas Filepath

	if Err.number = "91" Then
		Err.Clear
	end if

	If Err.Number <> 0 then
		Create_Batch = err.description
		write_log Now() & " * Error * " & networkInfo.UserName & " * Save Before Close " & "(" & Err.Description & ")"
		Err.Clear
		myExcelWorker.DefaultFilePath = strSaveDefaultPath
		myExcelWorker.Quit
		Set myExcelWorker = Nothing
		Set oWorkBook = Nothing
		if oFSO.fileexists(Filepath) then
			oFSO.DeleteFile Filepath, True
		end if
	End If

	oWorkBook.close

	If Err.Number <> 0 Then
		write_log Now() & " * Warning * " & networkInfo.UserName & " * Close " & "(" & Err.Description & ")"
		Err.Clear
	End If

	myExcelWorker.DefaultFilePath = strSaveDefaultPath

	Set myExcelWorker = Nothing
	Set oWorkBook = Nothing
end Function

private function SQL_Scripts(SQL_Filepath, mylist)
	Dim objFile, strLine

	if oFSO.fileexists(SQL_Filepath) then
		Set objFile = oFSO.OpenTextFile(SQL_Filepath)
		Do Until objFile.AtEndOfStream
			if len(strLine) > 0 then
				strLine= strLine & vbcrlf & objFile.ReadLine
			else
    				strLine= objFile.ReadLine
			end if
		Loop
		objfile.close

		Query_ODS strLine, mylist, iserr
		if iserr then
			SQL_Scripts = err.description
			if oFSO.fileexists(Filepath) then
				oFSO.DeleteFile Filepath, True
			end if
			write_log Now() & " * Error * " & networkInfo.UserName & " * SQL Script (" & SQL_Filepath & ") execution error"
		end if
	else
		SQL_Scripts = err.description
		write_log Now() & " * Error * " & networkInfo.UserName & " * SQL Script (" & SQL_Filepath & ") does not exist"
		if oFSO.fileexists(Filepath) then
			oFSO.DeleteFile Filepath, True
		end if
	end if

	set objFile = Nothing
end Function

private Sub Query_ODS(myquery, ReturnArray, iserr)
	On Error Resume Next
	Dim constr, conn, rs, myresults
	set ReturnArray = nothing

	constr = "Provider=SQLOLEDB;Data Source=" & My_SQL_Server & ";Initial Catalog=" & My_SQL_DB & ";Integrated Security=SSPI;"
	Set conn = CreateObject("ADODB.Connection")

	Set rs = CreateObject("ADODB.Recordset")

	conn.Open constr

	If Err.Number <> 0 Then
		iserr = err.description
		write_log Now() & " * Error * " & networkInfo.UserName & " * Open SQL Conn (" & Err.Description & ")"
		Set conn = Nothing
		exit sub
	end if

	conn.CommandTimeout = 1200

	rs.Open myquery, conn

	If Err.Number <> 0 Then
		iserr = err.description
		write_log Now() & " * Error * " & networkInfo.UserName & " * SQL Query (" & Err.Description & ")"
		Set conn = Nothing
		exit sub
	end if

	If Not rs.EOF Then
        	myresults = rs.getrows()
    	End If
    
	rs.Close

	If Err.Number <> 0 Then
		write_log Now() & " * Warning * " & networkInfo.UserName & " * SQL Close Conn (" & Err.Description & ")"
	end if

	conn.Close

	If Err.Number <> 0 Then
		write_log Now() & " * Warning * " & networkInfo.UserName & " * SQL Close Conn (" & Err.Description & ")"
	end if

	TransposeArray myresults, ReturnArray

    	Set rs = Nothing
	Set conn = Nothing
End Sub

Private Sub TransposeArray(ByRef InputArr, ByRef ReturnArray)
    If IsArray(InputArr) Then
        Dim RowNdx, ColNdx, delim, delim2
        Dim LB1, LB2, UB1, UB2
        LB1 = LBound(InputArr, 1)
        LB2 = LBound(InputArr, 2)
        UB1 = UBound(InputArr, 1)
        UB2 = UBound(InputArr, 2)
        
        set ReturnArray = nothing
        ReDim ReturnArray(UB2+1, UB1+1)
    
        For RowNdx = LB2 To UB2
        	For ColNdx = LB1 To UB1
            		If IsNull(InputArr(ColNdx, RowNdx)) Then
                		ReturnArray(RowNdx, ColNdx) = ""
            		Else
                		ReturnArray(RowNdx, ColNdx) = InputArr(ColNdx, RowNdx)
            		End If
		Next
        Next
    End If
End Sub

Sub Write_Log(ByVal text)
	msgbox(text)
	Dim objFile, strLine

	if oFSO.fileexists(Log_Path) then
		Set objFile = oFSO.OpenTextFile(Log_Path)
		Do Until objFile.AtEndOfStream
			if len(strLine) > 0 then
				strLine= strLine & vbcrlf & objFile.ReadLine
			else
    				strLine= objFile.ReadLine
			end if
		Loop
		objfile.close
		Set objfile = oFSO.CreateTextFile(Log_Path,True)
		objfile.write strLine & vbcrlf & text
		objfile.close
	else

		Set objfile = oFSO.CreateTextFile(Log_Path,True)
		objfile.write text
		objfile.close
	end if

	set objFile = Nothing
End Sub

Function IsArray(anArray)
    Dim I
    On Error Resume Next
    I = UBound(anArray, 1)
    If Err.Number = 0 Then
        IsArray = True
    Else
        IsArray = False
    End If
End Function
