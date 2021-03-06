Public OIEPgObj_ExpenseHome,OIEPgObj_GenPref,OIEPgObj_EmpBankDet,OIEPgObj_CERGenInfo,OIEPgObj_CERCashExp,OIEPgObj_CashExpDet,OIEPgObj_ExpAllocation,OIEPgObj_review
Set OIEPgObj_ExpenseHome = Browser("name:=Expense Home").Page("title:=Expense Home")
Set OIEPgObj_ExpenseSearch = Browser("name:=Expense Reports").Page("title:=Expense Reports")
Set OIEPgObj_GenPref = Browser("name:=General Preferences").Page("title:=General Preferences")
Set OIEPgObj_EmpBankDet = Browser("name:=Employee Bank Details Page").Page("title:=Employee Bank Details Page")
Set OIEPgObj_CERGenInfo = Browser("name:=Create Expense Report: General Information").Page("title:=Create Expense Report: General Information")
Set OIEPgObj_CERCashExp = Browser("name:=Create Expense Report: Cash and Other Expenses").Page("title:=Create Expense Report: Cash and Other Expenses")
Set OIEPgObj_CashExpDet = Browser("name:=Cash and Other Expenses: Details for Line 1").Page("title:=Cash and Other Expenses: Details for Line 1")
Set OIEPgObj_updateExpRep = Browser("name:=Update Expense Report.*").Page("title:=Update Expense Report.*")
Set OIEPgObj_ExpAllocation = Browser("name:=Create Expense Report: Expense Allocations").Page("title:=Create Expense Report: Expense Allocations")
Set OIEPgObj_review = Browser("name:=Create Expense Report: Review").Page("title:=Create Expense Report: Review")
Set OIEPgObj_Confirmation = Browser("name:=Expense Report.*").Page("title:=Expense Report.*")
Set OIEPgObj_NotificationDet = Browser("name:=Notification Details").Page("title:=Notification Details")
Set OIEPgObj_ExpenseAudit = Browser("name:=Oracle Internet Expenses Audit").Page("title:=Oracle Internet Expenses Audit")
Set orcNavigatorWindowObj = OracleFormWindow("title:=Navigator.*")
Set orcFRWindowObj = OracleFormWindow("title:=Find Requests.*")
Set orcSubReqObj = OracleFormWindow("short title:=Submit Request")
Set orcFlxWinObj = OracleFlexWindow("title:=Parameters")
Set orcDecisionObj = OracleNotification("title:=Decision")
Set orcSubNewReqObj = OracleFormWindow("title:=Submit a New Request")
'=============================================================
'*************************************************************************
'IExpense WebElemnts XPaths
'=============================================================
'*************************************************************************
Const empnum_xpath = "xpath:=//*[@id='employeeNum']"
Const cntryName_xpath = "xpath:=//*[@id='BankFlex0']"
Const bankBranch_code = "xpath:=//*[@id='BankFlex1']"
Const accNum_code = "xpath:=//*[@id='BankFlex2']"
Const benfStatus_code = "xpath:=//*[@id='BankFlex4']"
Const statusTable_xpath = "xpath:=//*[@id='TrackReportsRN.TrackSubExpRepTable']//TR[2]//TABLE"
Const IEHomeLink_xpath = "xpath:=//*[@id='OIEHOMEPAGE']"
Const prefrence_xpath = "xpath:=//*[@class='x6w']//td[7]/a"
Const empBankDtl_xpath = "xpath:=//*[@id='MMC_GLB_OIE_EMP_BANK_DTLS']"
Const SaveBtn_xpath = "xpath:=//BUTTON[@id='save']"
Const createExpRep_xpath = "xpath:=//button[@id='CreateButton']"
Const purpose_xpath = "xpath:=//*[@id='Purpose']"
Const projSrc_xpath = "xpath:=//*[@id='HeaderDFF0']"
Const projCode_xpath = "xpath:=//*[@id='HeaderDFF1']"
Const projCode2_xpath = "xpath:=//*[@id='DFF_346513']"
Const busiPur_xpath = "xpath:=//*[@id='HeaderDFF2']"
Const Next_xpath = "xpath:=//*[@id='pbb']//td[8]/button"
Const date_xpath = "xpath:=//*[@id='N51:Date:0']"
Const recAmt_xpath = "xpath:=//*[@id='N51:ReceiptCurrencyAmount:0']"
Const expType_xpath = "xpath:=//*[@id='N51:WebParameterId:0']"
Const desc_xpath = "xpath:=//*[@id='N51:Justification:0']"
Const dff_xpath = "xpath:=//*[@id='OIECashAndOtherList']//TR[2]/TD[8]/A/img"
Const merchName_xpath = "xpath:=//*[@id='DetailMerchantName']"
Const Next1_xpath = "xpath:=//*[@id='pbb']//td[10]/button"
Const attstn_xpath = "xpath:=//*[@id='AgreementCheckBox']"
Const submitBtn_xpath = "xpath:=//*[@id='OIESubmit']"
Const itemizeBtn_xpath = "xpath:=//BUTTON[@id='ItemizeButton']"
Const entity_xpath="xpath:=//*[@id='N55:KffSEGMENT1:1']"
Const RCfield_xpath="xpath:=//*[@id='N55:KffSEGMENT4:1']"
Const interComp_xpath="xpath:=//*[@id='N55:KffSEGMENT5:1']"
Const Next2_btn = "xpath:=//*[@id='pbb']//td[10]/button"
Const expText_xpath = "xpath:=//TABLE[@id='FwkErrorBeanId']//DIV[3]"
Const approverTab = "xpath:=//*[@id='ConfirmSubTabsRN']//TD[8]/A"
Const approverName = "xpath:=//*[@id='HrApproversListTable']/TABLE[2]"
Const logout_xpath = "xpath:=//*[@id='ConfirmationPG']//TR[2]//TD[5]/A"
Const approvebtn_xpath = "xpath:=//*[@id='rowLayout']/TD[2]/button"
Const missingRec_xpath = "xpath:=//*[@id='DetailReceiptMissing']"
Const searchByList_xpath = "xpath:=//*[@id='SearchByPoplist']"
Const searchInput_xpath = "xpath:=//*[@id='QuickSearchInput']"
Const goBtn_xpath = "xpath:=//*[@id='GoButton']"
Const receiveDate_xpath = "xpath:=//INPUT[@id='ReceiptPackageReceivedDate']"
'Const reciptVerifiedCB_xpath = "xpath:=//*[@id='N156:ReceiptVerified1:1']"
Const reciptVerifiedCB_xpath = "xpath:=//*[@id='N156:ReceiptVerified1:1_rc']"
Const applyBtn_xpath = "xpath:=//*[@id='AuditActionButton']"
Const cnfText_xpath = "xpath:=//*[@id='FwkErrorBeanId']//DIV[3]"
Const noRows_xpath = "xpath:=//SPAN[@id='NoRowsPrompt']"
Const aprvrNotif_xpath = "xpath:=//*[@id='NtfWorklist']/TABLE[2]"
Const expSearchLink_xpath = "xpath:=//*[@id='OIE_EXPENSE_REPORT_SEARCH']"
Const RepNumFld_xpath = "xpath:=//*[@id='SearchReportNum']"
Const searchGoBtn_xpath = "xpath:=//button[contains(text(),'Go')]"
Const searchResTbl_xpath = "xpath:=//SPAN[@id='HistoryResultsTbl']/TABLE[2]"
Const giftRecEmpName_xpath="xpath:=//*[@id='EmployeeTableRN']/TABLE[2]//TR[2]/TD[2]//INPUT"
Const newReqBtn = "description:=Submit a New Request"
'Const ledgerField = "prompt:=Ledger"
'Const endDateField = "prompt:=End Date"
Const repName = "description:=Name"
Const subBtn = "description:=Submit"
Const okBtn = "label:=Ok"
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  06/10/2021
'Name -                   fn_checkEmployeeBankDetails 
'Description:             fn_checkEmployeeBankDetails :  will check employee bank details
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=========================================================================
Function fn_checkEmployeeBankDetails()
    blnResultFlag = False
    On Error Resume Next
    If (fn_exist (OIEPgObj_ExpenseHome.Link(IEHomeLink_xpath))) Then
        fnReportEvent "Pass", "Oracle Iexpense page navigation status","Successfully navigated to OIE Home page ",False
        fn_Click OIEPgObj_ExpenseHome.Link(prefrence_xpath)
        fn_Click OIEPgObj_GenPref.Link(empBankDtl_xpath)
        If gstrTestCaseExec_id = "GSI.P2P.IE.SA.002" Then
            Call fn_validateBankDetails
        Else
            Call fn_getBankDetails
        End If
        blnResultFlag = True
    Else
        fnReportEvent "Fail","Oracle Iexpense page navigation status"," Failed to navigate OIE Home page",True
    End If
    fn_checkEmployeeBankDetails = blnResultFlag
    If err.number <> 0 Then
        fn_checkEmployeeBankDetails = False
        fnReportEvent "Fail","check bank details"," Failed to validate bank details",True
    End If
End Function
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  06/10/2021
'Name -                     fn_getBankDetails 
'Description:             fn_getBankDetails :  will get employee bank details from UI
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=========================================================================
Function fn_getBankDetails()
On error resume next 
fn_getBankDetails=false
If fn_exist (OIEPgObj_EmpBankDet)=true Then
fnReportEvent "Pass","Employee Bank Details Page","Employee Bank Details Page loaded successfully",true
	varEmpNum = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebElement(empnum_xpath),"innertext")
	If varEmpNum <> ""  Then
		fnReportEvent "Pass","Employee Number","Employee Number " & varEmpNum & " is present",False
	Else
		fnReportEvent "Fail","Employee Number","Employee Number does not Exist",True
	End If
	
	varCntryName = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebList(cntryName_xpath),"value")
	If varCntryName <> ""  Then
		fnReportEvent "Pass","Country Name","Bank Details-Country " & varCntryName & " is present",False
	Else
		fnReportEvent "Fail","Country Name","Country doesnt Exist",True
	End If
	
	varBankBCode = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebEdit(bankBranch_code),"value")
	If varBankBCode <> "" Then
		fnReportEvent "Pass","Bank and Branch Code","Bank and Branch Code " & varBankBCode & " is present",False
	Else
		fnReportEvent "Fail","Bank and Branch Code","Bank and Branch Code doesnt Exist",True
	End If
	
	varAccNum = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebEdit(accNum_code),"value")
	If varAccNum <> "" Then
		fnReportEvent "Pass","Account Number","Account Number " & varAccNum & " is present",False
	Else
		fnReportEvent "Fail","Account Number","Account Number doesnt Exist",True
	End If
	
	varBenfStatus = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebList(benfStatus_code),"value")
	If varBenfStatus <> "" Then
		fnReportEvent "Pass","Beneficiary Status","Beneficiary Status " & varBenfStatus & " is present",False
	Else
		fnReportEvent "Fail","Beneficiary Status","Beneficiary Status doesnt Exist",True
	End If
Else 
	fnReportEvent "Fail","Employee Bank Details Page","Unable to load Employee Bank Details Page",true
	fn_getBankDetails=false
	Exit Function 
End If

fn_getBankDetails=true

If err.number <> 0 Then
	fn_getBankDetails = False
	fnReportEvent "Fail","Employee Bank Details"," Failed to check bank details. Error is : "&Err.description,True
	Exit Function
End If
End Function
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  06/10/2021
'Name -                   fn_validateBankDetails 
'Description:             fn_validateBankDetails :will validate employee bank details with test data
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=========================================================================
Function fn_validateBankDetails()
    On Error Resume Next
    varEmpNum = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebElement(empnum_xpath),"innertext")
    strQuery="UPDATE [ExecutionResult$] SET Employee_No='"&varEmpNum&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"
    Call fn_updateQuery(strQuery)
    If varEmpNum = gb_TestDataDic.item("Employee_Number") Then
        fnReportEvent "Pass","Employee Number","Employee Number is same as test data",False
    Else
        fnSet_FieldName OIEPgObj_EmpBankDet.WebElement(empnum_xpath), gb_TestDataDic.item("Employee_Number"),"Employee Number"
    End If
    varCntryName = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebList(cntryName_xpath),"value")
    If varCntryName = gb_TestDataDic.item("Country_Name") Then
        fnReportEvent "Pass","Country name","Country name is same as test data",False
    Else
        fnSet_FieldName OIEPgObj_EmpBankDet.WebList(cntryName_xpath), gb_TestDataDic.item("Country_Name"),"Country name"
    End If
   varBankBCode = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebEdit(bankBranch_code),"value")
   If varBankBCode = gb_TestDataDic.item("Bank_Branch_code") Then
       fnSet_FieldName OIEPgObj_EmpBankDet.WebEdit(bankBranch_code), gb_TestDataDic.item("Enter_Bank_Branch_code"),"Bank and Branch Code"
       fnReportEvent "Pass","Bank and Branch Code","Edited Data in field : Bank and Branch Code",False 
   Else
     fnReportEvent "Pass","Bank and Branch Code","Bank and Branch Code is same as test data",False
    End If
   varAccNum = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebEdit(accNum_code),"value")
   If varAccNum = gb_TestDataDic.item("Account_Number") Then
   	 fnSet_FieldName OIEPgObj_EmpBankDet.WebEdit(accNum_code), gb_TestDataDic.item("Enter_Account_Number"),"Account Number"
        fnReportEvent "Pass","Account Number","Edited Data in field : Account Number",False
   Else
        fnReportEvent "Pass","Account Number","Account Number is same as test data",False
   End If
    Call fn_CaptureAccountNo
    varBenfStatus = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebList(benfStatus_code),"value")
    If varBenfStatus = gb_TestDataDic.item("Beneficiary_Status") Then
        fnReportEvent "Pass","Beneficiary Status","Beneficiary Status is same as test data",False
    Else
        fnSet_FieldName OIEPgObj_EmpBankDet.WebList(benfStatus_code), gb_TestDataDic.item("Beneficiary_Status"),"Beneficiary Status"
    End If
    fn_Click OIEPgObj_EmpBankDet.WebButton(SaveBtn_xpath)
'     If OIEPgObj_EmpBankDet.WebElement("html id:=FwkErrorBeanId").GetROProperty("innertext")="ConfirmationYour bank details are saved" then 
     If Instr(1,OIEPgObj_EmpBankDet.WebElement("html id:=FwkErrorBeanId").GetROProperty("innertext"),"Confirmation")>0 Then
        fnReportEvent "Pass","Bank Details Save Status","Confirmation message displayed & Bank Details saved successfully",False    
    Else 
        fnReportEvent "Fail","Bank Details Save Status","Confirmation message not displayed. Unable to save Bank Details",True
    End If
'   OIEPgObj_EmpBankDet.WebEdit(bankBranch_code)=Empty
'   OIEPgObj_EmpBankDet.WebEdit(accNum_code)=Empty
'   fn_Click OIEPgObj_EmpBankDet.WebButton(SaveBtn_xpath)
    If err.number <> 0 Then
        fnReportEvent "Fail","validate bank details"," Failed to set the bank details",True
    End If
    
End Function


Function fn_CaptureAccountNo()
On error resume next 
fn_CaptureAccountNo=false
	If fn_exist(OIEPgObj_EmpBankDet)=true Then		
		var_Account_No = OIEPgObj_EmpBankDet.WebEdit(accNum_code).GetROProperty("value")
		strQuery="UPDATE [ExecutionResult$] SET Account_No='"&var_Account_No&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"
		Call fn_updateQuery(strQuery)
		fn_CaptureAccountNo=true
	Else 
		fnReportEvent "Fail","Employee Bank Details Window Status","Unable to Load Employee Bank Details Window",true   
	End If
	
	If err.number <> 0 Then
		fnReportEvent "Fail","Capture Account No","Unable to capture Account No. Error is : "&Err.description,True
	End If
End Function


'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  06/10/2021
'Name -                   fn_cashExpenseClaim 
'Description:             fn_cashExpenseClaim :create an Expense report with cash expense claim 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=========================================================================
Function fn_cashExpenseClaim()
blnResultFlag = False
d = fn_getSysdateFormat("DD-MMM-YYYY")
On Error Resume Next
If fn_exist(OIEPgObj_ExpenseHome.Link(IEHomeLink_xpath)) = True  Then
	fnReportEvent "Pass", "Oracle Iexpense page navigation status","Successfully navigated to OIE Home page ",False
	
	fn_Click OIEPgObj_ExpenseHome.WebButton(createExpRep_xpath)
	fn_setGeneralInfo()
	fn_setCashExpenseDetails()      
	fn_click OIEPgObj_CERCashExp.Image(dff_xpath)
	fnSet_FieldName OIEPgObj_CashExpDet.WebEdit(merchName_xpath),gb_TestDataDic.item ("Merchant_Name"),"Merchant_Name"
		If gb_TestDataDic.item ("Is_missingReceipts")=Y  Then
			fn_Click OIEPgObj_CashExpDet.WebCheckBox(missingRec_xpath)   
		End  If
		fnSet_FieldName OIEPgObj_CashExpDet.WebEdit(projCode2_xpath),gb_TestDataDic.item ("Project Code"),"Project code"

		If gb_TestDataDic.item ("Expense Type")="Gifts" Then
			fnSet_FieldName OIEPgObj_CashExpDet.WebEdit(giftRecEmpName_xpath),gb_TestDataDic.item ("Gift_Emp_Name"),"Gift Employee Name"
		End If
	fn_Click OIEPgObj_CashExpDet.WebButton(itemizeBtn_xpath)
	fn_Click OIEPgObj_CashExpDet.WebButton(returnBtn_xpath)
	fn_Click OIEPgObj_CERCashExp.WebButton(Next1_xpath)
'	fn_checkRCDetails()
	If fn_checkRCDetails=false Then
		fnReportEvent "Fail", "Check RC Details","RC Details did not match",false
		fn_cashExpenseClaim=false
		Exit Function
	End If
	fn_Click OIEPgObj_ExpAllocation.WebButton(Next2_btn)
	fn_Click OIEPgObj_review.WebCheckBox(attstn_xpath)
	fn_Click OIEPgObj_review.WebButton(submitBtn_xpath)
	expRep_No= fn_getExpenseReportConfirmation()
'        If sendRec_flag  Then
'         fn_sendReceipts(expRep_No)	
'        End If
	Call fn_getApprover()
	blnResultFlag = True
Else
	fnReportEvent "Fail", "Oracle Iexpense page navigation status","Unable to navigate to OIE Home page ",true
	blnResultFlag = False
	Exit Function
End If

If err.number <> 0 Then
	fnReportEvent "Fail","Cash Expense report creation failed"," Failed to input details for Cash Expense report. & Error is : "&Err.description,True
	Exit Function
End If

fn_cashExpenseClaim = blnResultFlag
End Function


Function fn_setCashExpenseDetails()
On error resume next
d = fn_getSysdateFormat("DD-MMM-YYYY")
'	If fn_exist (OIEPgObj_CERCashExp)=true Then
		fnSet_FieldName OIEPgObj_CERCashExp.WebEdit(date_xpath),d,"Current Date"
		fnSet_FieldName OIEPgObj_CERCashExp.WebEdit(recAmt_xpath),gb_TestDataDic.item ("Receipt Amount"),"Receipt Amount"      
		Call fn_SelectWeblist(OIEPgObj_CERCashExp.WebList(expType_xpath), gb_TestDataDic.item ("Expense Type"),"Expense Type")
		fnSet_FieldName OIEPgObj_CERCashExp.WebEdit(desc_xpath),gb_TestDataDic.item ("Description"),"Description"
'	Else
'		fnReportEvent "Fail","Fail : Cash and Other Expenses Window Status","Unable to load Cash and Other Expenses Information Window",false  
'		fn_setCashExpenseDetails=false 
'		Exit Function
'	End If
	fn_setCashExpenseDetails=true

	If err.number <> 0 Then
		fnReportEvent "Fail","Cash and Other Expenses"," Failed to input Cash and Other Expenses Info. Error is : "&Err.Description,True
		fn_setCashExpenseDetails=false
		Exit function
	End If
End Function

Function fn_getExpenseReportConfirmation()
On Error Resume Next
fn_getExpenseReportConfirmation=false
strTest = "has been submitted."
	If fn_exist(OIEPgObj_Confirmation)=true Then
		fnReportEvent "Pass", "Expense Report Status","Successfully generated Expense Report",False
		var_Expno = OIEPgObj_Confirmation.WebElement(expText_xpath).GetROProperty("innertext")
		ExpNo = Split(var_Expno," ")
		strQuery = "UPDATE [ExecutionResult$] SET Expense_report_No='" & ExpNo(3) & "' where TC_ID='" & gstrTestCaseExec_id & "' and Start_Date='"&TstExecStart&"'"
		
		Call fn_updateQuery(strQuery)
		strReqConfirmStmt = OIEPgObj_Confirmation.WebElement(expText_xpath).GetROProperty("innertext")
			If InStr(strReqConfirmStmt,strTest) > 0 Then
				print ("Confirmation found.Expense Report  " & ExpNo(3) & " has been submitted")
				fnReportEvent "Pass","Confirmation","Confirmation found.Expense Report " & ExpNo(3) & " has been submitted",True
			Else
				fnReportEvent "Fail","Confirmation","Confirmation not found",True
			End If
			fn_getExpenseReportConfirmation =ExpNo(3)
	Else 
		fnReportEvent "Fail", "Expense Report Status","Unable to generate Expense Report",False
	Exit Function
	End If
	fn_getExpenseReportConfirmation=true
	
	If err.number <> 0 Then
		fnReportEvent "Fail","Expense Report Confirmation Status"," Failed to generate Expense Report. Error is : "&Err.Description,True
		fn_getExpenseReportConfirmation=false
		Exit function
	End If
End Function

Function fn_getApprover()
On error resume next 
fn_getApprover=false 
If fn_exists(OIEPgObj_Confirmation)=true Then
	fnReportEvent "Pass", "Expense Report Status","Successfully generated Expense Report",False
	fn_Click OIEPgObj_Confirmation.Link(approverTab)
	var_Row = OIEPgObj_Confirmation.WebTable(approverName).GetROProperty("rows")
	count = 1
	For i = 2 To var_Row
		var_appname = OIEPgObj_Confirmation.WebTable(approverName).GetCellData(i,2)
		strQuery = "UPDATE [ExecutionResult$] SET OIE_Approver" & count & "='" & var_appname & "' where TC_ID='" & gstrTestCaseExec_id & "' and Start_Date='" & TstExecStart & "'"
		Call fn_updateQuery(strQuery)
		If var_appname <> ""  Then
			fnReportEvent "Pass","Supervisor Name","Supervisor Name exists",False
			gstrApproverLogin = True	
		End If
	count = count + 1
	Next
Else 
	fnReportEvent "Fail", "Supervisor Name Status","Unable to get Supervisor Name",False
	fn_getApprover=false
	Exit Function
End If

fn_getApprover=true

If err.number <> 0 Then
	fnReportEvent "Fail","Supervisor Name Status","Failed to get Supervisor Name. Error is : "&Err.Description,True
	fn_getApprover=false
	Exit function
End If
End Function

Function fn_approveExpenseReport()
blnResultFlag = False
On Error Resume Next
varCnfNo = fn_getExecutionResultData(gstrTestCaseExec_id,"Expense_report_No")
If len(varCnfNo)=0 or varCnfNo="" or Isnull(varCnfNo) Then
	fnReportEvent "Fail","Expense Report No","Failed to fetch the Expense Number from the execution sheet",false
	fn_approveExpenseReport=false
	Exit Function
End If
resultctr = 1 
If fn_exist (OIEPgObj_ExpenseHome) = True Then
	var_Row = OIEPgObj_ExpenseHome.WebTable(aprvrNotif_xpath).GetROProperty("rows")
	For i = 2 To var_Row
	var_subject = OIEPgObj_ExpenseHome.WebTable(aprvrNotif_xpath).GetCellData(i,3)
	If InStr(var_subject,varCnfNo) > 0 Then
		fn_Click OIEPgObj_ExpenseHome.Link("xpath:=//*[@id='NtfWorklist']//tr[" & i & "]/td[3]/a")
		validateFlag = fn_validateReport()
		If validateFlag Then
			fn_Click OIEPgObj_NotificationDet.WebButton(approvebtn_xpath)
			blnResultFlag = True
			fnReportEvent "Pass","Validate expense report- Approver Login","Approver validated Expense report successfully",False
			resultctr=0
	Exit For
		End If 

'		Else
'			fnReportEvent "Fail","Validate expense report- Approver Login","Approver failed to validate Expense report",True
'		End If
	Else 
		fnReportEvent "Fail","Validate expense ID in worklist","Failed to find Expense Report ID" &varCnfNo,True
	End If
	Next
	If resultctr >= 1 Then
		fnReportEvent "Fail","Validate expense report- Approver Login","Approver failed to validate Expense report",True
	End If
Else
	fnReportEvent "Fail","Page Does not exist","Expense HomePage was not found-Approver Login .",True
End If
    
fn_approveExpenseReport = blnResultFlag
    
If err.number <> 0 Then
	fn_approveExpenseReport = False
	fnReportEvent "Fail","Approve Expense report failed"," failed to approve the Expense report",True
	Exit Function
End If
End Function

Function fn_validateReport()
    validateFlag = True
    fn_validateReport = validateFlag
End Function

Function fn_checkStatus()
On error resume next
    ' resultSearchCount=1
blnResultFlag = False
varCnfNo = fn_getExecutionResultData(gstrTestCaseExec_id,"Expense_report_No")
If len(varCnfNo)=0 or varCnfNo="" or Isnull(varCnfNo) Then
	fnReportEvent "Fail","Expense Report No","Failed to fetch the Expense Number from the execution sheet",false
	fn_checkStatus=false
	Exit Function
End If
print varCnfNo
	If  fn_Exist(OIEPgObj_ExpenseHome) = True Then
	
	var_status = fn_searchExpenseReport(varCnfNo)
	'resultSearchCount=0
		If var_status = "Pending Payables Approval" Then
			fnReportEvent "Pass","Report Status","Report Status is Pending Payables Approval",True
			blnResultFlag = True
		ElseIf  var_status = "Ready for Payment" Then
			fnReportEvent "Pass","Report Status","Report Status is Ready for Payment",True
			blnResultFlag = True
		ElseIf  var_status = "Pending Manager Approval" Then
			exception_Apr = fn_getExecutionResultData(gstrTestCaseExec_id,"OIE_Approver2")
			var_curApr = OIEPgObj_ExpenseHome.WebTable(searchResTbl_xpath).GetCellData(i,5)
			If var_curApr = exception_Apr Then
				fnReportEvent "Pass","Exception Approver","Expense report passed to Exception Approver successfully",True
				blnResultFlag = True
			Else
				fnReportEvent "Fail","Exception Approver","Expense report failed to pass to Exception Approver",True
			End If        
		Else
			fnReportEvent "Fail","Check Status","Incorrect status Updated",True
		End If
	'            End If
	'        Next      
	Else
		fnReportEvent "Fail","Search Expense Number","Fail to search the Expense number & value is =" & varCnfNo,True
	End If
fn_checkStatus = blnResultFlag
If err.number <> 0 Then
	fnReportEvent "Fail","Search Expense report failed"," failed to search the Expense report",false
	fn_checkStatus=false
	Exit Function
End If
End Function

Function fn_sendReceipts()
Dim objOutlook
Dim objOutlookMsg
'Dim olMailItem
On Error resume next
blnResultFlag = False
varCnfNo = fn_getExecutionResultData(gstrTestCaseExec_id,"Expense_report_No")
If len(varCnfNo)=0 or varCnfNo="" or Isnull(varCnfNo) Then
	fnReportEvent "Fail","Expense Report No","Failed to fetch the Expense Number from the execution sheet",false
	fn_sendReceipts=false
	Exit Function
End If
'fn_Click OIEPgObj_Confirmation.Link("name:="&varCnfNo)
'Browser("name:=Expense Report.*").Dialog("text:=Internet Explorer").WinButton("text:=&Allow").Click

Set objOutlook = CreateObject("Outlook.Application")
Set objOutlookMsg = objOutlook.CreateItem(0)
email_id = gb_TestDataDic.item("Email_ID")
'objOutlookMsg.To = "CONS.APAC.ExpenseReports.SIT@mmc.com"
objOutlookMsg.To = email_id
 
'doc = "C:\Users\u1207547\Downloads\Test Receipt attachment.pdf"
doc = "C:\Code Repository\RISHybridFramework\TestData\TestDocument.txt"

objOutlookMsg.Subject = ""&varCnfNo&""
objOutlookMsg.Body = "This is a test"

objOutlookMsg.Attachments.Add(doc)
objOutlookMsg.Display
objOutlookMsg.Send
blnResultFlag=true
'objOutlook.Quit
Set objOutlook = Nothing
Set objOutlookMsg = Nothing

	If err.number <> 0 Then
		fnReportEvent "Fail","Send receipts Email"," Failed to Send receipts Email. Error is : "&Err.description,True
		Exit Function
	End If
fn_sendReceipts=blnResultFlag
End Function

Function fn_run_ExpenseReportProgram()
    fn_run_ExpenseReportProgram = False
    orcNavigatorWindowObj.SelectMenu "View->Requests"
    orcFRWindowObj.OracleButton(newReqBtn).Click
    orcSubReqObj.OracleTextField(repName).Enter gb_TestDataDic.item ("repName_1")
    orcFlxWinObj.Approve
    orcSubReqObj.OracleButton(subBtn).Click
    If orcDecisionObj.Exist(5) Then
        strRequestNum = f_GetNumericValueFromString(orcDecisionObj.GetROProperty("message"))
        orcDecisionObj.Decline
    End If
    orcFRWindowObj.CloseWindow
    Call f_CheckRequestStatus(strRequestNum)
    strRequestNum = Empty
    fn_run_ExpenseReportProgram = True
End Function
Function fn_Bank_InterfaceProgram()
    blnResultFlag = False
    orcNavigatorWindowObj.SelectMenu "View->Requests"
    orcFRWindowObj.OracleButton(newReqBtn).Click
    orcSubReqObj.OracleTextField(repName).Enter gb_TestDataDic.item ("repName_1")
    orcSubReqObj.OracleButton(subBtn).Click
    OracleNotification("Caution").Approve
    If orcDecisionObj.Exist(5) Then
        strRequestNum = f_GetNumericValueFromString(orcDecisionObj.GetROProperty("message"))
        orcDecisionObj.Decline
    End If
    orcFRWindowObj.CloseWindow
    Call f_CheckRequestStatus(strRequestNum)
    strRequestNum = Empty
    blnResultFlag = True
    fn_Bank_InterfaceProgram = blnResultFlag
End Function
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  06/10/2021
'Name -                      fn_expenseAudit 
'Description:             fn_expenseAudit :Search and complete audit on any Expense Report Number
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=========================================================================
Function fn_expenseAudit() 
    On Error Resume Next
    var_date = fn_getSysdateFormat("DD-MMM-YYYY")
    str_cnfText = "The audit for this expense report is complete."
    blnResultFlag = False
    If fn_exist (OIEPgObj_ExpenseAudit) = True Then
        Call fn_SelectWeblist(OIEPgObj_ExpenseAudit.WebList(searchByList_xpath),"Expense Report Number","Expense Report Number")
        var_expNo = fn_getExecutionResultData(gstrTestCaseExec_id,"Expense_report_No")
        fnSet_FieldName OIEPgObj_ExpenseAudit.WebEdit(searchInput_xpath),var_expNo,"Expense Report Number"
        fn_Click OIEPgObj_ExpenseAudit.WebButton(goBtn_xpath)
        If OIEPgObj_ExpenseAudit.WebElement(noRows_xpath).Exist(5) Then
            fnReportEvent "Fail","Expense report not found  ","Expense report not found .Please check details " & var_expNo,False
        Else
            fnSet_FieldName OIEPgObj_ExpenseAudit.WebEdit(receiveDate_xpath),var_date,"Original Receipts Package Received Date"
            fn_Click OIEPgObj_ExpenseAudit.WebCheckBox(reciptVerifiedCB_xpath)
            fn_Click OIEPgObj_ExpenseAudit.WebButton(applyBtn_xpath)
            var_cnfTxt = fn_GetROPropertyValueByPropName (OIEPgObj_ExpenseAudit.WebElement(cnfText_xpath),"innertext")
            If InStr(var_cnfTxt,str_cnfText) > 0 Then
                fnReportEvent "Pass","Audit  ","Successfully Audited the expense report " & var_expNo,False
                blnResultFlag = True
            Else
                fnReportEvent "Fail","Audit  ","Failed to Audit the expense report " & var_expNo & ".Please check details",False
            End If
        End If
    End If
    fn_expenseAudit = blnResultFlag
    
    If err.number <> 0 Then
        fn_expenseAudit = False
        fnReportEvent "Fail","Audit expense report failed"," Failed to complete audit  Expense report  " & err.description ,True
    End If
End Function
Function fn_run_CreateAccounting()
    blnResultFlag = False
    d = fn_getSysdateFormat()
    fn_Click orcSubNewReqObj.OracleButton(okBtn)
    orcSubReqObj.OracleTextField(repName).Enter gb_TestDataDic.item ("repName_2")
    orcFlxWinObj.OracleTextField(ledgerField).OpenDialog
    orcFlxWinObj.OracleTextField(endDateField).Enter d
    orcFlxWinObj.Approve
    fn_Click orcSubReqObj.OracleButton(subBtn)
    If fn_exist (orcDecisionObj) Then
        strRequestNum = f_GetNumericValueFromString(orcDecisionObj.GetROProperty("message"))
        orcDecisionObj.Decline
    End If
    orcFRWindowObj.CloseWindow
    Call f_CheckRequestStatus(strRequestNum)
    strRequestNum = Empty
    blnResultFlag = True
    fn_CreateAccounting = blnResultFlag
End Function

Function fn_runReporting()
    fn_runReporting = False
    flag = fn_run_ExpenseReportProgram()
    If flag Then
        fn_switchResponsibility(gb_TestDataDic.item("Responsibility4"))
        Call fn_run_CreateAccounting()
        fn_runReporting = True
        
    Else
        fnReportEvent "Fail","Run reporting","Run Reporting failed",False
    End If
End Function

Function fn_searchExpenseReport(var_expNum)
On Error Resume Next
If fn_exist(OIEPgObj_ExpenseHome.Link(expSearchLink_xpath))=true Then
	fnReportEvent "Pass","Search Expense Report Link","Expense report Link is present",False
	fn_Click OIEPgObj_ExpenseHome.Link(expSearchLink_xpath)
	fn_Set OIEPgObj_ExpenseSearch.WebEdit(RepNumFld_xpath),var_expNum
	fn_Click OIEPgObj_ExpenseSearch.WebButton(searchGoBtn_xpath)
	var_status = OIEPgObj_ExpenseSearch.WebTable(searchResTbl_xpath).GetCellData(2,4)

	If var_status = ""  Then
		fnReportEvent "Fail","Search Expense","Expense report not found",False
	End If
	fn_searchExpenseReport = var_status
Else
	fnReportEvent "Fail","Search Expense Link Status","Expense report Link not found",False
	fn_searchExpenseReport=false	
End If
	
If err.number <> 0 Then
	fnReportEvent "Fail","Search Expense"," failed to search the Expense report",True
	fn_searchExpenseReport=false
End If
End Function

Function fn_setGeneralInfo()
On error resume next
fn_setGeneralInfo=false
	If fn_exist (OIEPgObj_CERGenInfo)=true Then
		fnReportEvent "Pass","Create Expense Report: General Information Window Status","Successfully loaded General Information Window",false
		fnSet_FieldName OIEPgObj_CERGenInfo.WebEdit(purpose_xpath),gb_TestDataDic.item("Purpose"),"Purpose"
		fn_WSSendKeys TAB 
		fnSet_FieldName OIEPgObj_CERGenInfo.WebEdit(projSrc_xpath),gb_TestDataDic.item("Project source"),"Project source"
		fnSet_FieldName OIEPgObj_CERGenInfo.WebEdit(projCode_xpath),gb_TestDataDic.item("Project Code"),"Project code"
		Wait 5
		Call fn_SelectWeblist(OIEPgObj_CERGenInfo.WebList(busiPur_xpath),gb_TestDataDic.item("Business Purpose"),"Business Purpose")	
		fn_Click OIEPgObj_CERGenInfo.WebButton(Next_xpath)	
	Else 
		fnReportEvent "Fail","Fail : General Information Window Status","Unable to load General Information Window",false  
		fn_setGeneralInfo=false 
		Exit Function
	End If       
	fn_setGeneralInfo=true
	If err.number <> 0 Then
		fnReportEvent "Fail","General Info Error"," Failed to input the general Info. Error is : "&Err.Description,True
		Exit function
	End If
End Function

Function fn_setCashExpenseDetails_MultiLine()
 d = fn_getSysdateFormat("DD-MMM-YYYY")
    On Error Resume Next

        expenseTypeCount = Split( gb_TestDataDic.item ("Expense Type"),"|")
        
        For i = 0 To UBound(expenseTypeCount)
          fnSet_FieldName OIEPgObj_CERCashExp.WebEdit("xpath:=//*[@id='N51:Date:"&i&"']"),d,"Current Date"
          fnSet_FieldName OIEPgObj_CERCashExp.WebEdit( "xpath:=//*[@id='N51:ReceiptCurrencyAmount:"&i&"']"),gb_TestDataDic.item ("Receipt Amount"),"Receipt Amount"
          call fn_SelectWeblist(OIEPgObj_CERCashExp.WebList("xpath:=//*[@id='N51:WebParameterId:"&i&"']"),arrResponsibility(iRespIndex),"Expense Type")
           fnSet_FieldName OIEPgObj_CERCashExp.WebEdit( "xpath:=//*[@id='N51:Justification:"&i&"']"),gb_TestDataDic.item ("Description"),"Description"  
        Next
End Function

Function fn_checkRCDetails()
On error resume next
'fn_checkRCDetails=false
resultflagctr = 0 
	If fn_exist(OIEPgObj_ExpAllocation)=true Then	
		fnReportEvent "Pass","Create Expense Report : Expense Allocations Window Status","Successfully loaded Expense Allocations Window",false
		varEntity = fn_GetROPropertyValueByPropName(OIEPgObj_ExpAllocation.WebEdit(entity_xpath),"value")
'		If OIEPgObj_ExpAllocation.WebEdit(entity_xpath).GetROProperty("disabled:=0")=true Then
'		fnReportEvent "Pass","Create Expense Report : Expense Allocations Window Status","Entity field is editable",false
			If varEntity = gb_TestDataDic.item("Entity") Then
				fnReportEvent "Pass","Entity","Entity is same as test data. Entity is : "&varEntity,False
			Else
				fnReportEvent "Fail","Entity","Entity does not match with Test data.Please check. Expected Entity : "&varEntity,True
				resultflagctr = resultflagctr+1
			End If
'		Else 
'		fnReportEvent "Fail","Create Expense Report : Expense Allocations Window Status","Entity field is not editable",false
'		End  If 
'		
		varRC = fn_GetROPropertyValueByPropName(OIEPgObj_ExpAllocation.WebEdit(RCfield_xpath),"value")
		If varRC = gb_TestDataDic.item("Responsiblity_Center") Then
			fnReportEvent "Pass","Responsiblity_Center","Responsiblity_Center is same as test data. Responsibility Center is : "&varRC,False		
		Else
			fnReportEvent "Fail","Responsiblity_Center","Responsiblity_Center does not match with Test data.Please check. Expected Responsibility Center is : "&varRC,True
			resultflagctr = resultflagctr+1
		End If
	
		varinterComp = fn_GetROPropertyValueByPropName(OIEPgObj_ExpAllocation.WebEdit(interComp_xpath),"value")
		If varinterComp = gb_TestDataDic.item("Inter_Company") Then
			fnReportEvent "Pass","Inter_Company","Inter_Company is same as test data. Inter Company is : "&varinterComp,False		
		Else
			fnReportEvent "Fail","Inter_Company","Inter_Company does not match with Test data.Please check. Expected Inter Company is : "&varinterComp,True
			resultflagctr = resultflagctr+1
		End If
	Else 
	fnReportEvent "Fail","Create Expense Report : Expense Allocations Window Status","Unable to load Expense Allocations Window",false
	End If
	
	If resultflagctr = 0 Then
		fn_checkRCDetails=true
	Else 
		fn_checkRCDetails=false
	End If

        If err.number <> 0 Then
        	fnReportEvent "Fail","validate RC details"," Failed to validate the RC Details. Error is : "&Err.description,True
        	fn_checkRCDetails=false
        	Exit Function
    	End If
	
End Function

