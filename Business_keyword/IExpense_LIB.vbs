Public OIEPgObj_ExpenseHome,OIEPgObj_GenPref,OIEPgObj_EmpBankDet,OIEPgObj_CERGenInfo,OIEPgObj_CERCashExp,OIEPgObj_CashExpDet,OIEPgObj_ExpAllocation,OIEPgObj_review
Set OIEPgObj_ExpenseHome = Browser("name:=Expense Home").Page("title:=Expense Home")
Set OIEPgObj_ExpenseSearch = Browser("name:=Expense Reports").Page("title:=Expense Reports")
Set OIEPgObj_GenPref = Browser("name:=General Preferences").Page("title:=General Preferences")
Set OIEPgObj_EmpBankDet = Browser("name:=Employee Bank Details Page").Page("title:=Employee Bank Details Page")
Set OIEPgObj_CERGenInfo = Browser("name:=Create Expense Report: General Information").Page("title:=Create Expense Report: General Information")
Set OIEPgObj_CERCashExp = Browser("name:=Create Expense Report: Cash and Other Expenses").Page("title:=Create Expense Report: Cash and Other Expenses")
Set OIEPgObj_CashExpDet = Browser("name:=Cash and Other Expenses: Details for Line 1").Page("title:=Cash and Other Expenses: Details for Line 1")
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
Const reciptVerifiedCB_xpath = "xpath:=//*[@id='N156:ReceiptVerified1:1']"
Const applyBtn_xpath = "xpath:=//*[@id='AuditActionButton']"
Const cnfText_xpath = "xpath:=//*[@id='FwkErrorBeanId']//DIV[3]"
Const noRows_xpath = "xpath:=//SPAN[@id='NoRowsPrompt']"
Const aprvrNotif_xpath = "xpath:=//*[@id='NtfWorklist']/TABLE[2]"
Const expSearchLink_xpath = "xpath:=//*[@id='OIE_EXPENSE_REPORT_SEARCH']"
Const RepNumFld_xpath = "xpath:=//*[@id='SearchReportNum']"
Const searchGoBtn_xpath = "xpath:=//button[contains(text(),'Go')]"
Const searchResTbl_xpath = "xpath:=//SPAN[@id='HistoryResultsTbl']/TABLE[2]"
Const newReqBtn = "description:=Submit a New Request"
Const ledgerField = "prompt:=Ledger"
Const endDateField = "prompt:=End Date"
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
    varEmpNum = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebElement(empnum_xpath),"innertext")
    If varEmpNum <> ""  Then
        fnReportEvent "Pass","Employee Number","Employee Number " & varEmpNum & " is present",False
    Else
        fnReportEvent "Fail","Employee Number","Employee Number doesnt Exist",True
    End If
    varCntryName = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebList(cntryName_xpath),"value")
    'print varCntryName
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
        fnReportEvent "Pass","Bank and Branch Code","Bank and Branch Code is same as test data",False
    Else
        fnSet_FieldName OIEPgObj_EmpBankDet.WebEdit(bankBranch_code), gb_TestDataDic.item("Bank_Branch_code"),"Bank and Branchdata:image/pjpeg;base64,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 Code"
    End If
    varAccNum = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebEdit(accNum_code),"value")
    If varAccNum = gb_TestDataDic.item("Account_Number") Then
        fnReportEvent "Pass","Account Number","Account Number is same as test data",False
    Else
        fnSet_FieldName OIEPgObj_EmpBankDet.WebEdit(accNum_code), gb_TestDataDic.item("Account_Number"),"Account Number"
    End If
    varBenfStatus = fn_GetROPropertyValueByPropName(OIEPgObj_EmpBankDet.WebList(benfStatus_code),"value")
    If varBenfStatus = gb_TestDataDic.item("Beneficiary_Status") Then
        fnReportEvent "Pass","Beneficiary Status","Beneficiary Status is same as test data",False
    Else
        fnSet_FieldName OIEPgObj_EmpBankDet.WebList(benfStatus_code), gb_TestDataDic.item("Beneficiary_Status"),"Beneficiary Status"
    End If
    OIEPgObj_EmpBankDet.WebButton(SaveBtn_xpath).Click
    
    If err.number <> 0 Then
        fnReportEvent "Fail","validate bank details"," Failed to set the bank details",True
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
    If fn_exist( OIEPgObj_ExpenseHome.Link(IEHomeLink_xpath)) = True  Then
        fnReportEvent "Pass", "Oracle Iexpense page navigation status","Successfully navigated to OIE Home page ",False
        
        fn_Click OIEPgObj_ExpenseHome.WebButton(createExpRep_xpath)
        fnSet_FieldName OIEPgObj_CERGenInfo.WebEdit(purpose_xpath),gb_TestDataDic.item ("Purpose"),"Purpose"
        
        fnSet_FieldName OIEPgObj_CERGenInfo.WebEdit(projSrc_xpath),gb_TestDataDic.item ("Project source"),"Project source"
        
        fnSet_FieldName OIEPgObj_CERGenInfo.WebEdit(projCode_xpath),gb_TestDataDic.item ("Project Code"),"Project code"
        
        Call fn_SelectWeblist(OIEPgObj_CERGenInfo.WebList(busiPur_xpath),gb_TestDataDic.item ("Business Purpose"),"Business Purpose")
        
        fn_Click OIEPgObj_CERGenInfo.WebButton(Next_xpath)
        fnSet_FieldName OIEPgObj_CERCashExp.WebEdit(date_xpath),d,"Current Date"
        
        fnSet_FieldName OIEPgObj_CERCashExp.WebEdit(recAmt_xpath),gb_TestDataDic.item ("Receipt Amount"),"Receipt Amount"
        
        Call fn_SelectWeblist(OIEPgObj_CERCashExp.WebList(expType_xpath), gb_TestDataDic.item ("Expense Type"),"Expense Type")
        fnSet_FieldName OIEPgObj_CERCashExp.WebEdit(desc_xpath),gb_TestDataDic.item ("Description"),"Description"
        
        fn_click OIEPgObj_CERCashExp.Image(dff_xpath)
        fnSet_FieldName OIEPgObj_CashExpDet.WebEdit(merchName_xpath),gb_TestDataDic.item ("Merchant_Name"),"Merchant_Name"
        
        fn_Click OIEPgObj_CashExpDet.WebCheckBox(missingRec_xpath)
        fnSet_FieldName OIEPgObj_CashExpDet.WebEdit(projCode2_xpath),gb_TestDataDic.item ("Project Code"),"Project code"
        fn_Click OIEPgObj_CashExpDet.WebButton(returnBtn_xpath)
        fn_Click OIEPgObj_CERCashExp.WebButton(Next1_xpath)
        fn_Click OIEPgObj_ExpAllocation.WebButton(Next2_btn)
        fn_Click OIEPgObj_review.WebCheckBox(attstn_xpath)
        fn_Click OIEPgObj_review.WebButton(submitBtn_xpath)
        fn_getExpenseReportConfirmation()
        Call fn_getApprover()'change name 
        blnResultFlag = True
    Else
        
    End If
    fn_cashExpenseClaim = blnResultFlag
End Function
Function fn_getExpenseReportConfirmation()
    On Error Resume Next
    cnfFlag = False
    strTest = "has been submitted."
    var_Expno = OIEPgObj_Confirmation.WebElement(expText_xpath).GetROProperty("innertext")
    ExpNo = Split(var_Expno," ")
    strQuery = "UPDATE [ExecutionResult$] SET Expense_report_No='" & ExpNo(3) & "' where TC_ID='" & gstrTestCaseExec_id & "'"
    Call fn_updateQuery(strQuery)
    strReqConfirmStmt = OIEPgObj_Confirmation.WebElement(expText_xpath).GetROProperty("innertext")
    If InStr(strReqConfirmStmt,strTest) > 0 Then
        print ("Confirmation found.Expense Report  " & ExpNo(3) & " has been submitted")
        fnReportEvent "Pass","Confirmation","Confirmation found.Expense Report " & ExpNo(3) & " has been submitted",True
        cnfFlag = True
        
    Else
        fnReportEvent "Fail","Confirmation","Confirmation not found",True
        cnfFlag = False
    End If
    fn_getExpenseReportConfirmation = cnfFlag
End Function
Function fn_getApprover()
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
End Function

Function fn_approveExpenseReport()
    blnResultFlag = False
    On Error Resume Next
    varCnfNo = fn_getExecutionResultData(gstrTestCaseExec_id,"Expense_report_No")
    
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
                    Exit For
                Else
                    fnReportEvent "Fail","Validate expense report- Approver Login","Approver failed to validate Expense report",True
                End If
            Else
            End If
        Next
    Else
        fnReportEvent "Fail","Page Does not exist","Expense HomePage was not found-Approver Login .",True
    End If
    
    fn_approveExpenseReport = blnResultFlag
    If err.number <> 0 Then
        fn_approveExpenseReport = False
        fnReportEvent "Fail","Approve Expense report failed"," failed to approve the Expense report",True
    End If
End Function

Function fn_validateReport()
    validateFlag = True
    fn_validateReport = validateFlag
End Function
Function fn_checkStatus()
    ' resultSearchCount=1
    blnResultFlag = False
    varCnfNo = fn_getExecutionResultData(gstrTestCaseExec_id,"Expense_report_No")
    print varCnfNo
    If  fn_Exist(OIEPgObj_ExpenseHome) = True Then
        
        'var_Row = fn_GetROPropertyValueByPropName (OIEPgObj_ExpenseHome.WebTable(statusTable_xpath),"rows")
        '        For i = 2 To var_Row
        '            var_subject = OIEPgObj_ExpenseHome.WebTable(statusTable_xpath).GetCellData(i,1)
        '            If varCnfNo = var_subject Then
        var_status = fn_searchExpenseReport(varCnfNo)'OIEPgObj_ExpenseHome.WebTable(statusTable_xpath).GetCellData(i,3)
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
            print "Check Status. Incorrect status Updated"
        End If
        '            End If
        '        Next
        
    Else
        fnReportEvent "Fail","Search Expense Number","Fail to search the Expense number &value is =" & varCnfNo,True
    End If
    
    fn_checkStatus = blnResultFlag
End Function
Function fn_sendReceipts()
    Dim objOutlook
    Dim objOutlookMsg
    Dim olMailItem
    ' Create the Outlook object and the new mail object.
    Set objOutlook = CreateObject("Outlook.Application")
    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
    ' Define mail recipients
    objOutlookMsg.To = "sample@email.com"
    objOutlookMsg.CC = "sample@email.com"
    objOutlookMsg.BCC = "sample@email.com"
    ' Define a file for attachment
    doc = "C:\temp\test.xls"
    ' Body of the message
    objOutlookMsg.Subject = "Subject sample"
    objOutlookMsg.Body = "This is a test"
    ' Add the attachment to the email
    objOutlookMsg.Attachments.Add(doc)
    ' Display the email
    objOutlookMsg.Display
    ' Send the message
    objOutlookMsg.Send
    ' Release the objects
    Set objOutlook = Nothing
    Set objOutlookMsg = Nothing
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
    fn_Click OIEPgObj_ExpenseHome.Link(expSearchLink_xpath)
    fn_Set OIEPgObj_ExpenseSearch.WebEdit(RepNumFld_xpath),var_expNum
    fn_Click OIEPgObj_ExpenseSearch.WebButton(searchGoBtn_xpath)
    var_status = OIEPgObj_ExpenseSearch.WebTable(searchResTbl_xpath).GetCellData(2,4)
    If var_status = ""  Then
        fnReportEvent "Fail","Search Expense","Expense report not found",False
    End If
    fn_searchExpenseReport = var_status
    If err.number <> 0 Then
        fnReportEvent "Fail","Search Expense"," failed to search the Expense report",True
    End If
End Function
