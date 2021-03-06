
Public orcTranscnWindowObj,orcLinesWindowObj,orcDistributionWindowObject,orcFormLineItemDescObj,orcRuleAccntObj,orcTransactionNoObj,orcDistTableObj,orcCreditTranscnWindowObj
Public orcReceiptWindowObj,orcRcptCustomerNum,orcRcptApplicationsWindowObj,orcReceiptSummaryWindowObj
Public orcRcptBatchWindowObj,orcRcptTitleObj,objRecSet
Public orcParametersWindowObj,orcRequestWindowObj,orcFindRequestObj,orcSubmitRequestObj,orcDecisionNotificationObj,orcSubmitNewRequestObj
Public orcAdjustmentsPage,orcFindAdjustmentsPage,ObjSubledgerJournalEntryPage ,orcInstallmentsForms,orcApprovalLimits,vstrAdjstNumber,orccleAppR12

'########## descriptive object identification variable #################
Set orcTranscnWindowObj = OracleFormWindow("title:=Transactions.*")
Set orcLinesWindowObj = OracleFormWindow("title:=Lines.*")
Set orcDistributionWindowObject = OracleFormWindow("title:=Distributions.*")
Set orcCreditTranscnWindowObj =  OracleFormWindow("title:=Credit Transactions.*")
Set orcReceiptWindowObj = OracleFormWindow("title:=Receipts.*")
Set orcLineItemDescObj = orcLinesWindowObj.OracleTabbedRegion("label:=Main").OracleTable("block name:=Table")
Set orcRuleAccntObj = orcLinesWindowObj.OracleTabbedRegion("label:=Rules").OracleTable("block name:=Table")
Set orcLineTransactionFlexField = OracleFlexWindow("title:=Line Transaction Flexfield")
Set orcDistTableObj = orcDistributionWindowObject.OracleTable("block name:=Table")
Set orcTransactionNoObj = OracleFormWindow("title:=Transactions.*").OracleTextField("description:=Number","tooltip:=Transaction Number")
Set orcCreditTransactions =   OracleFormWindow("title:=Credit Transactions.*").OracleTabbedRegion("label:=Transaction Amounts")
Set orcRcptCustomerNum = OracleFormWindow("title:=Receipts.*").OracleTabbedRegion("label:=Main").OracleTextField("description:=Number")
Set orcRcptApplicationsWindowObj =  OracleFormWindow("title:=Applications.*")
Set orcReceiptSummaryWindowObj = OracleFormWindow("title:=Receipts Summary.*")
Set orcRcptTitleObj = OracleFormWindow("short title:=Receipts")
Set orcRcptBatchWindowObj = OracleFormWindow("title:=Receipt Batches.*")
Set OracleNavigatorForm = OracleFormWindow("title:=Navigator.*").OracleTabbedRegion("label:=Functions")
Set orcHomePageNavigator = OracleFormWindow("title:=Navigator.*")
Set orcAdjustmentsPage = OracleFormWindow("title:=Adjustments.*")
Set orcFindAdjustmentsPage = OracleFormWindow("title:=Find Adjustments*")
Set ObjSubledgerJournalEntryPage = Browser("name:=Subledger Journal Entry.*").Page("title:=Subledger Journal Entry.*")
Set orcInstallmentsForms = OracleFormWindow("title:=Installments.*")
Set orcApprovalLimits = OracleFormWindow("title:=Approval Limits.*")

'Operational Reporting Case Objects 
Set orcParametersWindowObj = OracleFlexWindow("title:=Parameters")
Set orcRequestWindowObj = OracleFormWindow("title:=Requests")
Set orcFindRequestObj = OracleFormWindow("title:=Find Requests")
Set orcSubmitRequestObj = OracleFormWindow("title:=Submit Request")
Set orcDecisionNotificationObj = OracleNotification("title:=Decision")
Set orcSubmitNewRequestObj = OracleFormWindow("title:=Submit a New Request")

Set orccleAppR12 = Browser("title:=Oracle Applications R12").WinObject("text:=Do you want to open or save Create_Accounting.*")

'=============================================================
'*************************************************************************
'AR Oracle Form Objects
'=============================================================
'*************************************************************************
Const orcSource = "description:=Source"
Const orcClass = "description:=Class"
Const orcType = "description:=Type"
Const orcLineItems = "description:=Line Items"
Const orcCurrency = "description:=Currency"
Const orcPaymentTerm = "description:=Payment Term"
Const orcInvoicingRule = "description:=Invoicing Rule"
Const orcCancellationReason = "description:=Reason"
Const orcCreditLines = "description:=Credit Lines"
Const orcReceiptMethod = "description:=Receipt Method"
Const orcReceiptNumber = "description:=Receipt Number"
Const orcReceiptAmount = "description:=Net Receipt Amount.*"
Const orcUnappliedAmount = "description:=Unapplied"
Const orcAppliedAmount = "description:=Applied"
Const orcBatchSource = "description:=Batch Source"
Const orcPaymentMethod = "description:=Payment Method"
Const orcRcptBtn = "description:=Receipts"
Const orcOpenBtn = "description:=Open"
Const orcApplyBtn = "description:=Apply"
Const orcNetReceiptAmount = "description:=Net Receipt Amount"
Const intRecord_no = 1 
Const contextvalue = "prompt:=Context Value"
Const okbutton = "label:=OK"
Const distributions = "description:=Distributions"
Const popUpList = "class description:=popup list box"
Const creditMemomLP = "description:=Credit Memo: Line: Percent"
Const orcAdjustmentNum = "description:=Adjustment Number"
Const orcFind = "description:=Find"
Const orcApproveAdjst = "title:=Approve Adjustments"
Const billToName = "description:=Bill To: Name"
Const orctable ="block name:=Table"
Const orcRules = "label:=Rules"
Const orcMain = "label:=Main"
Const orcComment ="label:=Comments"
Const orcControlCount  = "description:=Totals: Control Count"
Const orcControlAmount = "description:=Totals: Control Amount"
Const orcCustNumber = "description:=Number"
Const orcRefreshData = "description:=Refresh Data"
Const orcSubmitNewRequest = "description:=Submit a New Request"
Const orcRequestName = "description:=Name"
Const orcOperatingUnit = "prompt:=Operating Unit"
Const orcSetOfBooks = "prompt:=Set Of Books"
Const orcSubmitBtn = "description:=Submit"
Const orcViewOutput = "description:=View Output"
Const orcTranscnDateLow = "prompt:=Transaction Date Low"
Const orcTranscnDateHigh = "prompt:=Transaction Date High"
Const orcSingleReq = "selected item:=Single Request"
Const orcAllReq = "selected item:=All My Requests"
Const orcGLDateFrom = "prompt:=GL Date From"
Const orcGLDateTo = "prompt:=GL Date To"
Const orcPrintFormat = "prompt:=Print Format"
Const orcAsOfDate = "prompt:=As of Date"
Const ledgerField = "prompt:=Ledger"
Const endDateField = "prompt:=End Date"
Const orcMode = "prompt:=Mode"
Const orcReport = "prompt:=Report"
Const orcTransferToGL = "prompt:=Transfer to General Ledger"
Const orcPostToGL = "prompt:=Post in General Ledger"
Const orcIncludeUTI = "prompt:=Include User Transaction Identifiers"

Const orcNote = "title:=Note"
Const orcSTReport = "short title:=Report"
Const OutputBrowserTitle = "title:=https://test.risebs.mmc.com.*"
Const OutputBrowserURL = "url:=https://test.risebs.mmc.com.*"
Const output_xpath = "xpath:=//PRE[1]"
Const orcReqID = "description:=Request ID"

Const winSave = "acc_name:=Save"
Const winOpen = "acc_name:=Open"

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  25/10/2021
'Name -                 fn_NavigatorOracle 
'description:         fn_NavigatorOracle :  Navigate/Login to application depending if application is open/closed
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_NavigatorOracle()
	On error Resume Next
	blnresult =false
	if OracleNavigatorForm.OracleList("description:=Function List").Exist(10) then
'		  validate if responsibilty is selected correctly 
	title =  trim(Replace(orcHomePageNavigator.GetROProperty("title"),"Navigator -",""))
		If   title    =Split(gb_TestDataDic(gstrTdIdentifer2),"|")(0)Then
			blnresult = fn_Navigator()
		else
			blnresult = fn_switchResponsibility()					
'			If OracleNavigatorForm.OracleList("description:=Function List").Exist(10) Then
'					
'		
'			else
'				blnresult = fn_switchResponsibility()					
'			End If		
			
		End If		 		
	else
 		call	fn_LoginSSO()
		call fn_NavigateMenu()
		blnresult=fn_NavigateResponsibility()
'		fn_NavigatorOracle= fn_Navigator()
	End If 
'	will return the final value to the function
	fn_NavigatorOracle =blnresult
	If err.number<>0 Then
		fn_NavigatorOracle =false
	End If
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  25/10/2021
'Name -                 fn_Navigator 
'description:         fn_Navigator :  Navigate to respective menu depending on responsibility used 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_Navigator()
On Error Resume Next
	fn_Navigator =false
'	fetching the test data identifier
		if gstrTdIdentifer2 = "Responsibility1" then
			vnavigator ="OracleNavigator1"
		elseif gstrTdIdentifer2 = "Responsibility2" then 		
			vnavigator = "OracleNavigator2"
		elseif gstrTdIdentifer2 = "Responsibility3" then 	
			vnavigator = "OracleNavigator3"
		 Else 
            		vnavigator = "OracleNavigator4"
        	End  If 
'Split the vnavigator column value based on the column separted delimeter
		arrNavigation = split(gb_TestDataDic.item(vnavigator),",")
		 For itr = 0 To Ubound( arrNavigation) 
			If itr = 0 Then
				pNav1 =split(gb_TestDataDic.item(vnavigator),",")(itr) 	
			ElseIf itr =1 Then 
				pNav2= split(gb_TestDataDic.item(vnavigator),",")(itr)	
			End If	 
		 Next
	
	Set OracleFormFuncationsList =OracleNavigatorForm.OracleList("description:=Function List")
	intcounter = 0 		
	fn_click OracleNavigatorForm.OracleButton("description:=Collapse All")
	fn_Highlight OracleFormFuncationsList
		Do					
'			navg1 = "+  "& pNav1			
			OracleFormFuncationsList.Select("+  "& pNav1)
			OracleFormFuncationsList.Activate("+  "& pNav1)
					
			If len(pNav2) >1Then	
				If lcase(pNav2) = Lcase("transactions-->transaction" ) Then
					OracleFormFuncationsList.Activate(4)
				ElseIf lcase(pNav2) = lcase("transactionsbatchsummary") then 
					OracleFormFuncationsList.Activate(pNav2)	
				ElseIf lcase(pNav2) = lcase("Receipts-->Receipts") Then
					OracleFormFuncationsList.Activate(4)
				ElseIf lcase(pNav2) = lcase("Receipts-->Batches") Then
					OracleFormFuncationsList.Activate(2)	
				ElseIf lcase(pNav2) = lcase("Requests-->View") Then
					OracleFormFuncationsList.Activate(9)
				ElseIf lcase(pNav2) = lcase("Requests-->Run") Then
					OracleFormFuncationsList.Activate(8)
				ElseIf lcase(pNav2) = lcase("Requests-->InterfaceMaintenance_Run") Then
					OracleFormFuncationsList.Activate(2)
				ElseIf lcase(pNav2) = lcase("Transactions-->Approval Limits") Then
		                    OracleFormFuncationsList.Activate(7)		                		         
		                    navg3 = split(pNav2,"-->")(1)
		                    OracleFormFuncationsList.Select(navg3)
		                    OracleFormFuncationsList.Activate(navg3)
				     'Added for AP TC GSI.P2P.AP.SA.025
	                    	ElseIf lcase(pNav2) = lcase("Requests-->APInterfaceMaintenance_Run") Then
		                    OracleFormFuncationsList.Activate(5)
				Else 
					OracleFormFuncationsList.Activate(pNav2)		
				End If
			End If	
				intcounter = intcounter +1	
		Loop until orcTranscnWindowObj.OracleTextField(orcSource).Exist=false or counter<=1
	If intcounter >=1  Then
		fn_Navigator =true
	End If
	
	If err.number<>0 Then
		fn_Navigator =false
		fnReportEvent "Fail","Oracle Navigator Page","Function name  : fn_Navigator , Fail to to navigate on the jav form and Error is : " &Err.description,true 
	End If
	
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  15/10/2021
'Name -                 fn_createRcpt_Refund_WriteOff 
'description:         fn_createRcpt_Refund_WriteOff :  Receipt Processing -  Create Refund/WriteOff
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_createRcpt_Refund_WriteOff()			'Use Responsibility - AR Receipts Processor
On error resume next
fn_createRcpt_Refund_WriteOff = false
If fn_exist(orcReceiptWindowObj) = true Then
	fnReportEvent "Pass","Receipt Window Status","Successfully loaded Oracle Receipt Form Window",false   
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptMethod),gb_TestDataDic.item("Receipt_Method"),"Receipt Method"
'	vreceiptnumber = "REC" & fn_RandomNumber(4)
	vreceiptnumber = fn_Create_CaptureReceiptNo()
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptNumber),vreceiptnumber,"Receipt Number"
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcNetReceiptAmount),gb_TestDataDic.item("Receipt_Amount"),"Net Receipt Amount"
	fn_ReportEnter orcRcptCustomerNum,gb_TestDataDic.item("Customer_Number"),"Customer Number"
	fn_Click orcReceiptWindowObj.OracleButton(orcApplyBtn)	
	fn_exist orcRcptApplicationsWindowObj.OracleTable(orctable)
	If orcRcptApplicationsWindowObj.Exist(10) Then
		fn_EnterField orcRcptApplicationsWindowObj.OracleTable(orctable),intRecord_no,"Apply To",gb_TestDataDic.item("Apply_To_Field"),"Apply To"
		fn_EnterField orcRcptApplicationsWindowObj.OracleTable(orctable),intRecord_no,"Amount Applied",gb_TestDataDic.item("Amount_Applied"),"Amount Applied"	

		If OracleFormWindow("title:=Applications.*").OracleButton("label:=Refund Attributes").Exist(10) Then
			fn_Click OracleFormWindow("title:=Applications.*").OracleButton("label:=Refund Attributes")
			fn_ReportEnter OracleFormWindow("title:=Refund Attributes").OracleTextField("description:=Refund Payment Method"),gb_TestDataDic.item("Payment_Method"),"Payment Method"
			fn_Click OracleFormWindow("title:=Refund Attributes").OracleButton("description:=Apply")
		Else 
			fn_EnterField orcRcptApplicationsWindowObj.OracleTable(orctable),intRecord_no,"Activity",gb_TestDataDic.item("Activity"),"Activity"
			fn_Click OracleFlexWindow("title:=Receipt Application Information").OracleButton("label:=OK")
			fn_Click OracleFlexWindow("title:=Additional information").OracleButton("label:=OK")
		End If
	Else
		fnReportEvent "Fail","Applications Window Status","Unable to Load Oracle Applications Form Window",true   
		fn_createRcpt_Refund_WriteOff=false
		Exit function
	End  If 
	
	fn_CloseWindow orcRcptApplicationsWindowObj
	fn_SelectMenu orcReceiptWindowObj,"filesave"
	
	vstrUnappliedAmount = orcReceiptWindowObj.OracleTextField(orcUnappliedAmount).GetROProperty("value")
	ExpUnappliedAmount = gb_TestDataDic.item("Receipt_Amount")-gb_TestDataDic.item("Amount_Applied")
	
	If Cint(vstrUnappliedAmount) = Cint(ExpUnappliedAmount) Then
		 fnReportEvent "Pass","Balances : UnApplied Amount Check","UnApplied amount is displayed as expected. Amount is = "&ExpUnappliedAmount,false
		 fn_createRcpt_Refund_WriteOff=true  
	Else 
	  	fnReportEvent "Fail","Balances : UnApplied Amount Check","UnApplied amount is not displayed is as expected. Expected Amount is = "&ExpUnappliedAmount,true      
		fn_createRcpt_Refund_WriteOff=false                        
	End If
	
	vstrAppliedAmount = orcReceiptWindowObj.OracleTextField(orcAppliedAmount).GetROProperty("value")
	ExpAppliedAmount = gb_TestDataDic.item("Amount_Applied")
	
	If Cint(vstrAppliedAmount) = Cint(ExpAppliedAmount)  Then
		fnReportEvent "Pass","Balances : Applied Amount Check","Applied amount is displayed as expected. Amount is = "&ExpAppliedAmount,false
		fn_createRcpt_Refund_WriteOff=true  
	Else 
		fnReportEvent "Fail","Balances : Applied Amount Check","Applied amount is not displayed is as expected. Expected Amount is = "&ExpAppliedAmount,true                 
	End If
Else 	
		fnReportEvent "Fail","Receipt Window Status","Unable to Load Oracle Receipt Form Window",true   
		Exit function
End If

fn_CloseWindow orcReceiptWindowObj

If Err.number <> 0 Then             
	'print Err.description,true
	fnReportEvent "Fail","Refund/Write-Off","Function name  : fn_createRcpt_Refund_WriteOff , Fail to Create Refund/Write-Off. Error is : " &Err.description,true 
	Exit function
End If

End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  14/10/2021
'Name -                 fn_ApplyReceiptByBatch 
'description:         fn_ApplyReceiptByBatch :  Receipt Proccessing - Apply Receipt using Batch 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_ApplyReceiptByBatch()				'Responsibility Used - AR Receipts Processor 
On error resume next
fn_ApplyReceiptByBatch = false
vtransactionnumber = fn_getExecutionResultData(gstrTestCaseExec_id,"Transaction_No")

If vtransactionnumber=0 or vtransactionnumber="" or Isnull(vtransactionnumber) Then
	fnReportEvent "Fail","Transaction Number","Failed to fetch the Transaction Number for : GSI.O2C.AR.SA.010",false
End If

If fn_exist (orcRcptBatchWindowObj) = true Then
	fnReportEvent "Pass","Receipt by Batch Window Status","Successfully loaded Receipt by Batch Form Window",false   
	fn_ReportEnter orcRcptBatchWindowObj.OracleTextField(orcBatchSource),gb_TestDataDic.item("Batch_Source"),"Batch Source"	
	fn_ReportEnter orcRcptBatchWindowObj.OracleTextField(orcPaymentMethod),gb_TestDataDic.item("Payment_Method_Receipt"),"Payment Method"
'	fn_ReportEnter orcRcptBatchWindowObj.OracleTextField(orcControlCount),gb_TestDataDic.item("Total_Count"),"Total Count"
'	fn_ReportEnter orcRcptBatchWindowObj.OracleTextField(orcControlAmount),gb_TestDataDic.item("Total_Amount"),"Total Amount"	
	fn_ReportEnter orcRcptBatchWindowObj.OracleTextField(orcControlCount),gb_TestDataDic.item("Quantity_Value"),"Total Count"
	fn_ReportEnter orcRcptBatchWindowObj.OracleTextField(orcControlAmount),gb_TestDataDic.item("Unit_Price"),"Total Amount"	
	fn_Click orcRcptBatchWindowObj.OracleButton(orcRcptBtn)
	
	If orcReceiptSummaryWindowObj.Exist(5) Then
'		vreceiptnumber = "REC" & fn_RandomNumber(4)
		vreceiptnumber = fn_Create_CaptureReceiptNo()
		fn_EnterField orcReceiptSummaryWindowObj.OracleTable(orctable),intRecord_no,"Receipt Number",vreceiptnumber,"Receipt No"
		fn_EnterField  orcReceiptSummaryWindowObj.OracleTable(orctable),intRecord_no,"Net Amount",gb_TestDataDic.item("Net_Amount"),"Net Amount"
		fn_Click orcReceiptSummaryWindowObj.OracleButton(orcOpenBtn)
	else	
		fnReportEvent "Fail","Receipt Summary Window: Fail", "Fail to Navigate to the receipt summary window",true
		fn_ApplyReceiptByBatch=false
		Exit function
	End If
	If orcRcptTitleObj.Exist(5) Then
		fn_Enter orcRcptTitleObj.OracleTabbedRegion(orcMain).OracleTextField(orcCustNumber),gb_TestDataDic.item("Customer_Number")
		fn_Click orcRcptTitleObj.OracleButton(orcApplyBtn)
	else	
		fnReportEvent "Fail","Receipt customer information Window: Fail", "Fail to Navigate to the receipt customner information window",true
		fn_ApplyReceiptByBatch=false
		Exit function
	End If
	If orcRcptApplicationsWindowObj.Exist(5) Then
		fn_EnterField orcRcptApplicationsWindowObj.OracleTable(orctable),intRecord_no,"Apply To",vtransactionnumber,"Apply To:Transcn No"
		fn_SelectMenu orcRcptApplicationsWindowObj,"filesave"
	else	
		fnReportEvent "Fail","Receipt Application Window: Fail", "Fail to Navigate to the Receipt Application window",true
		fn_ApplyReceiptByBatch=false
		Exit function
	End If	
	fn_CloseWindow orcRcptApplicationsWindowObj
	fn_CloseWindow orcRcptTitleObj
	fn_CloseWindow orcReceiptSummaryWindowObj
	
	vstrActualAmountCheck = OracleFormWindow("title:=Receipt Batches.*").OracleTextField("description:=Totals: Actual Amount").GetROProperty("value")
	ExpActualAmount = gb_TestDataDic.item("Net_Amount")
	
	If Cint(vstrActualAmountCheck) = Cint(ExpActualAmount) Then
		fnReportEvent "Pass","Actual Amount Check","Actual Amount & Receipt Amount are matching and Amount is = "&ExpActualAmount,true
		fn_ApplyReceiptByBatch=true
	Else 
		fnReportEvent "Fail","Actual Amount Check : Fail", "Actual Amount & Receipt Amount are not matching. Expected Amount is = "&ExpActualAmount,true
	End If	
	
	vstrDifferenceAmountCheck = OracleFormWindow("title:=Receipt Batches.*").OracleTextField("description:=Totals: Difference Amount").GetROProperty("value")
	ExpDifferenceAmount = gb_TestDataDic.item("Total_Amount") - gb_TestDataDic.item("Net_Amount")
	
	If Cint(vstrDifferenceAmountCheck) = Cint(ExpDifferenceAmount)  Then
		fnReportEvent "Pass","Difference Amount Check","Difference Amount is displayed as expected & Amount is = "&ExpDifferenceAmount,true
		fn_ApplyReceiptByBatch=true
	Else 
		fnReportEvent "Fail","Difference Amount Check : Fail", "Difference Amount is not displayed as expected. Expected Amount is = "&ExpDifferenceAmount,true
	End If	
Else 	
	fnReportEvent "Fail","Receipt by Batch Window Status","Unable to Load Receipt by Batch Form Window",true   
	Exit function
End If	

fn_CloseWindow orcRcptBatchWindowObj

If Err.number <> 0 Then             
 	'print Err.description,true
 	fnReportEvent "Fail","Apply Receipt","Function name  : fn_ApplyReceiptByBatch , Fail to Apply Receipt by Batch. Error is : " &Err.description,true 
    	Exit function
End If
End Function


'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  05/10/2021
'Name -                 fn_createReceipt 
'description:         fn_createReceipt :  Receipt Proccessing - Enter a receipt against an open invoice
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_createReceipt()		'AR Receipt Processor Responsibility
On error resume next 
fn_createReceipt = false
vtransactionnumber = fn_getExecutionResultData(gstrTestCaseExec_id,"Transaction_No")

If vtransactionnumber=0 or vtransactionnumber="" or Isnull(vtransactionnumber) Then
	fnReportEvent "Fail","Transaction Number","Failed to fetch the Transaction Number for : GSI.O2C.AR.SA.009",false
End If

If fn_exist(orcReceiptWindowObj) Then
	fnReportEvent "Pass","Receipt Window Status","Successfully loaded Oracle Receipt Form Window",false
'	vreceiptnumber = "REC" & fn_RandomNumber(4)	
	vreceiptnumber = fn_Create_CaptureReceiptNo()
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptMethod),gb_TestDataDic.item("Receipt_Method"),"Receipt Method"
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptNumber),vreceiptnumber,"Receipt Number"	
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptAmount),gb_TestDataDic.item("Receipt_Amount"),"Receipt Amount"	
	fn_ReportEnter orcRcptCustomerNum,gb_TestDataDic.item("Customer_Number"),"Customer Number"	
	fn_Click OracleFormWindow("title:=Receipts.*").OracleButton("description:=Apply")	
	
	If fn_exist(orcRcptApplicationsWindowObj) = True Then		
'		If   Instr(1,orcRcptApplicationsWindowObj.GetROProperty("title"),"Applications") > 1   Then	
			fnReportEvent "Pass","Receipt Application Form","Succesfully Navigate to the Receipt Application Form",false
			fn_EnterField orcRcptApplicationsWindowObj.OracleTable(orctable),intRecord_no,"Apply To",vtransactionnumber,"Apply To"		
			fn_WSSendKeys("TAB")
			fn_SelectMenu orcRcptApplicationsWindowObj,"filesave"
	Else 	
			fnReportEvent "Fail","Receipt Application Form","Fail to  Navigate to the Receipt Application Form",false	
			Exit function
	End If	
	fn_CloseWindow orcRcptApplicationsWindowObj	
	
	vAppliedAmount = orcReceiptWindowObj.OracleTextField(orcAppliedAmount).GetROProperty("value")
	vUnappliedAmount = orcReceiptWindowObj.OracleTextField(orcUnappliedAmount).GetROProperty("value")
	
	ExpAppliedTotalamount = (gb_TestDataDic.item("Quantity_Value")* gb_TestDataDic.item("Unit_Price"))

	If Cint(vAppliedAmount) = Cint(ExpAppliedTotalamount) Then
	         fnReportEvent "Pass","Applied Amount Check","Applied amount is displayed as expected and amount value is: = " &ExpAppliedTotalamount  ,false
	         fn_createReceipt=true
	Else 
	         fnReportEvent "Fail","Applied Amount Check","Fail to validate Applied amount and expected amount  value should be : =" &ExpAppliedTotalamount,true       	         
	End If
	
	ExpUnappliedTotalamount = (gb_TestDataDic.item("Receipt_Amount") - cint(vAppliedAmount))
	If Cint(vUnappliedAmount) =Cint(vUnappliedAmount) Then
	         fnReportEvent "Pass","Unapplied  Amount Check","UnApplied amount is displayed as expected and amount value is: = "&ExpUnappliedTotalamount,false
	         fn_createReceipt=true
	Else 
	         fnReportEvent "Fail","Unapplied  Amount Check","Fail to validate UnApplied Amount. Expected value = " &ExpUnappliedTotalamount,true       	         
	End If
Else 	
	fnReportEvent "Fail","Receipt Window Status","Unable to Load Oracle Receipt Form Window",true   
	Exit function
End If

fn_CloseWindow orcReceiptWindowObj

If Err.number <> 0 Then             
	'print Err.description,true
	fnReportEvent "Fail","Create Receipt","Function name  : fn_createReceipt , Fail to create Receipt. Error is : " &Err.description,true 	
	Exit function
End If

End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/10/2021
'Name -                 fn_CreateStdInv_CreditMemo 
'description:         fn_CreateStdInv_CreditMemo :  Invoice Proccessing - Std Invoice Creation, Non-AGIS Foreign Currency Invoice & Credit Memo Creation
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_CreateStdInv_CreditMemo()
On error resume next

if fn_exist(orcTranscnWindowObj.OracleTextField(orcSource) )= true Then
	fnReportEvent "Pass","Transaction Window Status","Successfully loaded Transaction Form Window",false 
	fn_ReportEnter orcTranscnWindowObj.OracleTextField(orcSource),gb_TestDataDic.item("Source_Field"),"Source Field"	
	fn_Select orcTranscnWindowObj.OracleList(orcClass),gb_TestDataDic.item("Class"),"Class"
	fn_ReportEnter orcTranscnWindowObj.OracleTextField(orcCurrency),gb_TestDataDic.item("Currency"),"Currency"
	fn_ReportEnter orcTranscnWindowObj.OracleTextField(orcType),gb_TestDataDic.item("Type"),"Type"
	fn_Enter orcTranscnWindowObj.OracleTabbedRegion(orcMain).OracleTextField(billToName),gb_TestDataDic.item("Bill_To")
	fn_WSSendKeys TAB
	
	fn_Enter orcTranscnWindowObj.OracleTabbedRegion(orcMain).OracleTextField(orcPaymentTerm),gb_TestDataDic.item("Payment_Term"),"Payment Term"
	fn_Select orcTranscnWindowObj.OracleTabbedRegion(orcMain).OracleList(orcInvoicingRule),gb_TestDataDic.item("Invoice_Rule"),"Invoice Rule"

	fn_Click orcTranscnWindowObj.OracleButton(orcLineItems)
	fn_Click OracleNotification("title:=Error").OracleButton("label:=OK")
	fn_EnterField orcLineItemDescObj,intRecord_no,"Description",gb_TestDataDic.item("Description_value"),"Description"
	fn_EnterField orcLineItemDescObj,intRecord_no,"Quantity",gb_TestDataDic.item("Quantity_Value"),"Quantity"
	fn_EnterField orcLineItemDescObj,intRecord_no,"Unit Price",gb_TestDataDic.item("Unit_Price"),"Unit Price"
	fn_EnterField orcLineItemDescObj,intRecord_no,"Tax Classification",gb_TestDataDic.item("Tax_Classification"),"Tax Classification"
	fn_Enter orcLineTransactionFlexField.OracleTextField(contextvalue) ,gb_TestDataDic.item("Operating_Unit")
	fn_Click orcLineTransactionFlexField.OracleButton(okbutton)
	fn_WSSendKeys TAB3
	fn_EnterField orcLinesWindowObj.OracleTabbedRegion(orcRules).OracleTable(orctable), intRecord_no,"Accounting",gb_TestDataDic.item("Accounting"),"Accounting"
	fn_Click orcLinesWindowObj.OracleButton(distributions)	
	fn_Select orcDistributionWindowObject.OracleList(popUpList),gb_TestDataDic.item("GLAccountSelect"),"GL Account Selection"
	
	Call fn_selectDistributions

	fn_CloseWindow orcDistWindowObj
	fn_CloseWindow orcLinesWindowObj

	call fn_SelectMenu(orcTranscnWindowObj,"FileSave")
	Call fn_CaptureTransactionNo
	
	fn_CloseWindow orcTranscnWindowObj
	fn_CreateStdInv_CreditMemo=true
Else 	
	fnReportEvent "Fail","Transaction Window Status","Unable to Load Transaction Form Window",true   
	fn_CreateStdInv_CreditMemo=false
	Exit function
End If

If Err.number <> 0 Then             
	'print Err.description,true
	fnReportEvent "Fail","Create Standard Invoice/Credit Memo","Function name  : fn_CreateStdInv_CreditMemo , Fail to create Std Invoice/Credit Memo. Error is : " &Err.description,true 
	fn_CreateStdInv_CreditMemo=false
	Exit function
End If

End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/10/2021
'Name -                 fn_selectDistributions 
'description:         fn_selectDistributions : Selection of  GL Accounts
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_selectDistributions()      
On error resume next
If fn_exist(orcDistTableObj)=true Then	
	rowCount = orcDistTableObj.GetROProperty("total rows")
	For i = 1 To rowCount 
	className = orcDistTableObj.GetFieldValue( i, "Class" )
		If className <> "" Then
			orcDistTableObj.OpenDialog i,"GL Account"
			OracleFlexWindow("title:=MMC CORPORATE FLEXFIELD").ShowCombinations
			OracleFlexWindow("title:=Enter Reduction Criteria.*").Approve
			If OracleListOfValues("title:=MMC CORPORATE FLEXFIELD").Exist(2) Then
				OracleListOfValues("title:=MMC CORPORATE FLEXFIELD").OracleButton("label:=OK").Click
			End If    
			glAccountInfo = orcDistTableObj.GetFieldValue( i, "GL Account" )
			fnReportEvent "Pass","Displaying distribution list", "Successfully selected the account=" &glAccountInfo& " For=" &className, False
'			If glAccountInfo="" Then
'				fnReportEvent "Fail","GL Account Status","GL Account field is empty",true   
'			End If
		Else 
	Exit For
		End If
	Next
		
Else 
	fnReportEvent "Fail","Distributions Table Window Status","Unable to Load Distributions Table Window",true   
End If

If Err.number <> 0 Then             
	'print Err.description,true
	fnReportEvent "Fail","Select GL Account","Function name  : fn_selectDistributions , Fail to select GL account. Error is : " &Err.description,true   
	Exit function
End If
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  06/10/2021
'Name -                 fn_CaptureTransactionNo 
'description:         fn_CaptureTransactionNo : Capture generated Transaction No & Save it is Execution Result Tab of Test Data Sheet
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_CaptureTransactionNo()
On error resume next 
If fn_exist(orcTransactionNoObj) Then	
	var_TransNo=orcTransactionNoObj.GetROProperty("value")		'Capture Transaction No Field in Oracle Form
	strQuery="UPDATE [ExecutionResult$] SET Transaction_No='"&var_TransNo&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"
	
	Call fn_updateQuery(strQuery)

	If var_TransNo <> "" Then
		fnReportEvent "Pass","Transaction No","Transaction No  "& var_TransNo &" has been generated",True
		fn_CaptureTransactionNo=true
	else
		fnReportEvent "Fail","Transaction No","Transaction No is not generated",True
	    	fn_CaptureTransactionNo=false
	End If
Else 
	fnReportEvent "Fail","Transaction Table Window Status","Unable to Load Transaction Table Window",true   
End If

If Err.number <> 0 Then             
	print Err.description,true
	fnReportEvent "Fail","Capture Transaction No","Function name  : fn_CaptureTransactionNo , Failed to capture transaction. Error is : " &Err.description,true   
	Exit function
End If
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  26/10/2021
'Name -                 fn_switchResponsibility 
'description:         fn_switchResponsibility : SwitchResponsibility
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_switchResponsibility()
On error resume next
blnresult = false    

	if gstrTdIdentifer2 = "Responsibility1" then
		respName =Split( gb_TestDataDic.item("Responsibility1"),"|")(0)
	elseif gstrTdIdentifer2 = "Responsibility2" then 		
		respName = Split( gb_TestDataDic.item("Responsibility2"),"|")(0)
	elseif gstrTdIdentifer2 = "Responsibility3" then 	
		respName = Split( gb_TestDataDic.item("Responsibility3"),"|")(0)
	End  If 

Set oraRespoForm = OracleListOfValues("title:=Responsibilities")

    If (fn_exist (orcHomePageNavigator)) Then
'        OracleFormWindow("title:=Navigator.*").SelectMenu "File->Switch Responsibility..." 	
        Call fn_SelectMenu(orcHomePageNavigator,"SwitchResponsibility")
	If oraRespoForm.Exist(5) Then
		oraRespoForm.Select respName	 	      
         Else 
        	fnReportEvent "Fail", "Switch Responsibility","Fail to naivigate to Switch Responsibility",true        	
    	End If    
    End  if 	
'    validate if responsibilty is selected correctly 
'Added the below code for AR Interface Maintaincace to close Submit New Request window
	If orcSubmitNewRequestObj.Exist(3) Then
		orcSubmitNewRequestObj.CloseWindow
	End If

title =  orcHomePageNavigator.GetROProperty("title")
	If Instr(1,title,respName) > 1 Then
		fnReportEvent "Pass", "Navigator Page Status","Navigator Page is displaying and User is able to switch the Responsibility to "&respName,false
		blnresult = true
	else
		fnReportEvent "Fail", "Navigator Page Status","Navigator Page is not displaying or Responsibility is not present for that user "& respName ,true		
	End If

blnresult = fn_Navigator()
fn_switchResponsibility =blnresult
'if switch responsibilty page exit then will close that form
'If oraRespoForm.Exist(1) Then
'		oraRespoForm.Cancel	
'End  If		
       If Err.number <> 0 Then             
              fnReportEvent "Fail", "Navigator Page Status","Navigator Page is not displaying or Responsibility is not present for that user "& respName ,true
              fn_switchResponsibility = false             	
        End If
End function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  07/10/2021
'Name -                 fn_ApproveTransaction 
'description:         fn_ApproveTransaction : Approve the generated Transaction No & Save it is Execution Result Tab of Test Data Sheet
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************

Function fn_ApproveTransaction()				'Responsibility Used - AR Transaction Approver 
On error resume next
	If fn_Exist(orcTranscnWindowObj) = true Then
		fnReportEvent "Pass","Transaction Window Status","Successfully loaded Transaction Form Window",false 
		OracleFormWindow("title:=Transactions.*").SelectMenu "View->Query By Example->Enter"
		fn_Exist OracleFormWindow("title:=Transactions.*").OracleTextField("description:=Source")
		fn_ReportEnter OracleFormWindow("title:=Transactions.*").OracleTextField("description:=Source"),gb_TestDataDic.item("Source_Field"),"Source Field"
	
		vtransactionnumber = fn_getExecutionResultData(gstrTestCaseExec_id,"Transaction_No")
		
		fn_ReportEnter OracleFormWindow("title:=Transactions.*").OracleTextField("description:=Number","tooltip:=Transaction Number"),vtransactionnumber,"Transaction No"
		OracleFormWindow("title:=Transactions.*").SelectMenu "View->Query By Example->Run"

		fn_Click OracleFormWindow("title:=Transactions.*").OracleButton("description:=Complete") 		
		vstrStatusCheck = OracleFormWindow("title:=Transactions.*").OracleButton("description:=Incomplete").GetROProperty("description")
		
		If vstrStatusCheck="Incomplete" Then
			fnReportEvent "Pass","Transaction Approval Status", "Transaction is completed/approved successfully",false
			fn_ApproveTransaction=true
		Else 
			fnReportEvent "Fail","Transaction Approval Status", "Unable to complete/approve Transaction successfully",true
			fn_ApproveTransaction=false
		End If
		
		Call fn_CaptureTransactionNo()
		
		Call fn_CloseWindow(orcTranscnWindowObj)
	Else 	
		fnReportEvent "Fail","Transaction Window Status","Unable to Load Transaction Form Window",true   
		fn_ApproveTransaction=false
		Exit function
	End If

If Err.number <> 0 Then             
	fnReportEvent "Fail","Approve Transaction","Function name  : fn_ApproveTransaction , Fail to approve transaction. Error is : " &Err.description,true   
	fn_ApproveTransaction=false
	Exit function
End If
End Function


'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  08/10/2021
'Name -                 fn_CreditNote 
'description:         fn_CreditNote : Invoice Processing - Create Credit Note against an Invoice
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_CreditNote()				'Responsibility Used - AR Transaction Processor 
On error resume next
	fn_CreditNote=false
	If fn_exist(orcTranscnWindowObj) = true Then
		fnReportEvent "Pass","Transaction Window Status","Successfully loaded Transaction Form Window",false 
		orcTranscnWindowObj.SelectMenu "View->Query By Example->Enter"	
		vtransactionnumber = fn_getExecutionResultData(gstrTestCaseExec_id,"Transaction_No")
		fn_ReportEnter orcTransactionNoObj,vtransactionnumber,"Transaction No"
		orcTranscnWindowObj.SelectMenu "View->Query By Example->Run"
		fn_SelectMenu orcTranscnWindowObj,"actionscredit"
		fn_ReportEnter orcCreditTranscnWindowObj.OracleTextField(orcCancellationReason),gb_TestDataDic.item("Cancellation_Reason"),"Cancellation Reason"
		fn_Select orcCreditTransactions.OracleList(popUpList), gb_TestDataDic.item("Credit_Allocation"),"Credit Allocation"
		fn_Enter orcCreditTransactions.OracleTextField(creditMemomLP),gb_TestDataDic.item("Percentage")
		fn_Click orcCreditTranscnWindowObj.OracleButton(orcCreditLines)
		fn_Click orcLinesWindowObj.OracleButton(distributions)	
		fn_exist orcDistributionWindowObject.OracleList(popUpList)
		fn_Select orcDistributionWindowObject.OracleList(popUpList),gb_TestDataDic.item("GLAccountSelect"),"Accounts For All Lines"
		
		vstrReceiveableAmountCheck = orcDistTableObj.GetFieldValue (1,"Distribution Amount")
		If fn_exist(vstrReceiveableAmountCheck)=true Then
			fnReportEvent "Pass","Receiveable Amount Field","Receiveable Amount field exists",false
			If vstrReceiveableAmountCheck <= 0  Then
				fnReportEvent "Pass","Receiveable Amount Check","Amount is Negative",false
				fn_CreditNote=true
			Else 
				fnReportEvent "Fail","Receiveable Amount Check", "Amount is not negative",true
			End If
		Else 
			fnReportEvent "Fail","Receiveable Amount Field","Receiveable Amount field does not exists",true
		End If
			
		vstrRevenueAmountCheck = orcDistTableObj.GetFieldValue (2,"Distribution Amount")
		If fn_exist(vstrReceiveableAmountCheck)=true Then
			fnReportEvent "Pass","Revenue Amount Field","Revenue Amount field exists",false
			If vstrReceiveableAmountCheck <= 0  Then
				fnReportEvent "Pass","Revenue Amount Check","Amount is Negative",false
				fn_CreditNote=true
			Else 
				fnReportEvent "Fail","Revenue Amount Check", "Amount is not negative",true
			End If
		Else 
			fnReportEvent "Fail","Revenue Amount Field","Revenue Amount field does not exists",false
		End If
		
	Else 	
		fnReportEvent "Fail","Transaction Window Status","Unable to Load Transaction Form Window",true   
		Exit function
	End If

fn_CloseWindow orcDistributionWindowObject
fn_CloseWindow orcLinesWindowObj
fn_CloseWindow orcCreditTranscnWindowObj
fn_SelectMenu orcTranscnWindowObj,"filesave"
fn_CloseWindow orcTranscnWindowObj

If Err.number <> 0 Then             
	fnReportEvent "Fail","CreditNote Creation","Functiona name  : fn_CreditNote , Fail to create the credit note.Error is : " &Err.description,true   
'	fn_CreditNote=false
	Exit function
End If
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  28/10/2021
'Name -                 fn_SelectMenu 
'description:         fn_SelectMenu : Menu Selection of different Oracle menu options
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_SelectMenu(objMenu,MenuNaviagation)
	On error resume Next
	
	Select Case Lcase(MenuNaviagation)
		Case "filesave"
			objMenu.SelectMenu "File->Save"
		Case "actionscredit"
			objMenu.SelectMenu "Actions->Credit"
		Case "switchresponsibility"
			objMenu.SelectMenu  "File->Switch Responsibility..." 	
		Case "actionsadjust"
            		objMenu.SelectMenu "Actions->Adjust"
	        Case "fileopen"
	            objMenu.SelectMenu "File->Open"
	        Case "viewquerybyexampleenter"
	            objMenu.SelectMenu "View->Query By Example->Enter"        
	        Case "viewquerybyexamplerun"
	            objMenu.SelectMenu "View->Query By Example->Run"
             Case "viewquerybyexamplecancel"
            objMenu.SelectMenu "View->Query By Example->Cancel"
	        Case "toolsviewaccounting"
	            objMenu.SelectMenu "Tools->View Accounting"
	        Case "viewrequests"
	            objMenu.SelectMenu "View->Requests"
		Case "toolsCopyFile"
			objMenu.SelectMenu "Tools->Copy File..."
	End Select
	
	If err.Number<> 0 Then
		fnReportEvent "Fail","Menu Selection","Functiona name  : Menu Selection ,Unable to select/click on the Menu " &MenuNaviagation ,true   
		print err.description
	End If
	
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  26/10/2021
'Name -                 fn_TransactionSearch 
'description:         fn_TransactionSearch : To search the Transaction Number
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_TransactionSearch(pTransNum)
    
    If orcTranscnWindowObj.Exist(20) Then      
        fn_SelectMenu orcTranscnWindowObj,"viewquerybyexampleenter"
        fn_Enter orcTransactionNoObj,pTransNum             
        fn_SelectMenu orcTranscnWindowObj,"viewquerybyexamplerun"           
    Else 
        fnReportEvent "Fail", "Failed to find Transactions Number Status","Failed to find Transactions Number on the form  " & pTransNum,true        
    End If          
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  27/10/2021
'Name -                 fn_setAdjustmentLimit 
'description:         fn_setAdjustmentLimit : To set the Adjustment Limits of User
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

  Function fn_setAdjustmentLimit(frmAmount,toAmount)
On error resume next
blnresult = false
intRowNumber = 1
    If (fn_exist (orcApprovalLimits)) Then
    'orcApprovalLimits.RefreshObject
    orcApprovalLimits.Highlight
    fn_SelectMenu orcApprovalLimits,"viewquerybyexampleenter"
    fn_EnterField orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable),intRowNumber,"User Name",Ucase(environment("SSO_Username")),"User Name"        
    fn_SelectMenu orcApprovalLimits,"viewquerybyexamplerun"
        
        Set objTable = orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable)
        vtotalrows = objTable.GetRoproperty("total rows")
        vCol = objTable.GetRoproperty("columns")
       
       For iRow = 1 To vtotalrows
             vDocType = trim(objTable.GetFieldValue(iRow,2))
             vcurrency =trim(objTable.GetFieldValue(iRow,4))
             print  "vDocType ==" & vDocType
             if (vDocType="Adjustment" or vDocType = "") and vCurrency = gb_TestDataDic.item("Currency") then 
                intRowNumber =iRow
                Exit for 
             End If                                                                        
       Next
        
'        set doctTypelist = orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable).ChildItem(intRecord_no,2,"OracleList",0)
'   	  fn_Select doctTypelist,"Adjustment","Adjustment"
	fn_EnterField  orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable),intRowNumber,"Adjustment","Adjustment","Document type"
        fn_EnterField orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable),intRowNumber,"From  Amount",frmAmount,"From Amount"
        fn_EnterField orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable),intRowNumber,"To Amount",toAmount,"To Amount"
        'orcApprovalLimits.SelectMenu "File->Save"
        fn_SelectMenu orcApprovalLimits, "filesave"
       
        fn_CloseWindow orcApprovalLimits
        fnReportEvent "Pass", "Adjustment Approval Limits Form","Adjustment Approval Limits  has been changed corresponding to the user and limit change to :="  & frmAmount & " to "  & toAmount  ,false
    blnresult = true
    Else 
        fnReportEvent "Fail", "Adjustment Approval Limits Form","Failed to Set Adjustment Approval Limits corresponding to the user",true
    End If
    fn_setAdjustmentLimit = blnresult
    If Err.number <> 0 Then       
    fn_setAdjustmentLimit = blnresult
        fnReportEvent "Fail", "Adjustment Approval Limits Form","Adjustment Approval Limits Page is not displaying and User is not able to set the Adjustment limits for selected User",true
         Exit function
    End If
    
    End  Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  08/10/2021
'Name -                 fn_UserAdjustmentApproval()
'description:         fn_UserAdjustmentApproval():Adjustment approve  
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_UserAdjustmentApproval()
On error resume next
    fn_UserAdjustmentApproval = false
    If (fn_exist (orcTranscnWindowObj)) Then
        fn_CloseWindow orcTranscnWindowObj
    End If
    
    If not (fn_Navigator =  true) Then
    fnReportEvent "Fail", "Function Name = fn_Navigator"," Failed to navigate to Approval Adjustments oracle form", false 
    Exit function
    End If 
'    
        If fn_exist (orcFindAdjustmentsPage.OracleTabbedRegion(orcMain).OracleTextField(orcAdjustmentNum)) = true Then
            fn_Enter orcFindAdjustmentsPage.OracleTabbedRegion(orcMain).OracleTextField(orcAdjustmentNum),vstrAdjstNumber
            fn_click orcFindAdjustmentsPage.OracleButton(orcFind)
            OracleFormWindow(orcApproveAdjst).OracleTable(orctable).OpenDialog 1,"Status"
            fn_Select OracleListOfValues("title:=Approval Statuses"),"Approved", "Approval Status"
            fn_SelectMenu OracleFormWindow(orcApproveAdjst),"filesave"                 
            fn_CloseWindow OracleFormWindow(orcApproveAdjst)
            fnReportEvent "Pass", "Approve Adjustment Page Status","Approve Adjustment Page is displaying and User is able to Approve the Adjustment",false
            fn_UserAdjustmentApproval = true
        Else 
            fnReportEvent "Fail", "Approve Adjustment Page Status","Approve Adjustment Page is not displaying and User is not able to Approve the Adjustment",true
        End If     
    
    If Err.number <> 0 Then  
        fn_UserAdjustmentApproval = false     
        print Err.description
        fnReportEvent "Fail", "Function Name = fn_UserAdjustmentApproval"," Failed to approve the adjustment transaction" &error.description,true
        Exit function
    End If
  End Function   

 '=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  27/10/2021
'Name -                 fn_UserLimitSetAndValidation
'description:         fn_UserLimitSetAndValidation : will set and validate user adjustment limit at start of TC 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_UserLimitSetAndValidation()
On error resume next
    blnresult = false
    If gstrTdIdentifer2 = "IntializeLimit" Then    
        blnresult = fn_setAdjustmentLimit("-9", "9")
    ElseIf gstrTdIdentifer2 ="LimitReset" Then
        blnresult = fn_setAdjustmentLimit("-9999", "9999")
        
    End If 
    
    fn_UserLimitSetAndValidation = blnresult
    
    If Err.number <> 0 Then  
        fn_UserLimitSetAndValidation = false     
        print Err.description
        fnReportEvent "Fail", "Function Name = verifyAccountingAGIStoAR"," Failed to validate Accounting" &error.description,true
        Exit function
    End If
    
    
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  27/10/2021
'Name -                 fn_verifyAccountingAGIStoAR 
'description:         fn_verifyAccountingAGIStoAR : Invoice Processing - Adjustment exceeding Approval Limit
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_verifyAccountingAGIStoAR()
    
    On error resume next   
    fn_TransactionSearch(gb_TestDataDic.item("Transaction_No")) 'Provided the Transaction Num from sheet for only AGIS to AR test case
    vstrTransSource = orcTranscnWindowObj.OracleTextField(orcSource).GetROProperty ("value")        
    If vstrTransSource = "Global Intercompany" Then
           fnReportEvent "Pass", "Transactions Source Status","Transactions Source is " & vstrTransSource & " which is as expected",true
        Else 
        fnReportEvent "Fail", "Transactions Source Status","Transactions Source is " & vstrTransSource & " which is not as expected",true
        Exit function    
    End If    
    fn_click orcTranscnWindowObj.OracleButton(distributions)    
    vstrReceivableGLAccnt = orcDistributionWindowObject.OracleTable(orctable).GetFieldValue( 1,"GL Account")
    vstrReceivableDistrAmnt = orcDistributionWindowObject.OracleTable(orctable).GetFieldValue( 1,"Distribution Amount")    
    vstrRevenueGLAccnt = orcDistributionWindowObject.OracleTable(orctable).GetFieldValue( 2,"GL Account")
    vstrRevenueDistrAmnt = orcDistributionWindowObject.OracleTable(orctable).GetFieldValue( 2,"Distribution Amount")
    
    If vstrReceivableDistrAmnt = vstrRevenueDistrAmnt Then
         fnReportEvent "Pass", "Distribution Amount Status","Distribution Amount of Receivable " & vstrReceivableDistrAmnt &" is matching with Revenue " & vstrRevenueDistrAmnt & " And Account Code combination of Receivable is " & vstrReceivableGLAccnt & " and Revenue is " & vstrRevenueGLAccnt,true
    Else 
        fnReportEvent "Fail", "Distribution Amount Status","Distribution Amount of  Receivable " & vstrReceivableDistrAmnt &" is not matching with " & vstrRevenueDistrAmnt,true
        Exit function
    End If
    orcDistributionWindowObject.CloseWindow            
    If (fn_exist (orcTranscnWindowObj)) Then    
        fn_SelectMenu orcTranscnWindowObj,"toolsviewaccounting"        
        
        If OracleNotification(orcNote).OracleButton(okbutton).Exist(3) Then            
                    OracleNotification(orcNote).OracleButton(okbutton).Click
                    fnReportEvent "Fail", "View accounting generation Status","No accounting exist for this transaction , Kindly run the create accounting first  ",true
                Exit Function   
		fn_verifyAccountingAGIStoAR = false           
            End If
        
            If (fn_exist(ObjSubledgerJournalEntryPage)) Then
                ObjSubledgerJournalEntryPage.Highlight
                fnReportEvent "Pass", "Subledger Journal Entry Page Status","View Accounting link is enabled and user is navigated to Subledger Journal Entry Page successfully ",true
            Else
                fnReportEvent "Fail", "Subledger Journal Entry Page Status","Subledger Journal Entry Page is not exist  ",true
                Exit function
            End If    
           Browser("name:=Subledger Journal Entry.*").Close
        fn_CloseWindow orcTranscnWindowObj
    Else
        fnReportEvent "Fail", "Transaction Page Status","Transaction page is not exist",true
        Exit function
    End If                    
        fn_verifyAccountingAGIStoAR = True    
    If Err.number <> 0 Then  
        fn_verifyAccountingAGIStoAR = false     
        'print Err.description
        fnReportEvent "Fail", "Function Name = verifyAccountingAGIStoAR"," Failed to validate Accounting" &error.description,true
        Exit function
    End If    
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  27/10/2021
'Name -                 fn_Adjustment 
'description:         fn_Adjustment : Invoice Processing - Adjustment exceeding Approval Limit
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_Adjustment()  

    blnresult = False
    On error resume next    
     
        vtransactionnumber = fn_getExecutionResultData(gstrTestCaseExec_id,"Transaction_No")
        fn_TransactionSearch(vtransactionnumber)
        'orcTranscnWindowObj.SelectMenu "Actions->Adjust"
        fn_SelectMenu orcTranscnWindowObj,"actionsadjust"            
        orcAdjustmentsPage.OracleTabbedRegion(orcMain).OracleTable(orctable).OpenDialog 1,"Activity Name"
        fn_Select OracleListOfValues("title:=Activity Names"),gb_TestDataDic.item("Activity_Name"),"Activity Name"
        orcAdjustmentsPage.OracleTabbedRegion(orcMain).OracleTable(orctable).OpenDialog 1,"Amount"
    If (fn_exit (orcAdjustmentsPage)) Then
        'OracleFormWindow("title:=Adjustments.*").SelectMenu "File->Save"
        fn_SelectMenu orcAdjustmentsPage,"filesave"  
        fnReportEvent "Pass", "Adjustments Page Status","Adjustments Page is displaying, User is able to save the Adjustment, Adjustment Number is generated",false
        Else 
        fnReportEvent "Fail", "Adjustments Page Status","Adjustments Page is not displaying , User is not able to save the Adjustment, Adjustment Number is not generated",true
        Exit function    
    End If
    
    Do while OracleNotification("title:=Caution").Exist(5) = True
        OracleNotification("title:=Caution").Approve
        wait 2
    Loop  
    fnReportEvent "Pass", "Caution Pop Up Status","Caution Pop Up is displaying message as User cannot approve this adjustment as its not within users Limit",false
    
    orcAdjustmentsPage.OracleTabbedRegion(orcComment).Select
    vstrAdjstNumber = orcAdjustmentsPage.OracleTabbedRegion(orcComment).OracleTable(orctable).GetFieldValue( 1,"Number")
    vstrAdjstStatus = orcAdjustmentsPage.OracleTabbedRegion(orcComment).OracleTable(orctable).GetFieldValue( 1,"Status")        
    If vstrAdjstStatus = "Waiting Approval" Then    
        fnReportEvent "Pass", "Adjustment Status Page","Adjustment Status is as expected , Adjustment is not Approved as its out of user limit",false
        blnresult = True
    Else  
        fnReportEvent "Fail", "Adjustment Status Page","Adjustment Status is not as expected , Adjustment is Approved as its within user limit",true
        fn_Adjustment = blnresult
        Exit function    
    End  IF    
    'Call fn_switchResponsibility(gb_TestDataDic.item("Responsibility2"))
    fn_CloseWindow orcAdjustmentsPage
    fn_CloseWindow orcInstallmentsForms
    fn_CloseWindow orcTranscnWindowObj
     fn_Adjustment  =blnresult
    
 If Err.number <> 0 Then  
        fn_Adjustment = false     
        print Err.description
        Exit function
    End If
End Function      

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/11/2021
'Name -                 fn_SubmitRequest 
'description:         fn_SubmitRequest : Submit Request - Report Generation
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_SubmitRequest()	

On error resume next 
fn_SubmitRequest=false
'	create the batch job folder location 
Call fn_CreateBatchJobfolder()
	
	If orcRequestWindowObj.exist(2) Then
		fn_CloseWindow orcRequestWindowObj
	End If
	If gstrTdIdentifer2 = "Request1" Then
		ReqName = gb_TestDataDic.item("Request_Name1")
	Else  
		ReqName = gb_TestDataDic.item("Request_Name")
	End If
		Call fn_selectRequestType(gb_TestDataDic.item("Request_Type"))   			'Call TypeOfRequest Fn 
		fn_ReportEnter orcSubmitRequestObj.OracleTextField(orcRequestName),ReqName,"Request Name"

	'Call Parameterfn 
	fn_EnterParameter ReqName
		
	fn_Click orcSubmitRequestObj.OracleButton(orcSubmitBtn)
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Adding this condition for AP brider extract Test case for verifying Caution button presence
	If OracleNotification("title:=Caution").OracleButton(okbutton).Exist(2) Then
		fn_Click OracleNotification("title:=Caution").OracleButton(okbutton)
		fnReportEvent "Pass", "Submit Request Caution Pop Up","Caution Pop Up OK button clicked succesfully",false
	End If
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	If orcDecisionNotificationObj.Exist(5) Then
		strRequestNum = fn_GetNumericValueFromString(orcDecisionNotificationObj.GetROProperty("message"))		
		fn_Click orcDecisionNotificationObj.OracleButton("label:=No")
		strQuery="UPDATE [ExecutionResult$] SET Request_No='"&strRequestNum&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"
		Call fn_updateQuery(strQuery)
	End If

	
	PhaseStatus = fn_CheckRequestStatus(strRequestNum)
	Status = orcRequestWindowObj.OracleTable(orctable).GetFieldValue(1,5)
	If PhaseStatus="Completed" And Status="Normal" Then
		fnReportEvent "Pass","Submitted Request : " &gb_TestDataDic.item("Request_Name"),"Function name  : fn_SubmitRequest , Successfully able to submit request. Request Id is : "&strRequestNum& " Phase Status is : " &PhaseStatus&  " & Status is : " &Status,true   
		fn_SubmitRequest=true
	Else 
		fnReportEvent "Fail","Submitted Request : " &gb_TestDataDic.item("Request_Name")& " Failed ","Function name  : fn_SubmitRequest , Oracle Batch job not run successfully. Request Id is : "&strRequestNum& " Phase Status is : " &PhaseStatus& " & Status is : " &Status,true   
		Exit Function
	End If
'	Adding the specific condition for test cases GSI.O2C.AR.SA.015 
	If ReqName ="MMC AR Invoice Print Selected Invoices (Global)" Then
		If orcRequestWindowObj.Exist(5) Then
			fn_CloseWindow orcRequestWindowObj
		End If
	End If
	
If Err.number <> 0 Then             
	fnReportEvent "Fail","Submit Request","Function name  : fn_SubmitRequest , Failed to submit request.Error is : " &Err.description,true   
	Exit function
End If
	
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  02/11/2021
'Name -                 fn_ViewOutput 
'description:         fn_ViewOutput : View Output in Report Window once Batch Job execution is completed
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_ViewOutput()
On error resume next
	If orcRequestWindowObj.OracleButton(orcViewOutput).Exist(5) Then
		fn_Click orcRequestWindowObj.OracleButton(orcViewOutput)
		If OracleNotification(orcNote).Exist(20) = true Then
			fnReportEvent "Pass","Output File Size 0KB","The output file for request : " &strRequestNum& " is empty (0 bytes).",true
			OracleNotification(orcNote).OracleButton(okbutton).Click
			fn_CloseWindow orcRequestWindowObj
			fn_ViewOutput="0KB"
		End  If
	End If
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  03/11/2021
'Name -                 fn_selectRequestType 
'description:         fn_selectRequestType : Select type of Request to be submitted (Single request/Request Set)
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_selectRequestType(requestType)
On error resume next

If fn_Exist(orcSubmitNewRequestObj)=true Then		
		orcSubmitNewRequestObj.OracleRadioGroup(orcSingleReq).Select requestType
		 fn_Click orcSubmitNewRequestObj.OracleButton(okbutton)
Else 
	fnReportEvent "Fail","Select Request","Function name  : fn_selectRequestType , Unable to select request type ",true
End If
                 
If Err.number <> 0 Then             
	fnReportEvent "Fail","Select Request","Function name  : fn_selectRequestType , Failed to select request.Error is : " &Err.description,true   
	Exit function
End If
	
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  02/11/2021
'Name -                 fn_EnterParameter 
'description:         fn_EnterParameter : Enter different Parameters based on Report Name
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_EnterParameter(RequestName)
	On error resume Next
	var_date = fn_getSysdateFormat("DD-MMM-YYYY")
	
	Select Case Ucase(RequestName)
	
	Case "MMC GLB AR BILLING AND RECEIPT HISTORY"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcOperatingUnit),gb_TestDataDic.item("Request_Operating_Unit"),"Operating Unit"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcSetOfBooks),gb_TestDataDic.item("Set_Of_Books"),"Set Of Books"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcTranscnDateLow),gb_TestDataDic.item("Transcn_Date_Low"),"Transaction Date Low"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcTranscnDateHigh),gb_TestDataDic.item("Transcn_Date_High"),"Transaction Date High"
	Case "MMC GLB AR OUTSTANDING INVOICES LISTING"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcOperatingUnit),gb_TestDataDic.item("Request_Operating_Unit"),"Operating Unit"
	Case "INVOICE EXCEPTION REPORT"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcGLDateFrom),gb_TestDataDic.item("GL_Date_From"),"GL Date From"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcGLDateTo),gb_TestDataDic.item("GL_Date_To"),"GL Date To"
	Case "REVENUE RECOGNITION PROGRAM"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcPrintFormat),gb_TestDataDic.item("Print_Format"),"Print Format"
	Case "REVENUE RECOGNITION MASTER PROGRAM"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcPrintFormat),gb_TestDataDic.item("Print_Format"),"Print Format"
	Case "MMC GLB AR2GL CREATE BAD DEBTS JOURNALS (PART1)"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcAsOfDate),gb_TestDataDic.item("As_On_Date"),"As On Date"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcOperatingUnit),gb_TestDataDic.item("Request_Operating_Unit"),"Operating Unit"
	Case "MMC GLB AR2GL CREATE BAD DEBTS JOURNALS (PART2)"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcOperatingUnit),gb_TestDataDic.item("Request_Operating_Unit"),"Operating Unit"
	Case "CREATE ACCOUNTING"
	        orcParametersWindowObj.OracleTextField(ledgerField).OpenDialog
	        fn_ReportEnter orcParametersWindowObj.OracleTextField(endDateField),var_date,"End Date"
	        fn_ReportEnter orcParametersWindowObj.OracleTextField(orcMode),gb_TestDataDic.item("Accounting_Mode"),"Accounting Mode"
	        fn_ReportEnter orcParametersWindowObj.OracleTextField(orcReport),"Detail"," Report Type"
	        fn_ReportEnter orcParametersWindowObj.OracleTextField(orcTransferToGL),gb_TestDataDic.item("Transfer_GL"),"Transfer To GL"
	        fn_ReportEnter orcParametersWindowObj.OracleTextField(orcPostToGL),gb_TestDataDic.item("Post_GL"),"Post To GL"
	        fn_ReportEnter orcParametersWindowObj.OracleTextField(orcIncludeUTI),"Yes","Include User Transaction Identifiers"	        
        Case "MMC AR INVOICE PRINT SELECTED INVOICES (GLOBAL)"
      		fn_Select OracleListOfValues("title:=Reports"),gb_TestDataDic.item("Request_Operating_Unit"),"Operating Unit"
	End  Select
	
	If RequestName <> "MMC AP Supplier Bridger Extract" Then
		fn_Click orcParametersWindowObj.OracleButton(okbutton)
	End If 
	
	If err.Number<> 0 Then
		fnReportEvent "Fail","Enter Parameters","Function name  : Enter Parameter,Unable to enter parameters for Request : " &RequestName ,true   
	
	End If
		
End Function


'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  08/11/2021
'Name -                 fn_CopyFile 
'description:         fn_CopyFile : Selects Copy_File Menu & Opens Output in Browser Window
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_CopyFile()
On error resume next
fn_CopyFile=false
	If OracleFormWindow(orcSTReport).exist(5) Then
		OracleFormWindow(orcSTReport).SelectMenu "Tools->Copy File..."
		fnReportEvent "Pass","Copy File","Function Name : fn_CopyFile. Copy File clicked successfully",true
		fn_CopyFile=true		
	End If
	
	If err.Number <> 0 Then
		fnReportEvent "Fail","Copy File","Function name  : fn_CopyFile,Unable to copy Output File." ,true   
		print err.description
	End If
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  09/11/2021
'Name -                 fn_ValidateOutput 
'description:         fn_ValidateOutput : Select View Output/ CopyFile Options in respective cases
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_ValidateOutput()
On error resume next
blnresult=false
	If gstrTdIdentifer2 = "Request1" Then
		ReqName = gb_TestDataDic.item("Request_Name1")
	Else  
		ReqName = gb_TestDataDic.item("Request_Name")
	End If
	
'	Adding the specific condition for TC GSI.O2C.AR.SA.015
	OutputFileSize=fn_ViewOutput
	If OutputFileSize="0KB" Then
		fn_ValidateOutput=true
		Exit Function
	End If
	'Added one more "OR" condition  for AP TC GSI.P2P.AP.SA.025
	If ReqName="Invoice Exception Report" OR ReqName="MMC GLB AR Billing and Receipt History" OR ReqName="MMC GLB AR Outstanding Invoices Listing" OR ReqName="Revenue Recognition" OR ReqName="Revenue Recognition Master Program" OR ReqName="MMC AP Supplier Bridger Extract"  Then
		If OracleFormWindow(orcSTReport).Exist(5) Then
			fnReportEvent "Pass","Validate Output","Function name  : fn_ValidateOutput. Output is generated for submitted Request : "&ReqName,true
			CopyFileStatus = fn_CopyFile
			If CopyFileStatus=true Then
				blnresult = fn_CopyReportOutputToExcel (ReqName)
				
			Else 	
				fnReportEvent "Fail","Copy File","Function name  : fn_CopyFile,Unable to copy Output File." ,true  
			End If
			fn_CloseWindow OracleFormWindow(orcSTReport)
		End If
	ElseIf ReqName = "Create Accounting" Then
		'Code to Save Create Accounting Report 
		
		fileLocation = environment.Value("BatchJobEntityfolder")& "\CreateAccounting" &fn_RandomNumber(3) &".rtf"
		blnresult = fn_fileDownload(fileLocation)
	End  If 
	
fn_CloseWindow orcRequestWindowObj
	If orcSubmitRequestObj.Exist(10) Then
		fn_CloseWindow orcSubmitRequestObj
	End If
fn_ValidateOutput = blnresult
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  09/11/2021
'Name -                 fn_ValidateOutputInFile 
'description:         fn_ValidateOutputInFile : Validate output of one test case into Report 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_ValidateOutputInFile(TextContent)
On error resume next
fn_ValidateOutputInFile=false
vTransactionNo = fn_getExecutionResultData("GSI.O2C.AR.SA.006","Transaction_No")

	If vTransactionNo=0 or vTransactionNo="" or Isnull(vTransactionNo) Then
		fnReportEvent "Fail","Transaction Number","Failed to fetch the Transaction Number for : GSI.O2C.AR.SA.006",false
		Exit Function
	ElseIf instr(1,TextContent,vTransactionNo)>0 Then
		fnReportEvent "Pass","Transaction Number","Transaction Number : "&vTransactionNo& " of test case GSI.O2C.AR.SA.006 is present in Output File",false	
		fn_ValidateOutputInFile=true
	End If
End Function


'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  09/11/2021
'Name -                 fn_CopyReportOutputToExcel 
'description:         fn_CopyReportOutputToExcel : Copy Report Content opened in Browser into Excel Sheet & save it in respective entity folder with timestamp 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_CopyReportOutputToExcel(ReportName)
On error resume next
blnresult = false
vdate = fn_getSysdateFormat("DD-MMM-YYYY")
vdatetime=now()
vtimestamp =replace(replace(vdatetime,"/",""),":","")

If Browser(OutputBrowserTitle).Page(OutputBrowserURL).Exist(5) Then
		Browser(OutputBrowserTitle).Page(OutputBrowserURL).WebElement(output_xpath).Click
		fnReportEvent "Pass","Detailed Report generated","Able to Launch Report in the browser",true
		sText = Browser(OutputBrowserTitle).Page(OutputBrowserURL).WebElement(output_xpath).GetROProperty("innertext")
		 strEntityfolder =environment.Value("BatchJobEntityfolder") 
			
		Set objExcel = CreateObject("Excel.Application")		
		sExcelPath =  strEntityfolder & "\" &ReportName& "_" &vtimestamp& ".xls"	
		objExcel.Visible = true

	 	 objExcel.Workbooks.Add()
		 objExcel.Workbooks.Open sExcelPath
		 objExcel.Cells(1,1).value = sText
		 objExcel.ActiveWorkbook.SaveAs (sExcelPath)
		objExcel.Quit

		If fn_FileExist(sExcelPath) Then
			blnresult = true
			 fnReportEvent "Pass","Report Output","Saved Report content successfully in " &ReportName& ".xls excel file. File Location : "&sExcelPath,false
		else
			blnresult = false
		End If	

		If ReportName="MMC GLB AR Billing and Receipt History" Then
			blnresult =  fn_ValidateOutputInFile(sText)
		End If	
	Else
		fnReportEvent "Fail","Detailed Report not generated","Unable to Launch Report in the browser",true
	End  If 	
		fn_CopyReportOutputToExcel = blnresult
		Browser(OutputBrowserTitle).Sync		
		Browser(OutputBrowserTitle).close

If err.Number <> 0 Then
	fnReportEvent "Fail","Validate Output","Function name  : fn_CopyReportOutputToExcel , Failed to Validate Output in Browser" ,true   
End If
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  03/11/2021
'Name -                 fn_CheckRequestStatus 
'description:         fn_CheckRequestStatus : Validate Status of submitted Request
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_CheckRequestStatus(strReqId)
On error resume next
ReqName = gb_TestDataDic.item("Request_Name") 
	If strReqId <> "" Then
		fn_SelectMenu orcHomePageNavigator,"viewrequests"
		orcFindRequestObj.OracleRadioGroup(orcAllReq).Select "Specific Requests"	
		'Add elseIF condition in case user need to find different reports		
		If ReqName = "MMC AR Invoice Print Selected Invoices (Global)" Then
			var_date = fn_getSysdateFormat("DD-MMM-YYYY")
			fn_ReportEnter orcFindRequestObj.OracleTextField(orcRequestName),"EN-US: (MMC AR Invoice Print Selected Invoices Print PDF)","Request Name"
			fn_ReportEnter orcFindRequestObj.OracleTextField("description:=Date Submitted"),var_date,"Date Submitted"
			orcFindRequestObj.OracleButton(orcFind).RefreshObject
		Else 			
			fn_ReportEnter orcFindRequestObj.OracleTextField("description:=Request ID"),strReqId,"Request ID"
		End If		
		
           	fn_Click orcFindRequestObj.OracleButton(orcFind)
'           	Added below condition because during batch execution its not able to click on the find button for below specific request 
		If ReqName = "MMC AR Invoice Print Selected Invoices (Global)" and orcRequestWindowObj.OracleButton(orcRefreshData).exist(3) = false Then					
				orcFindRequestObj.OracleButton(orcFind).RefreshObject
				fn_Click_fieldname orcFindRequestObj.OracleButton(orcFind),"FindButton"							
		End If
		
		intLoopCnt = 1
               Do
                    fn_Click orcRequestWindowObj.OracleButton(orcRefreshData)
                    strPhaseStatus = orcRequestWindowObj.OracleTable(orctable).GetFieldValue(1,4)
                    wait 2
                    If intLoopCnt = 120 Then
                        Exit Do
                    End If
                    intLoopCnt = intLoopCnt +1
                Loop While strPhaseStatus <>"Completed"
                fn_CheckRequestStatus = strPhaseStatus
              fnReportEvent "Pass","Check Request Status","Function name  : fn_CheckRequestStatus. Request Status is : "&strPhaseStatus ,false   
        Else 
        	fnReportEvent "Fail","Check Request Status","Function name  : fn_CheckRequestStatus , Request Status is not: "&strPhaseStatus ,true   
      	End  If 
If err.Number<> 0 Then
	fnReportEvent "Fail","Check Request Status","Function name  : fn_CheckRequestStatus , Failed to Check Request Status" ,true   
	
End If
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  03/11/2021
'Name -                 fn_GetNumericValueFromString 
'description:         fn_GetNumericValueFromString : Extract the Request No from the popup
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_GetNumericValueFromString(str)
Dim c,reqId
	for x=1 to len(str)
	c=mid(str,x,1)
	If isnumeric(c) OR c="," then
		reqId=reqId&c
	End If 
	next
fn_GetNumericValueFromString = reqId
End Function


'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  03/11/2021
'Name -                 fn_fileDownload 
'description:         fn_fileDownload : Window file download
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_fileDownload(fileLocation)

On error resume next

blnresult = false 

Set objOracleAppR12 = Browser("title:=Oracle Applications R12.*")
Set objDownloadFile = objOracleAppR12.WinObject("text:=Do you want to open or save Create_Accounting.*")

    If objDownloadFile.WinButton("acc_name:=6").Exist(5) Then
        objDownloadFile.WinButton("acc_name:=6").Click
        objOracleAppR12.WinMenu("MenuObjType:=3").Select "Save as"
        wait 4
        Browser("Oracle Applications R12").Dialog("Save As").Highlight        
        Browser("Oracle Applications R12").Dialog("Save As").WinEdit("File name:").Set fileLocation
        Browser("Oracle Applications R12").Dialog("Save As").WinButton("Save").Click
                
            If Browser("Oracle Applications R12").WinObject("Notification").WinButton("Close").Exist(10) then
                Browser("Oracle Applications R12").WinObject("Notification").WinButton("Close").Click
            End  If 
    Else
         fnReportEvent "Fail","Download File","Pop Up to downlod file is not existing" ,true    
        Exit function                
    End If
    
'    validating the file NEED TO ADD THE WAIT AS IT TAKE TIME TO DOWNLOAD
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Wait(2)
    If objFso.FileExists(fileLocation) Then
        'objFso.CreateFolder(fileLocation)
        fnReportEvent "Pass","Download File Location","Succefully Downloaded the file at:" &fileLocation ,false
        blnresult = true
    Else 
        fnReportEvent "Fail","Download File Location","Failed to Download the file at:" &fileLocation ,false
        Exit Function
    End If
    fn_fileDownload = blnresult
 
    If Err.number <> 0 Then             
          fnReportEvent "Fail", "Download File","Pop Up to downlod file is not existing" ,true
     End If  
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  03/11/2021
'Name -                 fn_viewAccounting 
'description:         fn_viewAccounting : Check Accounting Entries
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_viewAccounting()
On error resume next 
blnresult = false
        vtransactionnumber = fn_getExecutionResultData("GSI.O2C.AR.SA.007","Transaction_No")
        fn_TransactionSearch(vtransactionnumber)
        
        If orcTranscnWindowObj.Exist(5) Then    
        fn_SelectMenu orcTranscnWindowObj,"toolsviewaccounting"        
        
            If OracleNotification(orcNote).OracleButton(okbutton).Exist(3) Then            
                    OracleNotification(orcNote).OracleButton(okbutton).Click
                    fnReportEvent "Fail", "View accounting generation Status","No accounting exist for this transaction , Kindly run the create accounting first  ",true
                    fn_CloseWindow orcTranscnWindowObj
               	fn_viewAccounting = blnresult           
               	  Exit function
            End If
                 
                If (fn_exist(ObjSubledgerJournalEntryPage)) Then
                    ObjSubledgerJournalEntryPage.Highlight
                    blnresult = true
                    fnReportEvent "Pass", "Subledger Journal Entry Page Status","View Accounting link is enabled and user is navigated to Subledger Journal Entry Page successfully ",true
                Else
                    fnReportEvent "Fail", "Subledger Journal Entry Page Status","Subledger Journal Entry Page is not exist  ",true
                    fn_viewAccounting = blnresult
                     Exit function 
                End If      
                
            ObjSubledgerJournalEntryPage.WebButton("title:=View Journal Entry").Click
            
            If ObjSubledgerJournalEntryPage.WebElement("innerhtml:=Lines","innertext:=Lines","outertext:=Lines").Exist(5) Then
                strAcc1 =  ObjSubledgerJournalEntryPage.WebElement("html id:=N14:Account:0").GetROProperty("innertext")
                strAccClass1 =  ObjSubledgerJournalEntryPage.WebElement("html id:=N14:AccountingClass2:0").GetROProperty("innertext")
                strAcc2 =  ObjSubledgerJournalEntryPage.WebElement("html id:=N14:Account:1").GetROProperty("innertext")
                strAccClass2 =  ObjSubledgerJournalEntryPage.WebElement("html id:=N14:AccountingClass2:1").GetROProperty("innertext")
                strAcc3 =  ObjSubledgerJournalEntryPage.WebElement("html id:=N14:Account:2").GetROProperty("innertext")
                strAccClass3 =  ObjSubledgerJournalEntryPage.WebElement("html id:=N14:AccountingClass2:2").GetROProperty("innertext")
                strAcc4 =  ObjSubledgerJournalEntryPage.WebElement("html id:=N14:Account:3").GetROProperty("innertext")
                strAccClass4 =  ObjSubledgerJournalEntryPage.WebElement("html id:=N14:AccountingClass2:3").GetROProperty("innertext")
                
                fnReportEvent "Pass", "Subledger Journal Entry Lines status","Lines account details are displaying for " & strAccClass1 & " = " & strAcc1 & " , " & strAccClass2 & " = " & strAcc2 &  " & " & strAccClass3 & " = " & strAcc3& " & " & strAccClass4 & " = " & strAcc4,true
                blnresult = true
            Else
                fnReportEvent "fail", "Subledger Journal Entry Lines status","Lines account details are not displaying ",true       
 		 Browser("name:=Subledger Journal Entry.*").Close                 
            	 fn_viewAccounting = false 
            	 Exit Function
  		
            End If
            
                strAccDr = ObjSubledgerJournalEntryPage.WebElement("html id:=TotalAccountedDr").GetROProperty("innertext")
                strAccCr = ObjSubledgerJournalEntryPage.WebElement("html id:=TotalAccountedCr").GetROProperty("innertext")
                
                    If strAccDr = strAccCr Then
                    		blnresult =true
                            fnReportEvent "Pass", "View Journal Entry Page Status","Accounting entries are visible with all the line details and Account Dr =" & strAccDr & " is matching with Account Cr =" & strAccCr,true
                    Else
                            fnReportEvent "Fail", "View Journal Entry Page Status","SLA Accounting entries are not visible with all the line details",true                        
		               blnresult =False        
                    End If
             
            ObjSubledgerJournalEntryPage.Link("text:=Return to Subledger Journal.*").Click    
            Browser("name:=Subledger Journal Entry.*").Close 
            fn_CloseWindow orcTranscnWindowObj
    Else
       fnReportEvent "Fail", "Transaction Page Status","Transaction page is not exist",true        
	blnresult  = false           	
    End If                    
        fn_viewAccounting = blnresult    
        
    If Err.number <> 0 Then  
        fn_viewAccounting = false     
        fnReportEvent "Fail", "Function Name = fn_viewAccounting"," Failed to view Accounting" &error.description,true
        Exit function
    End If    
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  10/11/2021
'Name -                 fn_CreateBatchJobfolder 
'description:         fn_CreateBatchJobfolder : Create Batch Job Folder 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_CreateBatchJobfolder()
	
	Set objFso=CreateObject("Scripting.FileSystemObject")
	 vdate = fn_getSysdateFormat("DD-MMM-YYYY")
	vdatetime=now()
	vtimestamp =replace(replace(vdatetime,"/",""),":","")

	 'Create Batch job Report Folder 
	 strBatchJobReportFolderPath =  mid(environment("TestDir"),1,InStrRev(environment("TestDir"),"\")) & "Log\Batch Job Reports"	
	If Not(objFso.FolderExists(strBatchJobReportFolderPath)) Then
		objFso.CreateFolder(strBatchJobReportFolderPath)
	End If
	 
	 'Create Date Folder 
	 strDateFolderPath = strBatchJobReportFolderPath & "\"&vdate 
	 If Not(objFso.FolderExists(strDateFolderPath)) Then
		objFso.CreateFolder(strDateFolderPath)
	End If
	 
	 'Create Entity Folder 
	 strEntityfolder = strDateFolderPath & "\" &environment("Legal_entity")
	 environment.Value("BatchJobEntityfolder") = strEntityfolder
	If Not(objFso.FolderExists(strEntityfolder)) Then
		objFso.CreateFolder(strEntityfolder)
	End If
End Function
'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  17/11/2021
'Name -                 fn_Create_CaptureReceiptNo 
'description:         fn_Create_CaptureReceiptNo : Create & Capture generated ReceiptNo No & Save it is Execution Result Tab of Test Data Sheet
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_Create_CaptureReceiptNo()
On error resume next 
	vreceiptnumber = "REC" & fn_RandomNumber(4)
	strQuery="UPDATE [ExecutionResult$] SET Receipt_No='"&vreceiptnumber&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"
	Call fn_updateQuery(strQuery)
	If vreceiptnumber <> "" Then
		fnReportEvent "Pass","Receipt No","Receipt No  "& vreceiptnumber &" has been generated",falses
		fn_Create_CaptureReceiptNo=true
	else
		fnReportEvent "Fail","Receipt No","Receipt No is not generated",True
	    	fn_Create_CaptureReceiptNo=false
	End If
fn_Create_CaptureReceiptNo=vreceiptnumber
If Err.number <> 0 Then             
	print Err.description,true
	fnReportEvent "Fail","Capture Receipt No","Function name  : fn_Create_CaptureReceiptNo , Failed to Create & Capture Receipt No. Error is : " &Err.description,true   
	fn_Create_CaptureReceiptNo=false
	Exit function
End If
End Function


Function fn_ValidateLimitSetUp()
On error resume next
blnresult = false
intRowNumber = 1
limitcheck = cint(gb_TestDataDic.item("Amount_Applied"))

    If (fn_exist (orcApprovalLimits)) Then
    orcApprovalLimits.Highlight
    fn_SelectMenu orcApprovalLimits,"viewquerybyexampleenter"
    fn_EnterField orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable),intRowNumber,"User Name",Ucase(environment("SSO_Username")),"User Name"        
    fn_SelectMenu orcApprovalLimits,"viewquerybyexamplerun"
        
        Set objTable = orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable)
        vtotalrows = objTable.GetRoproperty("total rows")
        vCol = objTable.GetRoproperty("columns")
       
       For iRow = 1 To vtotalrows
             vDocType = trim(objTable.GetFieldValue(iRow,2))             
             if vDocType= gb_TestDataDic.item("Apply_To_Field") or vDocType = "" then 
                intRowNumber =iRow
                Exit for 
             End If                                                                        
       Next
else

End  IF
' fetech the from ammout and to Ammount 

 vToAmount  = cint(trim(orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable).GetFieldValue(irow,"To Amount")))
vFromAmount =orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable).GetFieldValue(irow,"From  Amount")      

If limitcheck <= vToAmount  Then
	blnresult = true 
Else 
	fnReportEvent "Fail","fn_ValidateLimitSetUp","Function name  : fn_ValidateLimitSetUp , Limit is not set within From and To Range",true   
End If
	fn_ValidateLimitSetUp = blnresult
	fn_CloseWindow orcApprovalLimits
If Err.number <> 0 Then             
	fnReportEvent "Fail","fn_ValidateLimitSetUp","Function name  : fn_ValidateLimitSetUp , Failed to validate setup  : " &Err.description,true   
	fn_ValidateLimitSetUp=false
	Exit function
End If

	
End Function
