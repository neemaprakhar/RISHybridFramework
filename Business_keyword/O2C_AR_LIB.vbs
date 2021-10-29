
Public orcTranscnWindowObj,orcLinesWindowObj,orcDistributionWindowObject,orcFormLineItemDescObj,orcRuleAccntObj,orcTransactionNoObj,orcDistTableObj,orcCreditTranscnWindowObj
Public orcReceiptWindowObj,orcRcptCustomerNum,orcRcptApplicationsWindowObj,orcReceiptSummaryWindowObj
Public orcRcptBatchWindowObj,orcRcptTitleObj,objRecSet
Public orcParametersWindowObj,orcRequestWindowObj,orcFindRequestObj,orcSubmitRequestObj,orcDecisionNotificationObj,orcSubmitNewRequestObj
Public orcAdjustmentsPage,orcFindAdjustmentsPage,ObjSubledgerJournalEntryPage ,orcInstallmentsForms,orcApprovalLimits,vstrAdjstNumber 

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



Function fn_Navigator()
On Error Resume Next
	fn_Navigator =false
	
		if gstrTdIdentifer2 = "Responsibility1" then
			vnavigator ="OracleNavigator1"
		elseif gstrTdIdentifer2 = "Responsibility2" then 		
			vnavigator = "OracleNavigator2"
		elseif gstrTdIdentifer2 = "Responsibility3" then 	
			vnavigator = "OracleNavigator3"
		 Else 
            		vnavigator = "OracleNavigator4"
        	End  If 

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
				ElseIf lcase(pNav2) = lcase("Transactions-->Approval Limits") Then
		                    OracleFormFuncationsList.Activate(7)
		                    navg3 = split(pNav2,"-->")(1)
		                    OracleFormFuncationsList.Select(navg3)
		                    OracleFormFuncationsList.Activate(navg3)
				Else 
					OracleFormFuncationsList.Activate(pNav2)		
				End If
			End If	
				intcounter = intcounter +1	
		Loop until orcTranscnWindowObj.OracleTextField(orcSource).Exist=false or counter<=1
	
	fn_Navigator =true
	If err.number<>0 Then
		fn_NavigateOraclePage =false
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
	vreceiptnumber = "REC" & fn_RandomNumber(4)
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptNumber),vreceiptnumber,"Receipt Number"
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcNetReceiptAmount),gb_TestDataDic.item("Receipt_Amount"),"Net Receipt Amount"
	fn_ReportEnter orcRcptCustomerNum,gb_TestDataDic.item("Customer_Number"),"Customer Number"
	fn_Click orcReceiptWindowObj.OracleButton(orcApplyBtn)	
	fn_exist orcRcptApplicationsWindowObj.OracleTable(orctable)
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
	fn_ReportEnter orcRcptBatchWindowObj.OracleTextField(orcControlCount),gb_TestDataDic.item("Total_Count"),"Count"
	fn_ReportEnter orcRcptBatchWindowObj.OracleTextField(orcControlAmount),gb_TestDataDic.item("Total_Amount"),"Amount"	
	fn_Click orcRcptBatchWindowObj.OracleButton(orcRcptBtn)
	vreceiptnumber = "REC" & fn_RandomNumber(4)
	fn_EnterField orcReceiptSummaryWindowObj.OracleTable(orctable),intRecord_no,"Receipt Number",vreceiptnumber,"Receipt No"
	fn_EnterField  orcReceiptSummaryWindowObj.OracleTable(orctable),intRecord_no,"Net Amount",gb_TestDataDic.item("Net_Amount"),"Net Amount"
	fn_Click orcReceiptSummaryWindowObj.OracleButton(orcOpenBtn)
	fn_Enter orcRcptTitleObj.OracleTabbedRegion(orcMain).OracleTextField(orcCustNumber),gb_TestDataDic.item("Customer_Number")
	fn_Click orcRcptTitleObj.OracleButton(orcApplyBtn)
	fn_EnterField orcRcptApplicationsWindowObj.OracleTable(orctable),intRecord_no,"Apply To",vtransactionnumber,"Apply To:Transcn No"
	fn_SelectMenu orcRcptApplicationsWindowObj,"filesave"
	fn_CloseWindow orcRcptApplicationsWindowObj
	fn_CloseWindow orcRcptTitleObj
	fn_CloseWindow orcReceiptSummaryWindowObj
	
	vstrActualAmountCheck = OracleFormWindow("title:=Receipt Batches.*").OracleTextField("description:=Totals: Actual Amount").GetROProperty("value")
	ExpActualAmount = gb_TestDataDic.item("Net_Amount")
	
	If Cint(vstrActualAmountCheck) = Cint(ExpActualAmount) Then
		fnReportEvent "Pass","Actual Amount Check","Actual Amount & Receipt Amount are matching and Amount is = "&ExpActualAmount,false
		fn_ApplyReceiptByBatch=true
	Else 
		fnReportEvent "Fail","Actual Amount Check : Fail", "Actual Amount & Receipt Amount are not matching. Expected Amount is = "&ExpActualAmount,true
	End If	
	
	vstrDifferenceAmountCheck = OracleFormWindow("title:=Receipt Batches.*").OracleTextField("description:=Totals: Difference Amount").GetROProperty("value")
	ExpDifferenceAmount = gb_TestDataDic.item("Total_Amount") - gb_TestDataDic.item("Net_Amount")
	
	If Cint(vstrDifferenceAmountCheck) = Cint(ExpDifferenceAmount)  Then
		fnReportEvent "Pass","Difference Amount Check","Difference Amount is displayed as expected & Amount is = "&ExpDifferenceAmount,false
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

If fn_exist(orcReceiptWindowObj) = true Then
	fnReportEvent "Pass","Receipt Window Status","Successfully loaded Oracle Receipt Form Window",false
	vreceiptnumber = "REC" & fn_RandomNumber(4)	
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptMethod),gb_TestDataDic.item("Receipt_Method"),"Receipt Method"
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptNumber),vreceiptnumber,"Receipt Number"	
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptAmount),gb_TestDataDic.item("Receipt_Amount"),"Receipt Amount"	
	fn_ReportEnter orcRcptCustomerNum,gb_TestDataDic.item("Customer_Number"),"Customer Number"	
	fn_Click OracleFormWindow("title:=Receipts.*").OracleButton("description:=Apply")	
	
	If fn_exist(orcRcptApplicationsWindowObj) Then		
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

If (fn_exist(orcTranscnWindowObj.OracleTextField(orcSource)))  Then
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
title =  orcHomePageNavigator.GetROProperty("title")
	If Instr(1,title,respName) > 1 Then
		fnReportEvent "Pass", "Navigator Page Status","Navigator Page is displaying and User is able to switch the Responsibility to "&respName,false
		blnresult = true
	else
		fnReportEvent "Fail", "Navigator Page Status","Navigator Page is not displaying or Responsibility is not present for that user "& respName ,true
		
	End If

blnresult = fn_Navigator()
fn_switchResponsibility =blnresult
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
	'print Err.description,true
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
'		fn_CreditNote=false
		Exit function
	End If
'Else 
'	fnReportEvent "Fail","Responsibility Status","Unable to Switch Responsibility",true  
'End if
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
intRecord_no = 1
    If (fn_exist (orcApprovalLimits)) Then
    'orcApprovalLimits.RefreshObject
    orcApprovalLimits.Highlight
    fn_SelectMenu orcApprovalLimits,"viewquerybyexampleenter"
    fn_EnterField orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable),intRecord_no,"User Name",Ucase(environment("SSO_Username")),"User Name"        
    fn_SelectMenu orcApprovalLimits,"viewquerybyexamplerun"
        
        Set objTable = orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable)
        vtotalrows = objTable.GetRoproperty("total rows")
        vCol = objTable.GetRoproperty("columns")
       
       For iRow = 1 To vtotalrows
             vDocType = objTable.GetFieldValue(iRow,2)
             print  "vDocType ==" & vDocType
             if vDocType="Adjusment" or vDocType = "" then 
                intRecord_no =iRow
                Exit for 
             End If                                                                        
       Next
        
        set doctTypelist = orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable).ChildItem(intRecord_no,2,"OracleList",0)
     fn_Select doctTypelist,"Adjusment","Adjustment"
        fn_EnterField orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable),intRecord_no,"From  Amount",frmAmount,"From Amount"
        fn_EnterField orcApprovalLimits.OracleTabbedRegion(orcMain).OracleTable(orctable),intRecord_no,"To Amount",toAmount,"To Amount"
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
            If (fn_exist(ObjSubledgerJournalEntryPage)) Then
                ObjSubledgerJournalEntryPage.Highlight
                fnReportEvent "Pass", "Subledger Journal Entry Page Status","View Accounting link is enabled and user is navigated to Subledger Journal Entry Page successfully ",true
            Else
                fnReportEvent "Fail", "Subledger Journal Entry Page Status","Subledger Journal Entry Page is not exist  ",true
                Exit function
            End If    
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

Function fn_SubmitRequest()	

On error resume next 
fn_SubmitRequest=false
'If fn_exist(orcFindRequestObj)=true Then
'fn_Click orcFindRequestObj.OracleButton(orcSubmitNewRequest)
	
	'Call TypeOfRequest Fn 
	Call fn_selectRequestType(gb_TestDataDic.item("Request_Type"))
	fn_ReportEnter orcSubmitRequestObj.OracleTextField(orcRequestName),gb_TestDataDic.item("Request_Name"),"Request Name"
	'fn_WSSendKeys TAB
	'Call Parameterfn 
	fn_EnterParameter gb_TestDataDic.item("Request_Name")
	
	fn_Click orcSubmitRequestObj.OracleButton(orcSubmitBtn)
	
	If orcDecisionNotificationObj.Exist(5) Then
		strRequestNum = fn_GetNumericValueFromString(orcDecisionNotificationObj.GetROProperty("message"))
		fn_Click orcDecisionNotificationObj.OracleButton("label:=No")
	End If

	'Call fn_CheckRequestStatus(strRequestNum)
	Status = fn_CheckRequestStatus(strRequestNum)
	If Status="Completed" Then
		fnReportEvent "Pass","Submit Request","Function name  : fn_SubmitRequest , Successfully able to submit request : " &gb_TestDataDic.item("Request_Name")&  " Request Id is : "&strRequestNum,true   
		fn_SubmitRequest=true
	Else 
		fnReportEvent "Fail","Submit Request Failed ","Function name  : fn_SubmitRequest , Oracle Batch job not run successfully " &gb_TestDataDic.item("Request_Name")&  " Request Id is : "&strRequestNum,true   
	End If
	
	If orcRequestWindowObj.OracleButton(orcViewOutput).Exist(5) Then
		fn_Click orcRequestWindowObj.OracleButton(orcViewOutput)
	End If
    	'strRequestNum = Empty
'Else 
	'fnReportEvent "Fail","Submit Request","Function name  : fn_SubmitRequest , Unable to submit request",true   
'End If
'Call Validate Fn
fn_CloseWindow OracleFormWindow("short title:=Report")
fn_CloseWindow orcRequestWindowObj

If Err.number <> 0 Then             
	fnReportEvent "Fail","Submit Request","Function name  : fn_SubmitRequest , Failed to submit request.Error is : " &Err.description,true   
	Exit function
End If
	
End Function

Function fn_selectRequestType(requestType)
On error resume next
fn_Highlight orcSubmitNewRequestObj
If fn_exist(orcSubmitNewRequestObj)=true Then
		'fn_Highlight orcSubmitNewRequestObj.OracleRadioGroup(orcSingleReq) 
		orcSubmitNewRequestObj.OracleRadioGroup(orcSingleReq).Select requestType
		 fn_Click orcSubmitNewRequestObj.OracleButton(okbutton)
'              If requestType = "Single Request" Then   
'			'fn_Highlight orcSubmitNewRequestObj.OracleRadioGroup(orcSingleReq)              
'                      orcSubmitNewRequestObj.OracleRadioGroup(orcSingleReq).Select "Single Request"
'              ElseIf requestType = "Request Set" Then                 
'                      orcSubmitNewRequestObj.OracleRadioGroup(orcSingleReq).Select "Request Set"
'              End If
Else 
	fnReportEvent "Fail","Select Request","Function name  : fn_selectRequestType , Unable to select request type ",true
End If
                 
If Err.number <> 0 Then             
	fnReportEvent "Fail","Select Request","Function name  : fn_selectRequestType , Failed to select request.Error is : " &Err.description,true   
	Exit function
End If
	
End Function


Function fn_EnterParameter(RequestName)
	On error resume Next
	
	Select Case Ucase(RequestName)
	
	Case "MMC GLB AR BILLING AND RECEIPT HISTORY"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcOperatingUnit),gb_TestDataDic.item("Request_Operating_Unit"),"Operating Unit"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcSetOfBooks),gb_TestDataDic.item("Set_Of_Books"),"Set Of Books"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcTranscnDateLow),gb_TestDataDic.item("Transcn_Date_Low"),"Transaction Date Low"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcTranscnDateHigh),gb_TestDataDic.item("Transcn_Date_High"),"Transaction Date High"
	Case "MMC GLB AR OUTSTANDING INVOICES LISTING"
		fn_ReportEnter orcParametersWindowObj.OracleTextField(orcOperatingUnit),gb_TestDataDic.item("Request_Operating_Unit"),"Operating Unit"
	End  Select
		fn_Click orcParametersWindowObj.OracleButton(okbutton)
	If err.Number<> 0 Then
		fnReportEvent "Fail","Enter Parameters","Function name  : Enter Parameter,Unable to enter parameters for Request : " &RequestName ,true   
		print err.description
	End If
		
End Function

Function fn_ValidateOutput()
On error resume next
fn_ValidateOutput = false

If  gstrTdIdentifer2 = "MMC GLB AR BILLING AND RECEIPT HISTORY"  Then

	'Validation - specific to report - create separate fn
	'fn_Click orcRequestWindowObj.OracleButton(orcViewOutput)
	'fn_SelectMenu OracleFormWindow("short title:=Report"),"toolsCopyFile"
	If OracleFormWindow("short title:=Report").exist(5) Then
		OracleFormWindow("short title:=Report").SelectMenu "Tools->Copy File..."
		Browser("creation time:=2").Sync
		If Browser("creation time:=2").Exist(5) Then
			fnReportEvent "Pass","Detailed Report generated","Able to Launch Report in the browser" ,false			
			fn_ValidateOutput=true
		Else 
			fnReportEvent "Fail","Detailed Report not generated","Unable to Launch Report in the browser" ,true	
		End If		
	End If
Else 
	fnReportEvent "Fail","fn_ValidateOutput","Unable to Validate Output in the browser" ,false	
End  If 

If err.Number<> 0 Then
	fnReportEvent "Fail","Validate Output","Function name  : fn_ValidateOutput , Failed to Validate Output in Browser" ,true   
	print err.description
End If
End Function

Function fn_CheckRequestStatus(strReqId)
On error resume next
	If strReqId <> "" Then
'		OracleFormWindow("title:=Navigator.*").OracleTabbedRegion("label:=Functions").OracleList("description:=Function List").Activate "       View"
		'OracleFormWindow("title:=Navigator.*").SelectMenu "View->Requests"
		fn_SelectMenu orcHomePageNavigator,"viewrequests"
		orcFindRequestObj.OracleRadioGroup(orcAllReq).Select "Specific Requests"
		fn_ReportEnter orcFindRequestObj.OracleTextField("description:=Request ID"),strReqId,"Request ID"
		fn_Click orcFindRequestObj.OracleButton(orcFind)
		intLoopCnt = 1
               Do
                    fn_Click orcRequestWindowObj.OracleButton(orcRefreshData)
                    strPhaseStatus = orcRequestWindowObj.OracleTable(orctable).GetFieldValue(1,4)
                    If intLoopCnt = 100 Then
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
	print err.description
End If
End Function

Function fn_GetNumericValueFromString(str)
Dim c,a
	for x=1 to len(str)
	c=mid(str,x,1)
	If isnumeric(c) OR c="," then
		a=a&c
	End If 
	next
fn_GetNumericValueFromString = a
End Function

'Function fn_OperationalReporting()	'change name - to SubmitRequest
'
'On error resume next 
'fn_OperationalReporting=false
'If fn_exist(orcFindRequestObj)=true Then
'	fn_Click orcFindRequestObj.OracleButton(orcSubmitNewRequest)
'	fn_ReportEnter orcSubmitRequestObj.OracleTextField(orcRequestName),gb_TestDataDic.item("Request_Name"),"Request Name"
'	'fn_WSSendKeys TAB
'	'Call Parameterfn 
'	fn_ReportEnter orcParametersWindowObj.OracleTextField(orcOperatingUnit),gb_TestDataDic.item("Request_Operating_Unit"),"Operating Unit"
'	fn_ReportEnter orcParametersWindowObj.OracleTextField(orcSetOfBooks),gb_TestDataDic.item("Set_Of_Books"),"Set Of Books"
'	fn_ReportEnter orcParametersWindowObj.OracleTextField(orcTranscnDateLow),gb_TestDataDic.item("Transcn_Date_Low"),"Transaction Date Low"
'	fn_ReportEnter orcParametersWindowObj.OracleTextField(orcTranscnDateHigh),gb_TestDataDic.item("Transcn_Date_High"),"Transaction Date High"
'	fn_Click orcParametersWindowObj.OracleButton(okbutton)
'	
'	fn_Click orcSubmitRequestObj.OracleButton(orcSubmitBtn)
'	'code to capture request id
'	If orcDecisionNotificationObj.Exist(5) Then
'		strRequestNum = f_GetNumericValueFromString(orcDecisionNotificationObj.GetROProperty("message"))
'		fn_Click orcDecisionNotificationObj.OracleButton("label:=No")
'	End If
'
'	Call fn_CheckOperationalReportingRequestStatus(strRequestNum)
'	fnReportEvent "Pass","Submit Request","Function name  : fn_OperationalReporting , Successfully able to submit request : " &gb_TestDataDic.item("Request_Name")&  " Request Id is : "&strRequestNum,true   
'	
'	'Validation - specific to report - create separate fn
'	fn_Click orcRequestWindowObj.OracleButton(orcViewOutput)
'	'fn_SelectMenu OracleFormWindow("short title:=Report"),"toolsCopyFile"
'	If OracleFormWindow("short title:=Report").exist(5) Then
'		OracleFormWindow("short title:=Report").SelectMenu "Tools->Copy File..."
'		Browser("creation time:=2").Sync
'		If Browser("creation time:=2").Exist(5) Then
'			fnReportEvent "Pass","Detailed Report generated","Able to Launch Report in the browser" ,true   
'		End If		
'	End If
'    	'strRequestNum = Empty
'	fn_OperationalReporting=true
'Else 
'	fnReportEvent "Fail","Submit Request","Function name  : fn_OperationalReporting , Unable to submit request",true   
'End If
'
'If Err.number <> 0 Then             
'	fnReportEvent "Fail","Operational Reporting","Function name  : fn_OperationalReporting , Failed to submit request.Error is : " &Err.description,true   
'	Exit function
'End If
'	
'End Function
