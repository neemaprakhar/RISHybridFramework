
Public orcTranscnWindowObj,orcLinesWindowObj,orcDistributionWindowObject,orcTranscnWindowMainObj,orcFormLineItemDescObj,orcRuleAccntObj,orcContextValueObj,orcContextValueOkBtnObj,orcDistributionsBtnObj,orcTransactionNoObj,orcGLAccountObj,orcDistTableObj
Public orcCreditTranscnWindowObj,orcCreditMemoPercObj,orcCreditAllocationObj
Public orcReceiptWindowObj,orcRcptCustomerNum,orcRcptApplicationsTable,orcRcptApplicationsWindowObj,orcReceiptSummaryObj,orcReceiptSummaryWindowObj
Public orcRcptBatchWindowObj
Public mySendKeys,objRecSet

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
'Const orcBillTo = "description:=Bill To: Name"
Const orcPaymentTerm = "description:=Payment Term"
Const orcInvoicingRule = "description:=Invoicing Rule"
Const orcCompleteBtn = "description:=Complete"
Const orcIncompleteBtn = "description:=Incomplete"
Const orcCancellationReason = "description:=Reason"
Const orcCreditLines = "description:=Credit Lines"
Const orcReceiptMethod = "description:=Receipt Method"
Const orcReceiptNumber = "description:=Receipt Number"
Const orcReceiptAmount = "description:=Receipt Amount"
Const orcUnappliedAmount = "description:=Unapplied"
Const orcAppliedAmount = "description:=Applied"
Const orcBatchSource = "description:=Batch Source"
Const orcPaymentMethod = "description:=Payment Method"
Const orcRcptBtn = "description:=Receipts"
Const orcOpenBtn = "description:=Open"
Const orcApplyBtn = "description:=Apply"
Const orcNetReceiptAmount = "description:=Net Receipt Amount"
Const intRecord_no = 1 

Set orcTranscnWindowObj = OracleFormWindow("title:=Transactions.*")
Set orcLinesWindowObj = OracleFormWindow("title:=Lines.*")
Set orcDistributionWindowObject = OracleFormWindow("title:=Distributions.*")
Set orcCreditTranscnWindowObj =  OracleFormWindow("title:=Credit Transactions.*")
Set orcReceiptWindowObj = OracleFormWindow("title:=Receipts.*")
Set mySendKeys = CreateObject("WScript.shell")
Set orcTranscnWindowMainObj = orcTranscnWindowObj.OracleTabbedRegion("label:=Main")
Set orcBillToObj = orcTranscnWindowMainObj.OracleTextField("description:=Bill To: Name")
Set orcLineItemDescObj = orcLinesWindowObj.OracleTabbedRegion("label:=Main").OracleTable("block name:=Table")
Set orcRuleAccntObj = orcLinesWindowObj.OracleTabbedRegion("label:=Rules").OracleTable("block name:=Table")
Set orcContextValueObj = OracleFlexWindow("title:=Line Transaction Flexfield").OracleTextField("prompt:=Context Value")
Set orcContextValueOkBtnObj = OracleFlexWindow("title:=Line Transaction Flexfield").OracleButton("label:=OK")
Set orcDistributionsBtnObj = orcLinesWindowObj.OracleButton("description:=Distributions")
Set orcGLAccountObj = orcDistributionWindowObject.OracleList("class description:=popup list box")
Set orcDistTableObj = orcDistributionWindowObject.OracleTable("block name:=Table")
Set orcTransactionNoObj = OracleFormWindow("title:=Transactions.*").OracleTextField("description:=Number","tooltip:=Transaction Number")
Set orcCreditMemoPercObj = OracleFormWindow("title:=Credit Transactions.*").OracleTabbedRegion("label:=Transaction Amounts").OracleTextField("description:=Credit Memo: Line: Percent")
Set orcCreditAllocationObj = OracleFormWindow("title:=Credit Transactions.*").OracleTabbedRegion("label:=Transaction Amounts").OracleList("class description:=popup list box")
Set orcRcptCustomerNum = OracleFormWindow("title:=Receipts.*").OracleTabbedRegion("label:=Main").OracleTextField("description:=Number")
Set orcRcptApplicationsTable = OracleFormWindow("title:=Applications.*").OracleTable("block name:=Table")
Set orcRcptApplicationsWindowObj = OracleFormWindow("title:=Applications.*")
Set orcReceiptSummaryObj = OracleFormWindow("title:=Receipts Summary.*").OracleTable("block name:=Table")
Set orcReceiptSummaryWindowObj = OracleFormWindow("title:=Receipts Summary.*")
Set orcRcptBatchWindowObj = OracleFormWindow("title:=Receipt Batches.*")
'Adjusment Test Case Objects
Set OracleFormAdjustmentsMain = OracleFormWindow("title:=Adjustments.*").OracleTabbedRegion("label:=Main").OracleTable("block name:=Table")
Set OracleFormAdjstComments = OracleFormWindow("title:=Adjustments.*").OracleTabbedRegion("label:=Comments")
Set OracleFormAdjstStatus =  OracleFormWindow("title:=Adjustments.*").OracleTabbedRegion("label:=Comments").OracleTable("block name:=Table")
Set OracleFindAdjustmentNum = OracleFormWindow("title:=Find Adjustments*").OracleTabbedRegion("label:=Main").OracleTextField("description:=Adjustment Number")
Set OracleAdjustmentNumFindBtn = OracleFormWindow("title:=Find Adjustments").OracleButton("description:=Find")
Set OracleFormFuncationsList = OracleFormWindow("title:=Navigator.*").OracleTabbedRegion("label:=Functions").OracleList("description:=Function List")
Set OracleFormApproveAdjstTable = OracleFormWindow("title:=Approve Adjustments").OracleTable("block name:=Table")
Set OracleFormApprovalLimitsMainTable = OracleFormWindow("title:=Approval Limits.*").OracleTabbedRegion("label:=Main").OracleTable("block name:=Table")
Set OracleFormTransDistributions = OracleFormWindow("title:=Transactions.*").OracleButton("description:=Distributions")
Set ObjSubledgerJournalEntryPage = Browser("name:=Subledger Journal Entry.*").Page("title:=Subledger Journal Entry.*")
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

If fn_exist(orcReceiptWindowObj) = true Then
	fnReportEvent "Pass","Receipt Window Status","Successfully loaded Oracle Receipt Form Window",false   
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptMethod),gb_TestDataDic.item("Receipt_Method"),"Receipt Method"
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptNumber),gb_TestDataDic.item("Receipt_Number"),"Receipt Number"
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcNetReceiptAmount),gb_TestDataDic.item("Receipt_Amount"),"Net Receipt Amount"
	fn_ReportEnter orcRcptCustomerNum,gb_TestDataDic.item("Customer_Number"),"Customer Number"
	fn_Click orcReceiptWindowObj.OracleButton(orcApplyBtn)	
	fn_exist orcRcptApplicationsTable
	fn_EnterField orcRcptApplicationsTable,intRecord_no,"Apply To",gb_TestDataDic.item("Apply_To_Field"),"Apply To"
	fn_EnterField orcRcptApplicationsTable,intRecord_no,"Amount Applied",gb_TestDataDic.item("Amount_Applied"),"Amount Applied"
	'fn_EnterField orcRcptApplicationsTable,intRecord_no,"Activity",gb_TestDataDic.item("Activity"),"Activity"
	fn_Click OracleFormWindow("title:=Applications.*").OracleButton("label:=Refund Attributes")
	fn_ReportEnter OracleFormWindow("title:=Refund Attributes").OracleTextField("description:=Refund Payment Method"),gb_TestDataDic.item("Payment_Method"),"Payment Method"
	fn_Click OracleFormWindow("title:=Refund Attributes").OracleButton("description:=Apply")
	orcRcptApplications.SelectMenu "File->Save"
	fn_CloseWindow orcRcptApplicationsWindowObj
	
	vstrUnappliedAmount = orcReceiptWindowObj.OracleTextField(orcUnappliedAmount).GetROProperty("value")
	
	If vstrUnappliedAmount = gb_TestDataDic.item("UnApplied_Amount") Then
		 fnReportEvent "Pass","Balance Amount Check","UnApplied amount is displayed as expected. Amount is = "&vstrUnappliedAmount,false
		 fn_createRcpt_Refund_WriteOff=true  
	Else 
	  	fnReportEvent "Fail","Balance Amount Check","UnApplied amount is not displayed is as expected. Expected Amount is = "&vstrUnappliedAmount,true      
		fn_createRcpt_Refund_WriteOff=false                        
	End If
	
	vstrAppliedAmount = orcReceiptWindowObj.OracleTextField(orcAppliedAmount).GetROProperty("value")
	
	If vstrAppliedAmount = gb_TestDataDic.item("Amount_Applied")  Then
		fnReportEvent "Pass","Balance Amount Check","Applied amount is displayed as expected. Amount is = "&vstrAppliedAmount,false
		fn_createRcpt_Refund_WriteOff=true  
	Else 
		fnReportEvent "Fail","Balance Amount Check","Applied amount is not displayed is as expected. Expected Amount is = "&vstrAppliedAmount,true   
		fn_createRcpt_Refund_WriteOff=false                   
	End If
Else 	
		fnReportEvent "Fail","Receipt Window Status","Unable to Load Oracle Receipt Form Window",true   
		fn_createRcpt_Refund_WriteOff=false
		Exit function
End If

If Err.number <> 0 Then             
	print Err.description,true
	fn_createRcpt_Refund_WriteOff=false
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

If fn_exist (orcRcptBatchWindowObj) = true Then
	fnReportEvent "Pass","Receipt by Batch Window Status","Successfully loaded Receipt by Batch Form Window",false   
	fn_ReportEnter orcRcptBatchWindowObj.OracleTextField(orcBatchSource),gb_TestDataDic.item("Batch_Source"),"Batch Source"	
	fn_ReportEnter orcRcptBatchWindowObj.OracleTextField(orcPaymentMethod),gb_TestDataDic.item("Payment_Method"),"Payment Method"
	fn_ReportEnter OracleFormWindow("title:=Receipt Batches.*").OracleTextField("description:=Totals: Control Count"),gb_TestDataDic.item("Total_Count"),"Count"
	fn_ReportEnter OracleFormWindow("title:=Receipt Batches.*").OracleTextField("description:=Totals: Control Amount"),gb_TestDataDic.item("Total_Amount"),"Amount"	
	fn_Click orcRcptBatchWindowObj.OracleButton(orcRcptBtn)
	fn_EnterField orcReceiptSummaryObj,intRecord_no,"Receipt Number",gb_TestDataDic.item("Receipt_No"),"Receipt No"
	fn_EnterField orcReceiptSummaryObj,intRecord_no,"Net Amount",gb_TestDataDic.item("Net_Amount"),"Net Amount"
	fn_Click orcReceiptSummaryWindowObj.OracleButton(orcOpenBtn)
	OracleFormWindow("Receipts").OracleTabbedRegion("Main").OracleTextField("Detail|Customer|Number").Enter "CITI81"
	fn_Click OracleFormWindow("Receipts").OracleButton("Apply")
	fn_EnterField orcRcptApplicationsTable,intRecord_no,"Apply To",gb_TestDataDic.item("Apply_To_TranscnNo"),"Apply To:Transcn No"
	OracleFormApplicationWindowObj.SelectMenu "File->Save"
	fn_CloseWindow orcRcptApplicationsWindowObj
	fn_CloseWindow OracleFormWindow("Receipts")
	fn_CloseWindow orcReceiptSummaryWindowObj
	
	vstrActualAmountCheck = OracleFormWindow("title:=Receipt Batches.*").OracleTextField("description:=Totals: Actual Amount").GetROProperty("value")
	
	If vstrActualAmountCheck = gb_TestDataDic.item("Actual_Value") Then
		fnReportEvent "Pass","Actual Amount Check","Actual Amount & Receipt Amount are matching and Amount is = "&vstrActualAmountCheck,false
		fn_ApplyReceiptByBatch=true
	Else 
		fnReportEvent "Fail","Actual Amount Check : Fail", "Actual Amount & Receipt Amount are not matching. Expected Amount is = "&vstrActualAmountCheck,true
		fn_ApplyReceiptByBatch=false
	End If	
	
	vstrDifferenceAmountCheck = OracleFormWindow("title:=Receipt Batches.*").OracleTextField("description:=Totals: Difference Amount").GetROProperty("value")
	
	If vstrDifferenceAmountCheck = gb_TestDataDic.item("Difference")  Then
		fnReportEvent "Pass","Difference Amount Check","Difference Amount is displayed as expected & Amount is = "&vstrDifferenceAmountCheck,false
		fn_ApplyReceiptByBatch=true
	Else 
		fnReportEvent "Fail","Difference Amount Check : Fail", "Difference Amount is not displayed as expected. Expected Amount is = "&vstrDifferenceAmountCheck,true
		fn_ApplyReceiptByBatch=false
	End If	
Else 	
	fnReportEvent "Fail","Receipt by Batch Window Status","Unable to Load Receipt by Batch Form Window",true   
	fn_ApplyReceiptByBatch=false
	Exit function
End If	

If Err.number <> 0 Then             
 	print Err.description,true
 	fn_ApplyReceiptByBatch=false
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

If fn_exist(orcReceiptWindowObj) = true Then
	fnReportEvent "Pass","Receipt Window Status","Successfully loaded Oracle Receipt Form Window",false   
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptMethod),gb_TestDataDic.item("Receipt_Method"),"Receipt Method"
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptNumber),gb_TestDataDic.item("Receipt_Number"),"Receipt Number"	
	fn_ReportEnter orcReceiptWindowObj.OracleTextField(orcReceiptAmount),gb_TestDataDic.item("Receipt_Amount"),"Receipt Amount"	
	fn_ReportEnter orcRcptCustomerNum,gb_TestDataDic.item("Customer_Number"),"Customer Number"	
	fn_Click OracleFormWindow("title:=Receipts.*").OracleButton("description:=Apply")	
	'fn_EnterField orcRcptApplicationsTable,intRecord_no,"Apply To",gb_TestDataDic.item("Apply_To"),"Apply To"	
	fn_EnterField orcReceiptSummaryObj,intRecord_no,"Net Amount",gb_TestDataDic.item("Amount_Applied"),"Amount Applied"
	orcRcptApplicationsWindowObj.SelectMenu "File->Save"	
	fn_CloseWindow orcReceiptSummaryWindowObj	
	
	vstrUnappliedAmount = orcReceiptWindowObj.OracleTextField(orcUnappliedAmount).GetROProperty("value")
	
	If vstrUnappliedAmount = gb_TestDataDic.item("Amount_Applied") Then
	         fnReportEvent "Pass","Balance Amount Check","UnApplied amount is displayed as expected",false
	         fn_createReceipt=true
	Else 
	         fnReportEvent "Fail","Balance Amount Check","UnApplied amount is not displayed is as expected",true       
	         fn_createReceipt=false
	End If
Else 	
	fnReportEvent "Fail","Receipt Window Status","Unable to Load Oracle Receipt Form Window",true   
	fn_createReceipt=false
	Exit function
End If

If Err.number <> 0 Then             
	print Err.description,true
	fn_createReceipt=false
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
'fn_exist orcTranscnWindowObj.OracleTextField(orcSource)
If (fn_exist(orcTranscnWindowObj.OracleTextField(orcSource)))  Then
	fnReportEvent "Pass","Transaction Window Status","Successfully loaded Transaction Form Window",false 
	fn_ReportEnter orcTranscnWindowObj.OracleTextField(orcSource),gb_TestDataDic.item("Source_Field"),"Source Field"	
	fn_Select orcTranscnWindowObj.OracleList(orcClass),gb_TestDataDic.item("Class"),"Class"
	fn_ReportEnter orcTranscnWindowObj.OracleTextField(orcCurrency),gb_TestDataDic.item("Currency"),"Currency"
	fn_ReportEnter orcTranscnWindowObj.OracleTextField(orcType),gb_TestDataDic.item("Type"),"Type"
	fn_Enter orcBillToObj,gb_TestDataDic.item("Bill_To")
	fn_WSSendKeys TAB
	fn_Enter orcTranscnWindowMainObj.OracleTextField(orcPaymentTerm),gb_TestDataDic.item("Payment_Term"),"Payment Term"
	fn_Select orcTranscnWindowMainObj.OracleList(orcInvoicingRule),gb_TestDataDic.item("Invoice_Rule"),"Invoice Rule"
	fn_Click orcTranscnWindowObj.OracleButton(orcLineItems)
	fn_Click OracleNotification("title:=Error").OracleButton("label:=OK")
	fn_EnterField orcLineItemDescObj,intRecord_no,"Description",gb_TestDataDic.item("Description_value"),"Description"
	fn_EnterField orcLineItemDescObj,intRecord_no,"Quantity",gb_TestDataDic.item("Quantity_Value"),"Quantity"
	fn_EnterField orcLineItemDescObj,intRecord_no,"Unit Price",gb_TestDataDic.item("Unit_Price"),"Unit Price"
	fn_EnterField orcLineItemDescObj,intRecord_no,"Tax Classification",gb_TestDataDic.item("Tax_Classification"),"Tax Classification"
	fn_Enter orcContextValueObj,gb_TestDataDic.item("Operating_Unit")
	fn_Click orcContextValueOkBtnObj
	fn_WSSendKeys TAB3
	fn_EnterField orcRuleAccntObj,intRecord_no,"Accounting",gb_TestDataDic.item("Accounting"),"Accounting"
	fn_Click orcDistributionsBtnObj	
	fn_Select orcGLAccountObj,gb_TestDataDic.item("GLAccountSelect"),"GL Account Selection"
	
	Call fn_selectDistributions
	
	orcDistWindowObj.SelectMenu "File->Save"
	fn_CloseWindow orcDistWindowObj
	fn_CloseWindow orcLinesWindowObj
	
	Call fn_CaptureTransactionNo
	fn_CreateStdInv_CreditMemo=true
	'Code to close oracle form
	'Code to call Switch Responsibility Function 
Else 	
	fnReportEvent "Fail","Transaction Window Status","Unable to Load Transaction Form Window",true   
	fn_CreateStdInv_CreditMemo=false
Exit function
End If

If Err.number <> 0 Then             
	print Err.description,true
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
			If glAccountInfo="" Then
				fnReportEvent "Fail","GL Account Status","GL Account field is empty",true   
			End If
		Else 
	Exit For
		End If
	Next
		
Else 
	fnReportEvent "Fail","Distributions Table Window Status","Unable to Load Distributions Table Window",true   
End If

If Err.number <> 0 Then             
	print Err.description,true
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
	Exit function
End If
End Function

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
'=============================================================
Function fn_ApproveTransaction()				'Responsibility Used - AR Transaction Approver 
On error resume next

If fn_Exist(orcTranscnWindowObj) = true Then
	fnReportEvent "Pass","Transaction Window Status","Successfully loaded Transaction Form Window",false 
	orcTranscnWindowObj.SelectMenu "View->Query By Example->Enter"
	fn_Exist orcTranscnWindowObj.OracleTextField(orcSource)
	fn_ReportEnter orcTranscnWindowObj.OracleTextField(orcSource),gb_TestDataDic.item("Source_Field"),"Source Field"
	'varTxnNo = fn_ReadTransactionNo
	'fn_ReportEnter orcTransactionNoObj,varTxnNo,"Transaction No"
	vtransactionnumber = fn_getExecutionResultData(gstrTestCaseExec_id,"Transaction_No")
	
	fn_ReportEnter orcTransactionNoObj,gb_TestDataDic.item("Transaction_No"),"Transaction No"
	orcTranscnWindowObj.SelectMenu "View->Query By Example->Run"
	fn_Click orcTranscnWindowObj.OracleButton(orcCompleteBtn) 
	
	vstrStatusCheck = orcTranscnWindowObj.OracleButton(orcIncompleteBtn).GetROProperty("description")
	
	If vstrStatusCheck="Incomplete" Then
		fnReportEvent "Pass","Transaction Approval Status", "Transaction is completed/approved successfully",false
		fn_ApproveTransaction=true
	Else 
		fnReportEvent "Fail","Transaction Approval Status", "Unable to complete/approve Transaction successfully",true
		fn_ApproveTransaction=false
	End If
	
	Call fn_CaptureTransactionNo
Else 	
	fnReportEvent "Fail","Transaction Window Status","Unable to Load Transaction Form Window",true   
	fn_ApproveTransaction=false
	Exit function
End If

If Err.number <> 0 Then             
	print Err.description,true
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

If fn_exist(orcTranscnWindowObj) = true Then
	fnReportEvent "Pass","Transaction Window Status","Successfully loaded Transaction Form Window",false 
	'fn_exist orcTranscnWindowObj.OracleTextField(orcSource)
	orcTranscnWindowObj.SelectMenu "View->Query By Example->Enter"
	'fn_ReportEnter orcTranscnWindowObj.OracleTextField(orcSource),gb_TestDataDic.item("Source_Field"),"Source Field"
	'varTxnNo = fn_ReadTransactionNo
	'fn_ReportEnter orcTransactionNoObj,varTxnNo,"Transaction No"
	fn_ReportEnter orcTransactionNoObj,gb_TestDataDic.item("Transaction_No"),"Transaction No"
	orcTranscnWindowObj.SelectMenu "View->Query By Example->Run"
	orcTranscnWindowObj.SelectMenu "Actions->Credit"
	fn_ReportEnter orcCreditTranscnWindowObj.OracleTextField(orcCancellationReason),gb_TestDataDic.item("Cancellation_Reason"),"Cancellation Reason"
	fn_Select orcCreditAllocationObj,gb_TestDataDic.item("Credit_Allocation"),"Credit Allocation"
	fn_Enter orcCreditMemoPercObj,gb_TestDataDic.item("Percentage")
	fn_Click orcCreditTranscnWindowObj.OracleButton(orcCreditLines)
	fn_Click orcDistributionsBtnObj
	fn_exist orcGLAccountObj
	fn_Select orcGLAccountObj,gb_TestDataDic.item("GLAccountSelect"),"Accounts For All Lines"
	
	vstrReceiveableAmountCheck = orcDistTableObj.GetFieldValue (1,"Distribution Amount")
	If vstrReceiveableAmountCheck <= 0  Then
		fnReportEvent "Pass","Receiveable Amount Check","Amount is Negative",false
	Else 
		fnReportEvent "Fail","Receiveable Amount Check", "Amount is not negative",true
	End If	
	
	vstrRevenueAmountCheck = orcDistTableObj.GetFieldValue (2,"Distribution Amount")
	If vstrReceiveableAmountCheck <= 0  Then
		fnReportEvent "Pass","Revenue Amount Check","Amount is Negative",false
	Else 
		fnReportEvent "Fail","Revenue Amount Check", "Amount is not negative",true
	End If	
	
	fn_CreditNote=true
Else 	
	fnReportEvent "Fail","Transaction Window Status","Unable to Load Transaction Form Window",true   
	fn_CreditNote=false
	Exit function
End If

If Err.number <> 0 Then             
	fnReportEvent "Fail","CreditNote Creation","Functiona name  : fn_CreditNote , Fail to create the credit note ",true   
	fn_CreditNote=false
	Exit function
End If
End Function

