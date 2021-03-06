Public OracleIProcPageObj,OracleIProcPageObj_1,OracleIProcPageObj_2,OracleIProcPageObj_3
Public approvalList()
Public Count_appr
Public AutoCreatedPONumber
Set OracleIProcPageObj = Browser("name:=Oracle iProcurement: Shop").Page("title:=Oracle iProcurement: Shop")
Set OracleIProcPageObj_1 =  Browser("name:=Oracle iProcurement: Checkout").Page("title:=Oracle iProcurement: Checkout")
Set OracleIProcPageObj_2= Browser("name:=Approval Group").Page("title:=Approval Group")
Set OracleIProcPageObj_3 = Browser("name:=Confirmation").Page("title:=Confirmation")
Set OracleIProcPageObj_4 = Browser("name:=Oracle iProcurement: Requisitions").Page("title:=Oracle iProcurement: Requisitions")
Set OracleIProcPageObj_ChangeOrder = Browser("name:=Change Order.*").Page("title:=Change Order.*")
Set OracleIProcPageObj_ChangeOrderApprovals = Browser("name:=Change Order: Select Approvals").Page("title:=Change Order: Select Approvals")
Set objApproverDict = CreateObject("Scripting.Dictionary")
Set wflowconfigobj = Browser("name:=Workflow Configuration").Page("title:=Workflow Configuration")
Set notificationsobj = Browser("name:=Notifications").Page("title:=Notifications")
Set notificDetailsobj = Browser("name:=Notification Details").Page("title:=Notification Details")
Set OracleIProcPageObj_Receiving = Browser("name:=Oracle iProcurement: Receiving").Page("title:=Oracle iProcurement: Receiving")
Set OracleIProcPageObj_ReturnItems = Browser("name:=Oracle iProcurement: Return Items").Page("title:=Oracle iProcurement: Return Items")
Set OracleIProcPageObj_SearchSelectLOV = Browser("name:=Search and Select List of Values.*").Page("title:=Search and Select List of Values.*")

'=============================================================
'*************************************************************************
'Iproc WebElemnts XPaths
'=============================================================
'*************************************************************************
Const NonCatalaogLink_xpath= "xpath:=//a[@title='Non-Catalog Request']"
Const clearAll_xpath="xpath:=//button[@id='ClearAll']"
Const addToCart_xpath="xpath:=.//*[@id='AddToCart']"
const itemDesc_xpath="xpath:=.//*[@id='ItemDescription']"
Const category_xpath="xpath:=//input[@id='Category']"
Const quantity_xpath="xpath:=//input[@id='Quantity']"
Const unitOfMeasure_xpath="xpath:=//input[@id='UnitOfMeasureTl']"
Const unitPrice_xpath="xpath:=//input[@id='UnitPrice']"
Const supplier_xpath="xpath:=//input[@id='SupplierOnNonCat']"
Const supplierSite_xpath="xpath:=//input[@id='SupplierSiteOnNonCat']"
Const checkoutBtn_xpath="xpath:=.//*[@id='Checkout']"
Const projectSrc_xpath="xpath:=.//*[@id='N3:DescFlex2:0']"
Const projectCode_xpath="xpath:=.//*[@id='N3:DescFlex3:0']"
Const nextBtn1_xpath="xpath:=.//table[@id='PageButtonsRN']/tbody/tr/td[12]/button[1]"
Const requesterApproverLink_xpath="xpath:=.//a[contains(text(),""Requester's Supervisor"")]"
Const commodityApproverLink_xpath="xpath:=.//a[contains(text(),""Commodity Approvers"")]"
Const nextBtn2_xpath="xpath:=//TABLE[@id='PageActionButtonsBar']/TBODY/TR/TD[10]/BUTTON[1]"
Const reqSubmitBtn_xpath="xpath:=.//*[@id='SubmitButton']"
Const confirmationText_xpath="xpath:=.//*[@id='ApproverText']"
Const reqText_xpath="xpath:=.//*[@id='ApproverText']/b[1]"
Const approverDetailTable_xpath="xpath:=.//SPAN[@id=""GroupHeader.ApprovalGroupTable""]/TABLE[2]"
Const returnBtn_xpath="xpath:=.//*[@id='ReturnButton']"
Const showDetails_xpath = "xpath:=//*[@id='N6dd0']/img"
Const chargeAccountNum = "title:=Charge Account"
Const approvalPagexpath = "xpath:=//h1[text()='Checkout: Requisition Information']" 
Const checkoutTable = "xpath:=//SPAN[@id='ItemTableRN']/TABLE[2]"
Const editlines_btn = "//BUTTON[@id='EditLines']"
Const currencydropdown = "xpath:=//*[@id='Currency']"
Const exchangeCurrAmount = "xpath:=//*[@id='N6:Amount:0']"
Const approvalListError = "xpath:=//*[@id='FwkErrorBeanId']"
Const GoButton = "xpath:=//*[@id='Go']"
Const requisitionTabXpath = "xpath:=//*[@name='ICXPOR_REQSTATUS']"
Const changeButton = "xpath:=//button[@id='Change']"
Const changeOrderError = "xpath:=//*[@id='FwkErrorBeanId']"
Const reasonTextXpath = "xpath:=//textarea[@title='Reason']"
Const updatedPOPriceTextBoxXpath = "xpath:=//*[@title='Price']"
Const changeOrderNextButton = "xpath:=//*[@id='NavCell']/table/tbody/tr/td[3]/button"
Const requestorSupervisorXpath = "xpath:=//*[text()='Requester's Supervisor']"
Const selectApprNextButton = "xpath:=//*[@id='NavCell']/table/tbody/tr/td[5]/button"
Const changeOrderConfirmation = "xpath:=//h1[text()='Confirmation']"
Const officeSuppliesXpath = "xpath:=//a[contains(text(),'Office Supplies')]"
Const needToBuyGoodsXpath = "xpath:=//a[contains(text(),'I Need to Buy (Goods)')]"
Const foreignReqStatus = "xpath:=//*[@id='N19:ApprovalStatus:0']"
Const changeOrderApprText = "xpath:=//*[@id='ApprListText']"
Const receivingtab = "xpath:=//a[@name='ICXPOR_RECEIVING_HOME']"
Const receiveItemsLink = "xpath:= //a[@id='ReceiveItemsLink1']"
Const reqNumbertextbox = "xpath:=//input[@id='ReqNumber']"
Const goButtonSearchReq = "xpath:=//button[text()='Go']"
Const returntItemsLink = "xpath:=//a[@id='ReturnItemsLink1']"
Const receiptQuantity = "xpath:=//input[@id='N3:ReceiptQuantity:0']"
Const returnQuantity = "xpath:=//input[@id='N3:ReturnQuantity:0']"
Const selectReceivingReq = "xpath://*[@name='N3:selected:0']"
Const fromUserListOfValues = "xpath:=//*[@id='wfUserName3__xc_']/a"
Const SelectSearchBy = "xpath:=//select[@id='CategoryChoice']"
Const enterUserId = "//*[@id='categoryChoice']/following-sibling::input"
Const GoButtonToSearchUserID = "xpath:=//*[@id='categoryChoice']/following-sibling::button]"
Const quickSelectSearchValue = "xpath:=//*[@title='Quick Select']"
Const submitReceiving = "//button[@id='SubmitButton_uixr']"
Const FullyReceiveNextButton = "xpath:=//*[@id='NavigationCell']/table/tbody/tr/td[3]/button"
Const FullyReceiveNextNextButton = "xpath:=//*[@id='NavigationCell']/table/tbody/tr/td[5]/button"
Const returnReason = "xpath:=//*[@id='Reason']"
Const receievReqItemsDue = "xpath:=//*[@id='ItemsDue']"
Const homePagelinkXpath = "xpath:=//*[@class='x6w']//td[3]/a"
Const homePageLinkAdmin = "xpath:=//*[@class='x6w']//td[2]/a"
'===============================================================

'=============================================================
'*************************************************************************
'Iproc Administration XPaths
'=============================================================
'*************************************************************************
Const WfNotificationSearch = "xpath:=.//*[@id='WF_WORKLIST_SEARCH']"
Const frmNotification = "xpath:=.//*[@id='wfUserName2']"
Const Subject = "xpath:=.//*[@id='Subject']"
Const admin_goButton = "xpath:=//button[@id='Go']"
Const searchedPurchaseReqLink = "xpath:=.//*[contains(text(),'Purchase Requisition')]"
Const approveButton = "xpath:=(//button[@title='Approve'])[1]"

'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_RequisitionCreation 
'description:             fn_RequisitionCreation :  will create a new requisition and submit
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_RequisitionCreation()
 blnResultFlag=false
 On error resume next
    OracleIProcPageObj.Sync
       If OracleIProcPageObj.Link(NonCatalaogLink_xpath).Exists(5) Then
        fnReportEvent "Pass", "Oracle Iproc page navigation status","Successfully navigated to Oracle Iproc Home page ",false
        If OracleIProcPageObj.WebButton(checkoutBtn_xpath).Exist(10) then
        	Call fn_deleteCart
          	blnResultflag=true
        End if
    Else
        fnReportEvent "Fail","Oracle Iproc page navigation status"," Failed to navigate Oracle Iproc Home page",true
        Exit Function
    End If  

	Call fn_NonCatlogAddToCart()
 If fn_enterAndCheckoutNCRequest Then
		call fn_vrfyapproversSubmitReq 
		blnResultFlag=fn_verifyConfirmation  	
 End If 
       If Err.number <> 0 Then             
             fnReportEvent "Fail","Requisition Creation Status","Fail to create the requisition" & Err.description,false
             fn_RequisitionCreation = false
             Exit function
      End If
       fn_RequisitionCreation=blnResultFlag             
End Function
'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_NonCatlogAddToCart 
'description:         fn_NonCatlogAddToCart :  will navigate to Not catalog Request page and enter the values 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_NonCatlogAddToCart()
blnResultFlag=false
If Not(fn_Click(OracleIProcPageObj.Link(NonCatalaogLink_xpath)))Then
    fnReportEvent "Fail","Non-Catalog Request Link","Either Non-Catalog Request link not found in the Oracle iProcurement: Shop screen or could not reach at the screen.",True
    Exit Function
else
    blnResultFlag=true
End If
fn_Click OracleIProcPageObj.WebButton(clearAll_xpath)
fn_NonCatlogAddToCart=blnResultFlag
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_AddtoCart 
'description         fn_AddtoCart: Will click on the Add to cart button
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_AddtoCart()
	call fn_Click_fieldname(OracleIProcPageObj.WebButton(addToCart_xpath),"AddToCart Button")
End Function


'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_selectBillableOrNonbillable 
'description:         fn_selectBillableOrNonbillable :  select the billable and non billable as per the Business unit 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_selectBillableOrNonbillable()
On Error Resume Next
    If  gb_TestDataDic("Business_Line") = "Mercer" OR  gb_TestDataDic("Business_Line") = "OWG" OR gb_TestDataDic("Business_Line") = "NERA" Then

	If OracleIProcPageObj_1.WebEdit(projectSrc_xpath).Exist(5) Then
		
		totalRowCount = OracleIProcPageObj_1.Webtable(checkoutTable).GetROProperty("Rows")
		
		For irow = 1 To totalRowCount-2
			call fn_SetCheckoutTable(irow,"Project Source",gb_TestDataDic.item("Project_source"))
			call fn_SetCheckoutTable(irow,"Project Code",gb_TestDataDic.item("Project_code"))
		Next
		
		result = fn_ClickCheckOutButton
		fn_selectBillableOrNonbillable = result
    Else
		Call fn_ClickCheckOutButton
		fn_selectBillableOrNonbillable = True
    End If
    
    Else
     fnReportEvent "Fail","Project Source EditBox","Project Source EditBox doesnt Exist",True  
  End If
	If Err.number <> 0 Then             
              fnReportEvent "Fail","billable/non billable/Checkout","Failed to check out " & Err.description,false
             fn_selectBillableOrNonbillable = false             
      End If
End Function
'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_verifyApprovers 
'description:             fn_verifyApprovers :  verify whether Approval list is present for requisition.
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_verifyApprovers()
On Error Resume Next
Set objApproverDict = CreateObject("Scripting.Dictionary")

    If fn_Click(OracleIProcPageObj_1.WebButton(nextBtn1_xpath)) Then        
        fnReportEvent "Pass","Next Button","Succesfully Click on the Next Button",False
    Else
        fnReportEvent "Fail","Next Button","Next Button doesnt Exist",True    
    End  If
    
If fn_getApprovalHierarchy=False Then
	fn_verifyApprovers = False
	Exit Function
Else 
	fn_verifyApprovers = True	
End If

'Call fn_getApprovalHierarchy()
 
	call fn_Click_fieldname(OracleIProcPageObj_1.WebButton(nextBtn2_xpath),"Next button")

	If Err.number <> 0 Then             
              fnReportEvent "Fail","verify Approvers","Failed to verify approvers " & Err.description,false
             fn_verifyApprovers = false
      End If
      
End Function
'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_SubmitRequisition 
'description:             fn_SubmitRequisition :  Submit the requisition
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_SubmitRequisition()
On Error Resume Next
    If OracleIProcPageObj_1.WebButton(reqSubmitBtn_xpath).Exist(2) Then        		
		call fn_Click(OracleIProcPageObj_1.WebButton(reqSubmitBtn_xpath))
		fnReportEvent "Pass","Submit Button","Succesfully submit the  requisition request",False
		fn_SubmitRequisition = true
    Else           
		fnReportEvent "Fail","Submit Button","Submit Button doesnt Exist " & Err.description,false
		fn_SubmitRequisition = false		
	End If
	
End Function
'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_verifyConfirmation 
'description:             fn_verifyConfirmation :  will verify the req no and confirmation .
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_verifyConfirmation()
On error resume next 
blnResultFlag=false

    strTest="has been submitted."
    var_Reqno=OracleIProcPageObj_3.WebElement(reqText_xpath).GetROProperty("innertext")
    ReqNo=split(var_Reqno," ")
    'print ReqNo(1)
    
    strQuery="UPDATE [ExecutionResult$] SET Requisition_No='"&ReqNo(1)&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"

    'print strQuery
    Call fn_updateQuery(strQuery)
    
    strReqConfirmStmt=OracleIProcPageObj_3.WebElement(confirmationText_xpath).GetROProperty("innertext")
    If InStr(strReqConfirmStmt,strTest) >0 Then
	    print ("Confirmation found.Requisition no "& ReqNo(1) &" has been submitted.")
	    fnReportEvent "Pass","Confirmation","Confirmation found.Requisition no "& ReqNo(1) &" has been submitted",True
	    blnResultFlag=true
    else
	    fnReportEvent "Fail","Confirmation","Confirmation not found",True
	    blnResultFlag=false
End If
	fn_Click_fieldname OracleIProcPageObj_3.WebElement(homePagelinkXpath),"Home Page Link"
	fn_verifyConfirmation = blnResultFlag
  	
  If Err.number <> 0 Then             
             fnReportEvent "Fail","Verify Confirmation","Failed to the confirmation of submitted requisition." & Err.description,false
             fn_verifyConfirmation = false
  End If 
    	
End Function

Function fn_ClickCheckOutButton()
Set OracleIProcPageObj= Browser("name:=Oracle iProcurement.*").Page("title:=Oracle iProcurement.*")
	If OracleIProcPageObj.WebButton(checkoutBtn_xpath).Exist(3) Then        
             fn_Click OracleIProcPageObj.WebButton(checkoutBtn_xpath)
             fnReportEvent "Pass","Checkout Button","Click on the Checkout Button ",False
             fn_ClickCheckOutButton=true
        Else
            fnReportEvent "Fail","Checkout Button","Checkout Button doesnt Exist",True
            fn_ClickCheckOutButton=false
        End If  
End Function

Function fn_deleteCart()
		
			OracleIProcPageObj.WebButton(checkoutBtn_xpath).Click	
			' Added the following code to handle the error
			OracleIProcPageObj_1.Sync
			
			If  OracleIProcPageObj_1.Exist(15) Then
				strRowCount=OracleIProcPageObj_1.WebTable("xpath:=.//*[@id='ItemTableRN']/table[2]").GetROProperty("rows")
			
				If strRowCount>2  Then
				   For i=2 to strRowCount-1
						OracleIProcPageObj_1.Image("xpath:=//SPAN[@id='ItemTableRN']//td[19]//a/img").Click
						fnReportEvent "Pass","delete button","Clicked on delete button.",False
				   Next
				
				Else
               			fnReportEvent "Fail","delete button","delete button not available",True
		      		 Exit Function
		       End If
			End If
			OracleIProcPageObj_1.Link("xpath:=.//*[@id='ReturnToShoppingLink']").Click		
			OracleIProcPageObj.Sync
			OracleIProcPageObj.Link(NonCatalaogLink_xpath).Click
			Call fnNonCatlogAddToCart()
							
				

End Function

Function fn_getApprovalHierarchy()
On error resume Next

Dim ObjLinks, ObjChild,intCount

Set ObjLinks = Description.Create
ObjLinks("html tag").Value="A"

Set ObjChild = OracleIProcPageObj_1.WebTable("html id:=ApproverListRN").ChildObjects(ObjLinks)
print ObjChild.Count
Dim Approval_Levels(3)
'Dim arrayObjForList(5)

'This Condition is Specific to  test case related to validation of Approval flow hierarachy (GSI.A2R.IP.03.003)
		   If len(gb_TestDataDic.item("Approver_1")) > 1 Then
                 
		          For i = 0 To ObjChild.Count - 1
		                 Approval_Levels(i) = ObjChild(i).GetROProperty("innertext")
	                 Next
			If fn_VerifyApproverList(Approval_Levels)=False Then
				fn_getApprovalHierarchy = False
				Exit Function
			Else
				fn_getApprovalHierarchy = True
			End If
                End If


ReDim preserve approvalList(ObjChild.Count - 1)

		For j = 0 To ObjChild.Count - 1
                	approvalList(j) = ObjChild(j).GetROProperty("innertext")
                Next
                   
        If fn_putApproverNametoTestData(approvalList) = False Then
        	fn_getApprovalHierarchy = False
        	Exit Function
        Else
        	fn_getApprovalHierarchy = True
        End If
  If Err.number <> 0 Then             
             fnReportEvent "Fail","Function ::fn_getApprovalHierarchy","Failed to fetch the Approval Hierarchy from the UI" & Err.description,false
             fn_getApprovalHierarchy = false
  End If 

End function

'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_putApproverNametoTestData
'description:             fn_putApproverNametoTestData : will update approval name in the execution result 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================


Function fn_putApproverNametoTestData(approvalList)

On error resume next
blnResultCounter = 1
Count_appr = 1
If OracleIProcPageObj_1.Exist(5) Then

	For k=0 To UBound(approvalList)
		arrLinkName = approvalList(k)
		appXpath="xpath:=.//a[contains(text(),"&chr(34) &arrLinkName& chr(34)&")]"
		OracleIProcPageObj_1.Link(appXpath).Click
		If OracleIProcPageObj_2.Exist(10) Then
			'Update Constants for constant xpath
			var_Row=OracleIProcPageObj_2.WebTable("xpath:=.//SPAN[@id=""GroupHeader.ApprovalGroupTable""]/TABLE[2]").GetROProperty("rows")
			lastrowcounter = var_Row
				For i=2 to var_Row
					var_ApproverName=OracleIProcPageObj_2.WebTable("xpath:=.//SPAN[@id=""GroupHeader.ApprovalGroupTable""]/TABLE[2]").GetCellData(lastrowcounter,1)
		               	 strQuery="UPDATE [ExecutionResult$] SET Approver"&Count_appr&"='"&var_ApproverName&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"
					Count_appr = Count_appr + 1				
					print strQuery    					
						If var_ApproverName <>""  Then
							Call fn_updateQuery(strQuery)
							fnReportEvent "Pass",""&arrLinkName&"",""&arrLinkName&" exists",False	
							blnResultCounter = 0
						Else
							blnResultCounter = blnResultCounter + 1
							fnReportEvent "Fail","Approval Name","Failed to find the Approval Name.",False						
						End If
					lastrowcounter = lastrowcounter - 1				
				Next
'					OracleIProcPageObj_2.WebButton("xpath:=.//*[@id='ReturnButton']").Click
					fn_Click OracleIProcPageObj_2.WebButton("xpath:=.//*[@id='ReturnButton']")
				Else
					fnReportEvent "Fail",""&arrLinkName&"",""&arrLinkName&" not available",True
					Exit Function
					arrLinkName = Null		
		End If
		
		Next
End  If

If blnResultCounter = 0 Then
	fn_putApproverNametoTestData = True
Else
	fn_putApproverNametoTestData = False
End If

  If Err.number <> 0 Then             
             fnReportEvent "Fail","Putting Approver Name to Test Data","Failed to write the approver name to Test Data file." & Err.description,false
             fn_putApproverNametoTestData = false
             Exit function
      End If
End Function



'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_RequisitionCreationGlobalAccValidation
'description:             fn_RequisitionCreationGlobalAccValidation :  will create a new requisition will vaildate the global account and submit
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_RequisitionCreationGlobalAccValidation()
 blnResultFlag=false
 On error resume next
    OracleIProcPageObj.Sync
       If OracleIProcPageObj.Link(NonCatalaogLink_xpath).Exists(5) Then
        fnReportEvent "Pass", "Oracle Iproc page navigation status","Successfully navigated to Oracle Iproc Home page ",false
        If OracleIProcPageObj.WebButton(checkoutBtn_xpath).Exist(10) then
        	Call fn_deleteCart
          	blnResultflag=true
        End if
    Else
        fnReportEvent "Fail","Oracle Iproc page navigation status"," Failed to navigate Oracle Iproc Home page",true
        Exit Function
    End If  

	Call fn_NonCatlogAddToCart()
 If fn_enterValuesonNonCatlogRequestPage Then
		Call fn_selectBillableOrNonbillable
		Call fn_verifyApprovers
		Call fn_GlobalAccountValidation
		Call fn_SubmitRequisition
		blnResultFlag=fn_verifyConfirmation  	
 End If 
       If Err.number <> 0 Then             
             fnReportEvent "Fail","Requisition Creation Status","Fail to create the requisition" & Err.description,false
             fn_RequisitionCreation = false
             Exit function
      End If
       fn_RequisitionCreation=blnResultFlag             
End Function

Function fn_validateApprovalSubmitReq()
		Call fn_verifyApprovers
		Call fn_SubmitRequisition
End Function


'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_enterAndCheckoutNCRequest
'description:             fn_enterAndCheckoutNCRequest :  will fill the requistion details
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_enterAndCheckoutNCRequest()
blnResultFlag=false
 On error resume next
  
   OracleIProcPageObj.Sync
'   mercer and OWG  =Print "This is the Order form Test Case."
   If gstrTestCaseExec_id = "GSI.P2P.IP.SA.007" Or  gstrTestCaseExec_id = "GSI.P2P.IP.SA.008" Then
   Print "This is Smart form Test Case"
   Else
   	Call fn_NonCatlogAddToCart()
   End If
   
       If OracleIProcPageObj.Link(NonCatalaogLink_xpath).Exists(5) Then
        	fnReportEvent "Pass", "Oracle Iproc page navigation status","Successfully navigated to Oracle Iproc Home page ",false
        If OracleIProcPageObj.WebButton(checkoutBtn_xpath).Exist(10) then
        	Call fn_deleteCart
          	blnResultflag=true
        End if
    Else
        fnReportEvent "Fail","Oracle Iproc page navigation status"," Failed to navigate Oracle Iproc Home page",true
        fn_enterAndCheckoutNCRequest=false
        Exit Function
    End If  
    fn_Highlight(OracleIProcPageObj)
If OracleIProcPageObj.WebEdit(itemDesc_xpath).Exist(10) then 
	fnSet_FieldName OracleIProcPageObj.WebEdit(itemDesc_xpath),  gb_TestDataDic.item("Item_Description"),"ItemDescription"
  	fnSet_FieldName OracleIProcPageObj.WebEdit(category_xpath),gb_TestDataDic.item ("Category"),"Category"
 	fnSet_FieldName OracleIProcPageObj.WebEdit(quantity_xpath),gb_TestDataDic.item ("Qty"),"Qty"
    	fnSet_FieldName OracleIProcPageObj.WebEdit(unitOfMeasure_xpath), gb_TestDataDic.item ("Unit Of Measure"),"Unit of measure"
   	fnSet_FieldName OracleIProcPageObj.WebEdit(unitPrice_xpath),gb_TestDataDic.item ("Unit Price"),"Unit Price"
	If len(gb_TestDataDic.item("Foreign Currency")) > 1 Then
   		blnResultflag = fn_validateForeignCurrency
   		If blnResultflag = False Then
   			Exit Function
   		End If
   		
   		fn_SelectWeblist OracleIProcPageObj.WebList(currencydropdown),gb_TestDataDic.item("Foreign Currency"),"Foreign Currency"
   	End If
	
	fnSet_FieldName OracleIProcPageObj.WebEdit(supplier_xpath), gb_TestDataDic.item ("Supplier Name"),"Supplier Name"   
    	fnSet_FieldName OracleIProcPageObj.WebEdit(supplierSite_xpath),gb_TestDataDic.item ("Site"),"Supplier Site details"
   	'This will add the requisition to the cart
       fn_AddtoCart
	
		If OracleIProcPageObj.WebElement("innertext:=Error.*","class:=x5y").Exist(10) Then
			fnReportEvent "Fail","Requistion creation","Fail to create requistion due to invalid test data  ",true
			fn_enterAndCheckoutNCRequest=false
			Exit Function
		End If	
			       
		If  fn_ClickCheckOutButton Then
			fn_enterAndCheckoutNCRequest=true
		End If	    
     Else 
      		fnReportEvent "Fail","Non Catalog page","Fail to load/navigate the Non Catalog page",True
	End If     
	blnResultFlag = fn_selectBillableOrNonbillable
	If Err.number <> 0 Then             
              fnReportEvent "Fail","Enter values on Non Catalogue request","Failed to Enter values on Non Catalogue request page" & Err.description,false
              fn_enterAndCheckoutNCRequest = false
             Exit function
      End If    	

	  	fn_enterAndCheckoutNCRequest = blnResultFlag
End Function


'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_vrfyapproversSubmitReq
'description:             fn_vrfyapproversSubmitReq :  submit the requistion 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_vrfyapproversSubmitReq()
blnResultFlag=false
 On error resume next
  If OracleIProcPageObj.WebElement(approvalPagexpath).Exist(15)  Then
       	fnReportEvent "Pass", "Oracle Iproc approval page navigation status","Successfully navigated to Oracle Iproc approval page ",false
	Call fn_verifyApprovers	
	blnResultFlag =fn_SubmitRequisition
'	fn_vrfyapproversSubmitReq = blnResultFlag
Else
	fnReportEvent "Fail", "Oracle Iproc Approval page navigation status","Not able to navigate to Oracle Iproc Approval page ",True
End  If
	fn_vrfyapproversSubmitReq = blnResultFlag

End Function


'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_glbaccvalidSubmitReq
'description:             fn_glbaccvalidSubmitReq :  submit the requistion 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_glbaccvalidSubmitReq()
blnResultFlag=false
 On error resume next
 
  If OracleIProcPageObj.WebElement(approvalPagexpath).Exist(5)  Then
       	fnReportEvent "Pass", "Oracle Iproc approval page navigation status","Successfully navigated to Oracle Iproc approval page ",false
	If fn_verifyApprovers = True Then
'		blnResultFlag= fn_GlobalAccountValidation

			Set dict_GAInfo = CreateObject("Scripting.Dictionary")
			Call fn_Click_fieldname(OracleIProcPageObj_1.WebElement(showDetails_xpath),"Show Details Button")
			Set dict_GAInfo = fn_getChargeAccountInformation
		
			glblaccountString = Cstr(dict_GAInfo.Item("Global Account"))
		    	
		    	threshold_value = CLng(gb_TestDataDic.item("FA Threshold Amount"))
		    	 req_threshold_Str=   gb_TestDataDic.item("Unit Price")
		    	 req_threshold= CLng(req_threshold_Str)
		    	 
		If Len(threshold_value) > 1 Then    	
	
'		 compresult = strcomp (req_threshold, threshold_value)
		
		    	If req_threshold > threshold_value Then ':Above Threshold Validation
'			If compresult=1 Then

			    	If glblaccountString = "15104" Then
			    		fnReportEvent "Pass","Global Account Number","Global Account Number value is :="&dict_GAInfo("Global Account"),False
			    		blnResultFlag = True
			    	Else
			      	 	fnReportEvent "Fail","Global Account Number","Failed to validate the Global Account Number value is :="&dict_GAInfo("Global Account"),True   
			    		Exit Function
			    	End  If
			ElseIf req_threshold < threshold_value Then ':Below Threshold Validation
'			ElseIf compresult=0 Then
				If glblaccountString = "53861" Then
			    		fnReportEvent "Pass","Global Account Number","Global Account Number value is :="&dict_GAInfo("Global Account"),False
			    		blnResultFlag = True
			    	Else
			      	 	fnReportEvent "Fail","Global Account Number","Failed to validate the Global Account number. The value is :="&dict_GAInfo("Global Account"),True   
			    		Exit Function
			    	End  If
		    	End If
			blnResultFlag= fn_SubmitRequisition
			blnResultFlag= fn_verifyConfirmation
		Else
			blnResultFlag= fn_SubmitRequisition
			blnResultFlag= fn_verifyConfirmation
		End If
	Else
	fnReportEvent "Fail", "Oracle Iproc Approval page navigation status","Not able to navigate to Oracle Iproc Approval page ",True
	Exit Function
End  If
	Else
	       	fnReportEvent "Fail", "Oracle Iproc approval page navigation status","Could not  navigate to Oracle Iproc approval page ",false
End  If
	fn_glbaccvalidSubmitReq = blnResultFlag
	
If Err.number <> 0 Then             
             fnReportEvent "Fail","Global Account Validation","Failed to validate the global account." & Err.description,false
             fn_putApproverNametoTestData = false
        	   Exit function
End if 
	
End Function


Function fn_verifyLERCGlobalSubAccount()
On Error Resume Next
blnResultCounter= 1
Call fn_Click_fieldname(OracleIProcPageObj_1.WebElement(showDetails_xpath),"Next Button")     

Set ChargeAccountInformation = fn_getChargeAccountInformation
   	
   If ChargeAccountInformation("Legal Entity") = gb_TestDataDic.item("Legal Entity") Then
  	 fnReportEvent "Pass","Legal Entity","Legal entity is showing the expected value",False	
   Else
  	 fnReportEvent "Fail","Legal Entity","Legal entity is not showing the expected value",False	
  	 blnResultCounter = blnResultCounter + 1
   End If
  
   If ChargeAccountInformation("Global Account") = gb_TestDataDic.item("Global Account") Then
   	fnReportEvent "Pass","Global Account","Successfuly validated the Global Account and the value as per the expected data is:-"&gb_TestDataDic.item("Global Account"),False	
	
 Else
   	fnReportEvent "Fail","Global Account","Failed to validateb the Global Account is not showing the expected value",False	
  	blnResultCounter = blnResultCounter + 1
  End If	
   
   If ChargeAccountInformation("Sub Account") = gb_TestDataDic.item("SubAccount") Then
   	fnReportEvent "Pass","Sub Account","Sub Account is showing the expected value",False	
   	
   Else
   	fnReportEvent "Fail","Sub Account","Sub Account is not showing the expected value",False	
   	blnResultCounter = blnResultCounter + 1
   End If	
   
If ChargeAccountInformation("RC")  = gb_TestDataDic.item("RC") Then
	fnReportEvent "Pass","Sub Account","Sub Account is showing the expected value",False	
	
Else
	fnReportEvent "Fail","Sub Account","Sub Account is not showing the expected value",False	
	blnResultCounter = blnResultCounter + 1
End If	
   
   If blnResultCounter>=2 Then
   	fn_verifyLERCGlobalSubAccount = false
   Else
    	fn_verifyLERCGlobalSubAccount = true
   End If
   
   If Err.number <> 0 Then             
              fnReportEvent "Fail","Verify LE RC Global and Sub Account","Failed to Verify LE RC Global and Sub Account" & Err.description,false
             fn_verifyLERCGlobalSubAccount = false
             Exit function
End If 

End Function

Function fn_LERCGlobalSubAccountValidSubmitReq()
blnResultFlag=false
 On error resume next
 
  If OracleIProcPageObj.WebElement(approvalPagexpath).Exist(5)  Then
       	fnReportEvent "Pass", "Oracle Iproc approval page navigation status","Successfully navigated to Oracle Iproc approval page ",false
'Modify and add the error conditioning
If fn_verifyApprovers=True Then
	blnResultFlag = fn_verifyLERCGlobalSubAccount
	blnResultFlag = fn_SubmitRequisition
End If

Else
	fnReportEvent "Fail", "Oracle Iproc Approval page navigation status","Not able to navigate to Oracle Iproc Approval page ",True
	Exit Function
End  If
	fn_LERCGlobalSubAccountValidSubmitReq = blnResultFlag
	
	 If Err.number <> 0 Then             
              fnReportEvent "Fail","Verify LE RC Global Sub Account & Submit Requisition","Failed to Verify LE RC Global and Sub Account & Submit Requisition" & Err.description,false
             fn_LERCGlobalSubAccountValidSubmitReq = false
             Exit function
          End if 
End Function

Function fn_enterMultiLineValuesonNonCatlogRequestPage()
On Error Resume Next
blnresultflag =false
'Call  fn_AddtoCart
Call fn_NonCatlogAddToCart
itemDescArray = Split (gb_TestDataDic.item("Item_Description"),"|")
If OracleIProcPageObj.WebEdit(itemDesc_xpath).Exist(10) then 

For index = 0 To Ubound(itemDescArray)
	fnSet_FieldName OracleIProcPageObj.WebEdit(itemDesc_xpath),Split (gb_TestDataDic.item("Item_Description"),"|")(index),"ItemDescription"		
  	fnSet_FieldName OracleIProcPageObj.WebEdit(category_xpath),Split (gb_TestDataDic.item("Category"),"|")(index),"Category"
	
	if index = 0 then 
		fnSet_FieldName OracleIProcPageObj.WebEdit(quantity_xpath),gb_TestDataDic.item("Qty"),"Qty"  
	ElseIf index=1 then 
		fnSet_FieldName OracleIProcPageObj.WebEdit(quantity_xpath),gb_TestDataDic.item("Qty1"),"Qty"     
	End  If 	
	 
	 fnSet_FieldName OracleIProcPageObj.WebEdit(unitOfMeasure_xpath),Split (gb_TestDataDic.item("Unit Of Measure"),"|")(index),"Unit of measure"
	
	if index = 0 then 
		fnSet_FieldName OracleIProcPageObj.WebEdit(unitPrice_xpath),gb_TestDataDic.item("Unit Price"),"Unit Price"  
	ElseIf index=1 then 
		fnSet_FieldName OracleIProcPageObj.WebEdit(unitPrice_xpath),gb_TestDataDic.item("UnitPrice1"),"Unit Price"     
	End  If 

	fnSet_FieldName OracleIProcPageObj.WebEdit(supplier_xpath),Split ( gb_TestDataDic.item("Supplier Name"),"|")(index),"Supplier Name" 	    
	fnSet_FieldName OracleIProcPageObj.WebEdit(supplierSite_xpath),Split (gb_TestDataDic.item("Site"),"|")(index),"Supplier Site details"
	Call fn_AddtoCart()
		
		If OracleIProcPageObj.WebElement("innertext:=Error.*","class:=x5y").Exist(10) Then
			fnReportEvent "Fail","Requistion creation","Fail to create requistion due to invalid test data  ",true
			fn_enterMultiLineValuesonNonCatlogRequestPage=blnresultflag
			Exit Function
		End If
Next
			
			       
		If  fn_ClickCheckOutButton Then
			fn_enterMultiLineValuesonNonCatlogRequestPage=true
		End If	    
Else
		fnReportEvent "Fail","Non Catalog page","Fail to load/navigate the Non Catalog page",True
End If

If Err.number <> 0 Then             
              fnReportEvent "Fail","Enter values on Non Catalogue request","Failed to Enter values on Non Catalogue request page" & Err.description,false
             fn_enterMultiLineValuesonNonCatlogRequestPage = false
             
End If 

End  Function


Function fn_getChargeAccountInformation()
On Error Resume Next

Set ChargeAccountInfo = CreateObject("Scripting.Dictionary")

If OracleIProcPageObj_1.WebElement("title:=Charge Account","index:=0").Exist(5) Then
'    	fn_Highlight OracleIProcPageObj_1.WebElement("title:=Charge Account","index:=0")
    	retChargeAccountNum = OracleIProcPageObj_1.WebElement("title:=Charge Account","index:=0").GetROProperty("innertext")
    	fnReportEvent "Pass","Charge Account Number","Charge Account Number is present on page",False
Else
        fnReportEvent "Fail","Charge Account Number","Charge Account Number is not present on page",True    
End  If

splitChargeAccount = split(retChargeAccountNum,".")

ChargeAccountInfo.Add "Legal Entity",splitChargeAccount(0)
ChargeAccountInfo.Add "Global Account",splitChargeAccount(1)
ChargeAccountInfo.Add "Sub Account",splitChargeAccount(2)
ChargeAccountInfo.Add "RC",splitChargeAccount(3)

If Err.number <> 0 Then             
	fnReportEvent "Fail","Charge Account Information","Failed to Fetch theCharge account information" & Err.description,false
	fn_getChargeAccountInformation = false
  End  If

Set fn_getChargeAccountInformation = ChargeAccountInfo

End Function



Function fn_SetCheckoutTable(irow,pColname,value)
On error Resume Next

If len(value)>1 Then
	xpathvalue ="//input[@title='" & pColname &"' ]"	
	Set objWebEdit =  Description.Create
	objWebEdit("micclass").value = "WebEdit"
	objWebEdit("xpath").value  = xpathvalue
	If  OracleIProcPageObj_1.Webtable(checkoutTable).Exist(10) = true Then
		set objprojectsource  = OracleIProcPageObj_1.Webtable(checkoutTable).ChildObjects(objWebEdit)
		objprojectsource(irow-1).set value
		fn_WSSendKeys ("TAB")
		wait(1)
	End If	
End If

If Err.number<> 0 Then
	fnReportEvent "FAIL",pColname& "FieldName","Function =fn_SetCheckoutTable ,Fail to enter  " &  pColname & "and  Value is:= "& Value , true	
End If

End Function

Function fn_SubmitReqUsingtwoRequester()

	On error Resume Next 	
	if  fn_selectBillableOrNonbillable = true then 
		fn_click OracleIProcPageObj_1.webbutton(editlines_btn)
	End  If 
End Function

Function fn_validateForeignCurrency()

On Error Resume Next
blnResultFlag=false
		currOnPage = fn_GetROPropertyValueByPropName(OracleIProcPageObj.WebList(exchangeCurrAmount),"value")
'		currOnPage = OracleIProcPageObj.WebList(exchangeCurrAmount).GetROProperty("value")
		currencyVal = gb_TestDataDic.item("Foreign Currency")
		If currencyVal <> currOnPage Then
			fnReportEvent "Pass","Foreign Currency Requisition","This is foreign currency requisition.",False
			blnResultFlag=true
		Else
			fnReportEvent "Fail","Foreign Currency Requisition","This is not foreign currency requisition. Foreign currency should be different than "&currOnPage&" in the test data",False			
		End If
		
	fn_validateForeignCurrency = blnResultFlag
	
End Function


Function fn_VerifyExchangeCurrAmount()
On Error Resume Next

		exchangeCurrencyVal = OracleIProcPageObj.WebElement(exchangeCurrAmount).GetROProperty("innertext")
	If len(exchangeCurrencyVal) > 1 Then
		fnReportEvent "Pass","Exchange Currency","The Exchanged currency value is:="& exchangeCurrencyVal,False
    		fn_VerifyExchangeCurrAmount = true
    	Else           
		fnReportEvent "Fail","Exchange Currency Value","The currency has not be exchanged with the Foreign currency value." & Err.description,false
		fn_VerifyExchangeCurrAmount = false	
	End If
   
End Function

Function fn_ForeignCurrValidationSubmitReq()
 On error resume next
 blnResultFlag=false
  If OracleIProcPageObj.WebElement(approvalPagexpath).Exist(5)  Then
       	fnReportEvent "Pass", "Oracle Iproc approval page ","Successfully navigated to Oracle Iproc approval page ",false
	
	If fn_verifyApprovers=False Then
		Exit Function	
	Else	
		blnResultFlag=fn_VerifyExchangeCurrAmount
		blnResultFlag=fn_SubmitRequisition
		blnResultFlag=fn_verifyConfirmation
	
	End  If
	
Else
	fnReportEvent "Fail", "Oracle Iproc Approval page navigation status","Not able to navigate to Oracle Iproc Approval page ",True
	Exit Function
End  If

	fn_Click OracleIProcPageObj_3.WebElement(requisitionTabXpath)
	
	If OracleIProcPageObj_4.WebButton(GoButton).Exist Then
		fnReportEvent "Pass","Oracle iProcurement:Requisitions Page Navigation","Successfully navigated to Oracle iProcurement:Requisitions Page",False 
	Else
		fnReportEvent "Fail","Oracle iProcurement:Requisitions Page Navigation","Not able to navigate to Oracle iProcurement:Requisitions Page",True
	End If
	
	strStatus = OracleIProcPageObj_4.WebElement(foreignReqStatus).GetROProperty("innertext")
	
	If strStatus = "In Process" Then
		fnReportEvent "Pass","Foreign Requisition Status","Successfully validated the status of newly created foreign requisition and it is :In Process",False 
	Else
		fnReportEvent "Fail","Foreign Requisition Status","The status of the newly created requisition is not as expected. i.e In Process",True
	End If

	If Err.number <> 0 Then             
		fnReportEvent "Fail","Foreign Currency Validation Submit Requisition","Failed to validate the foreign currency and SUbmit Requisition." & Err.description,false
		fn_ForeignCurrValidationSubmitReq = false		
  	End  If
	
End Function

Function fn_VerifyApproverList(approvalList)
On error resume next
 blnResultFlag=false
 
If (approvalList(0) =  gb_TestDataDic.item("Approver_1") AND approvalList(1) =  gb_TestDataDic.item("Approver_2") AND approvalList(2) =  gb_TestDataDic.item("Approver_3"))Then
	blnResultFlag = true
	fnReportEvent "Pass", "Approval Level hierarchy validation","Approval hierarchy validation is successfull.",True
Else
 	blnResultFlag = False 
 	fnReportEvent "Fail", "Approval Level hierarchy validation","Approval hierarchy validation is Failed.",True
End If

fn_VerifyApproverList = blnResultFlag

End Function

Function fn_verifyIfTheProjectCodeIsMandatory()
	
	On error resume next
	blnResultFlag = false	
		fn_Click_fieldname 	OracleIProcPageObj_1.WebButton(nextBtn1_xpath),"Next Button"
	    	approvalErrorText = OracleIProcPageObj_1.WebElement("html id:=FwkErrorBeanId").GetROProperty("innertext")
		
		If len(approvalErrorText) > 1 Then
		
		    	If( len(approvalErrorText) > 1 AND IsNull((gb_TestDataDic.item("Project_source")))) Then
		    		If InStr(approvalErrorText,"Approval List could not be generated") Then
		    			fnReportEvent "Pass","OWG/NERA Project Code Mandatory Check","As Project Code is mandatory for OWG Requisition Creation, We are getting expected Error :-"&approvalErrorText,True 
		    			blnResultFlag = True
		    		End  If		    		
		    	Else 
		    			blnResultFlag = False		    		
		    	End  If
		
		Else		
	    	ApprovalNoErrortext = OracleIProcPageObj_1.WebElement("html id:=ApprListText").GetROProperty("innertext")
			If InStr(ApprovalNoErrortext,"Your requisition will be sent to the following list of approvers.") Then
	    			fnReportEvent "Pass","MERCER/MARSH Project Code Mandatory Check","As Project Code is not mandatory for Mercer/Marsh Requisition Creation, We are getting the expected msg as :-" & ApprovalNoErrortext,True 
	    			blnResultFlag = True
	    		Else
	    			blnResultFlag = False
	    		End If
	    	End  If
	    	
	    	
	    	 If Err.number <> 0 Then
	             fnReportEvent "Fail","Verify if the Project Code is Mandatory","Failed to validate if the project code is mandatory." & Err.description,false
	             fn_verifyIfTheProjectCodeIsMandatory = false
	             Exit function
      		End If
	
	fn_verifyIfTheProjectCodeIsMandatory = blnResultFlag
		
	
End Function

Function fn_CreateChangeOrderfromApprvedPO()
	
	On error resume next
	blnResultFlag = false
	
	fn_Click OracleIProcPageObj.WebElement(requisitionTabXpath)
	If OracleIProcPageObj_4.WebButton(GoButton).Exist Then
		fnReportEvent "Pass","Oracle iProcurement:Requisitions Page Navigation","Successfully navigated to Oracle iProcurement:Requisitions Page",False 
	Else
		fnReportEvent "Fail","Oracle iProcurement:Requisitions Page Navigation","Not able to navigate to Oracle iProcurement:Requisitions Page",True
	End If
	
'	PO_Num = "11200259927"
'	searchTextInColumn = "7"
	
	PO_Num = AutoCreatedPONumber
	
	If PO_Num=0 or PO_Num="" or Isnull(PO_Num) Then
		fnReportEvent "Fail","Auto Created PO Number","Failed to fetch the auto created PO Number.",True
		Exit Function
	End If
	
	searchTextInColumn = "7"   ' WRITE THE FUNCTION TO FETECH COLUMN NUMBER BASED ON COLUMN NAME
	
	reqTableRows = OracleIProcPageObj_4.WebTable("name:=Requisition").GetROProperty("Rows")
	reqPOFound = False
	for Crow = 1 to reqTableRows
		
		If PO_Num = Trim(OracleIProcPageObj_4.WebTable("name:=Requisition").GetCellData(Crow,searchTextInColumn)) Then
			FoundRow = Crow
			reqPOFound = True
			Exit For 
		End if 
	Next
		
		If reqPOFound Then
			Set reqPOLinkRadioButton = OracleIProcPageObj_4.WebTable("name:=Requisition").ChildItem(FoundRow,1,"WebRadioGroup",0)
			reqPOLinkRadioButton.Select FoundRow - 2
			fnReportEvent "Pass","The PO Number in Requisitions table","Successfully found and clicked the Selected Requition against the PO Number",False
		Else 
			fnReportEvent "Fail","The PO Number in Requisitions table","Couldn't find the PO number on Requisitions Page",False
		End If
	 
	If OracleIProcPageObj_4.WebButton(changeButton).Exist(5) = true	Then
		fn_Click OracleIProcPageObj_4.WebButton(changeButton)
		If OracleIProcPageObj_ChangeOrder.Exist(5)=true Then
			If OracleIProcPageObj_ChangeOrder.WebEdit(reasonTextXpath).Exist(5) Then
				fn_Highlight OracleIProcPageObj_ChangeOrder.WebEdit(reasonTextXpath)			
				fnReportEvent "Pass","Oracle iProcurement:Requisitions Change Button","Successfully Clicked on Change button and It navigated to Change order page.",False 
			Else
				fnReportEvent "Fail","Oracle iProcurement:Requisitions Change Button","Clicked on Change button but It could not navigate to Change order page.",True 
			End If
		Else
			changeOrderErrortext = OracleIProcPageObj_4.WebElement("html id:=FwkErrorBeanId").GetROProperty("innertext")
			If InStr (changeOrderErrortext, "Change request cannot be initiated") Then
				fnReportEvent "Fail","Change Request for the requisition","Change request cannot be initiated for this requisition since none of the associated purchase orders are eligible for change.",True 
			Else 
				fnReportEvent "Fail","Change Request for the requisition","Change request cannot be initiated.",True 
			End If		
		End If	
	Else
		fnReportEvent "Fail","Oracle iProcurement:Requisitions Change Button","Couldn't find and click Change button on iProcurement:Requisitions Page",True 
	End  If
	
	PO_amt = gb_TestDataDic.item("PO_Change_Amount")
'	PO_amt = "1550000"
	fnSet_FieldName OracleIProcPageObj_ChangeOrder.WebEdit(updatedPOPriceTextBoxXpath),PO_amt,"Updated Unit Price Amount"
	fnSet_FieldName OracleIProcPageObj_ChangeOrder.WebEdit(reasonTextXpath),"This is the reason","Reason for Change OrderS"
	OracleIProcPageObj_ChangeOrder.WebButton(changeOrderNextButton).Click
	
	strApprText = OracleIProcPageObj_ChangeOrder.WebElement(changeOrderApprText).GetROProperty("innertext")
	
	fn_Click_fieldname OracleIProcPageObj_ChangeOrder.WebButton(selectApprNextButton),"Next Button"
	
	If InStr(strApprText,"Your changes will be sent to the following list of approvers.") >0 Then
		fnReportEvent "Pass","Change Order: Select Approvers","After Change Request of PO, Successfully approval flow is generated",False
	Else
		fnReportEvent "Fail","Change Order: Select Approvers","After Change Request of PO, approval flow is not generated",True 	
		Exit Function
	End If

	fn_Click_fieldname OracleIProcPageObj_ChangeOrder.WebButton(reqSubmitBtn_xpath), "Submit button"
'		If fn_Click(OracleIProcPageObj_ChangeOrder.WebButton(reqSubmitBtn_xpath)) Then        
'	       		 fnReportEvent "Pass","Submit Button","Succesfully Click on the Submit Button",False
'	    	Else
'	       		 fnReportEvent "Fail","Submit Button","Submit Button doesnt Exist",True    
'	    	End  If
	
	POChangeRequestConfirm = OracleIProcPageObj_ChangeOrder.WebTable("html id:=FwkErrorBeanId").GetROProperty("innertext")
'	msgbox textwegot
	
		If InStr(POChangeRequestConfirm,"have been submitted for processing. View status of the change request(s) from the Requisition Status page") >0 Then
			fnReportEvent "Pass","Change Order Confirmation","After Change Request of PO, It is submitted successfully for processing.",False
			blnResultFlag = True
		Else
			fnReportEvent "Fail","Change Order Confirmation","Change Request PO is not submitted successfully.",True 	
		End If
		
		 If Err.number <> 0 Then             
	             fnReportEvent "Fail","Change Order from Approved PO","Failed to Change the Order from Approved PO" & Err.description,false
	             fn_CreateChangeOrderfromApprvedPO = false
	             Exit function
      		End If
		
		fn_CreateChangeOrderfromApprvedPO = blnResultFlag
	
End Function

Function fn_SmartOrderPageNavigation()

On error resume next
blnResultFlag = false
	fn_Click_fieldname OracleIProcPageObj.Link(officeSuppliesXpath),"officeSupplies link"
	blnResultFlag=	fn_Click_fieldname( OracleIProcPageObj.Link(needToBuyGoodsXpath),"I need to Buy goods Link")
'	If OracleIProcPageObj.Link(officeSuppliesXpath).Exist Then
'        	fnReportEvent "Pass", "Office Supplies Link","Office Supplies Link exist",false
'        	 fn_click OracleIProcPageObj.Link(officeSuppliesXpath)
'        	If OracleIProcPageObj.Link(needToBuyGoodsXpath).Exist Then
''	        	OracleIProcPageObj.Link(needToBuyGoodsXpath).Click
'	        	fn_Click OracleIProcPageObj.Link(needToBuyGoodsXpath)
'	        	fnReportEvent "Pass", "I need to Buy goods Link","I need to Buy goods Link exist",false
'	        	blnResultFlag = True
'        	Else
'        		fnReportEvent "Fail", "I need to Buy goods Link","Failed to find := I need to Buy goods Link on the page.",True
'        		blnResultFlag = false
'        	End If
'        Else
'                fnReportEvent "Fail", "Office Supplies Link","Office Supplies Link does not exist",false
'                blnResultFlag = false
'	End If
'	
	 If Err.number <> 0 Then             
	             fnReportEvent "Fail","Smart Order Page Navigation","Failed to validate Smart Order Page navigation." & Err.description,false
	             fn_SmartOrderPageNavigation = false
	             Exit function
      		End If
	
	fn_SmartOrderPageNavigation = blnResultFlag
	
End Function

Function fn_RequisitionApproval()

On error resume next
blnResultFlag = false

reqNumber = fn_getExecutionResultData(gstrTestCaseExec_id,"Requisition_No")

If len(reqNumber)= 0 or reqNumber="" or Isnull(reqNumber) Then
	fnReportEvent "Fail","Requisition Number","Failed to fetch the Requisition Number.",True
	Exit Function
End If
'Array of 
Set notificationSearchlink = wflowconfigobj.Link(WfNotificationSearch)
fn_Click_fieldname notificationSearchlink,"Notification Search"
'If notificationSearchlink.Exist(5) Then
'	notificationSearchlink.Click
'	
'Else
'	fnReportEvent "Fail","Notification Search","Notification Search link not found",True	
'End If
'
cnt=1
For K=0 To Count_appr-2
	If notificationsobj.Exist(5) Then
	 
		strToUser = fn_getExecutionResultData(gstrTestCaseExec_id,"Approver"&cnt&"")
		cnt = cnt + 1
		fnSet_FieldName notificationsobj.WebEdit(frmNotification),strToUser,"User Name"
		fnSet_FieldName notificationsobj.webEdit(Subject),"%"&reqNumber&"%","Requistion Number"
		fn_Click notificationsobj.webButton(admin_goButton)
	'	notificationsobj.WebEdit(frmNotification).Set strToUser
	'	notificationsobj.webEdit(Subject).Set "%"&reqNumber&"%"
	'	notificationsobj.webButton(admin_goButton).Click
	'	
	Else
		fnReportEvent "Fail","Notifications Page","Notifications Page not found",True
	End If
		
		intLoopCnt = 0
							Do
									fn_Click notificationsobj.WebButton(goButton) 
									notificationsobj.WebButton(goButton).Click
									intLoopCnt = intLoopCnt +1
									If intLoopCnt = 60 Then
										Exit Do
									End If
							Loop While NOT notificationsobj.Link(searchedPurchaseReqLink).exist(6)
	
	
		If notificationsobj.Link(searchedPurchaseReqLink).exist(7) Then
		'	notificationsobj.Link(searchedPurchaseReqLink).click
			fn_Click notificationsobj.Link(searchedPurchaseReqLink)
		Else 
			fnReportEvent "Fail","Seacrhed Purchae Requisition Link","Seacrhed Purchae Requisition Link not found",False
			fn_RequisitionApproval = false
			Exit function 
		End If
	
		If notificDetailsobj.WebButton(approveButton).Exist(30) Then
			fnReportEvent "Pass","Notification Details Page","Notification Details Page  found",False
			fn_Click notificDetailsobj.WebButton(approveButton)
	'		notificDetailsobj.WebButton(approveButton).Click
			fnReportEvent "Pass","Requisition Approval","Requisition is approved",False
			blnResultFlag = True
	
		Else
			fnReportEvent "Fail","Approve Button","Approve Button not  found",True
			blnResultFlag = False
		End If
	strToUser = NULL
Next

  fn_Click_fieldname notificationsobj.Link(homePageLinkAdmin),"Home Page Link"
'		If notificationsobj.Link(homePageLinkAdmin).Exist(2) Then  
'			notificationsobj.Link(homePageLinkAdmin).Click
'			fnReportEvent "Pass","Home Page Link","Succesfully Clicked on Home Page Link",False
'		Else
'			fnReportEvent "Fail","Home Page Link","Couldn't Click on Home Page Link",True
'		End  If
'
 If Err.number <> 0 Then             
             fnReportEvent "Fail","Requisition Approval","Failed to Approve the Requisition" & Err.description,false
             fn_RequisitionApproval = false
             Exit function
      End If

fn_RequisitionApproval = blnResultFlag

End Function



Function fn_ReceiveRequisition()

blnResultFlag = false
On error resume next

	If OracleIProcPageObj.WebElement(receivingtab).Exist Then
		OracleIProcPageObj.WebElement(receivingtab).Click
		fnReportEvent "Pass","Receiving tab","Receiving tab found",True		
			If OracleIProcPageObj_Receiving.WebElement(receiveItemsLink).Exist Then
				fnReportEvent "Pass","Receiving Page","Receiving Page found",True
				OracleIProcPageObj_Receiving.WebElement(receiveItemsLink).Click
							getReqNumber = fn_getExecutionResultData("GSI.P2P.IP.SA.013","Requisition_No")
							If getReqNumber=0 or getReqNumber="" or Isnull(getReqNumber) Then
							fnReportEvent "Fail","Requisition Number","Failed to fetch the Requisition Number.",True
							End If
							
							If OracleIProcPageObj_Receiving.WebElement(reqNumbertextbox).Exist Then
								OracleIProcPageObj_Receiving.WebEdit(reqNumbertextbox).set  getReqNumber
								 fn_SelectWeblist OracleIProcPageObj_Receiving.WebList(receievReqItemsDue),"Next 30 Days","Receive Requisition Items Due"
								fnReportEvent "Pass","Enter Req Number Text box","Enter Req Number Text box found",False
							
									If OracleIProcPageObj_Receiving.WebButton(goButtonSearchReq).Exist Then
	 								 	OracleIProcPageObj_Receiving.WebElement(goButtonSearchReq).Click
	 									 fnReportEvent "Pass","Go Button","Clicked on Go button on receiving Page.",False
									
												If OracleIProcPageObj_Receiving.WebElement(receiptQuantity).Exist Then
'	 											OracleIProcPageObj_Receiving.WebElement(receiptQuantity).Click
	 											'Write the code here to put the Recieve Receipt quantity
	 											 fnReportEvent "Pass","Receipt Quantity Edit box","Receipt Quantity Edit box is present on screen.",False
												
'														If fn_Click(OracleIProcPageObj_Receiving.WebElement(selectReceivingReq)) Then     
														If OracleIProcPageObj_Receiving.WebElement("name:=N3:selected:0").Exist Then
															OracleIProcPageObj_Receiving.WebCheckBox("name:=N3:selected:0").Click
	       									 				fnReportEvent "Pass","Select Requisition Check box","Select Requisition Check box exist and clicked on it.",False
	    													
	    														If fn_Click(OracleIProcPageObj_Receiving.WebButton(FullyReceiveNextButton)) Then        
	       									 					fnReportEvent "Pass","Next Button","Succesfully Click on the Next Button",False
	       									 					
	       									 						If fn_Click(OracleIProcPageObj_Receiving.WebButton(FullyReceiveNextNextButton)) Then        
	       									 						fnReportEvent "Pass","Next Button","Succesfully Click on the Submit Button",False
	    															Else
	       									 						fnReportEvent "Fail","Next Button Button","Submit Button doesnt Exist",True    
	    															End  If
	    															
	    															If fn_Click(OracleIProcPageObj_Receiving.WebButton(reqSubmitBtn_xpath)) Then        
	       									 						fnReportEvent "Pass","Submit Button","Succesfully Click on the Submit Button",False
	       									 						receiptconfirmation = OracleIProcPageObj_Receiving.WebTable("html id:=FwkErrorBeanId").GetROProperty("innertext")
	       									 						
	       									 							If InStr(receiptconfirmation,"has been created for you") >0 Then
																	fnReportEvent "Pass","Receipt Creation","Receipt creation is successfull",False
																	Else
																	fnReportEvent "Fail","Receipt Creation","Receipt creation is not successfull",True 	
																	End If
	       		
	       									 							blnResultFlag = True
	       									 							
	    															Else
	       									 						fnReportEvent "Fail","Submit Button","Submit Button doesnt Exist",True    
	    															End  If
	    														Else
	       									 					fnReportEvent "Fail","Next Button","Next Button doesnt Exist",True    
	    														End  If		
	    													    						
	    													Else
	       													 fnReportEvent "Fail","Next Button","Select Requisition Check box does not exist.",True    
	    													End  If
												
												Else
										 		fnReportEvent "Fail","Receipt Quantity Edit box","Receipt Quantity Edit box is not present on screen.",False
												End If
									
									Else
										 fnReportEvent "Fail","Go Button","Go Button is not present on screen.",False
									End If
						
							Else
	 							fnReportEvent "Fail","Enter Req Number Text box","Enter Req Number Text box not found",True
	
							End If
			Else
	 			fnReportEvent "Fail","Receiving Page","Receiving Page/Receive Items Link found",True
			End If
		
	Else
		fnReportEvent "Fail","Receiving tab","Receiving tab not found",True
	End If

	fn_ReceiveRequisition = blnResultFlag
	
	 If Err.number <> 0 Then             
             fnReportEvent "Fail","Requisition Receiving","Failed to Receive the Requisition" & Err.description,false
             fn_ReceiveRequisition = false
             Exit function
      End If
	
End Function


Function fn_ReturnRequisition()
	
blnResultFlag = false
On error resume next

	If OracleIProcPageObj.WebElement(receivingtab).Exist Then
		OracleIProcPageObj.WebElement(receivingtab).Click
		fnReportEvent "Pass","Receiving tab","Receiving tab found",True
		
			If OracleIProcPageObj_Receiving.WebElement(returntItemsLink).Exist Then
				fnReportEvent "Pass","Receiving Page","Receiving Page/Return Items page found",True
				OracleIProcPageObj_Receiving.WebElement(returntItemsLink).Click
							getReqNumber = fn_getExecutionResultData("GSI.P2P.IP.SA.013","Requisition_No")
							If getReqNumber=0 or getReqNumber="" or Isnull(getReqNumber) Then
							fnReportEvent "Fail","Requisition Number","Failed to fetch the Requisition Number.",True
							Exit Function
							End If
							
							If OracleIProcPageObj_ReturnItems.WebElement(reqNumbertextbox).Exist Then
								OracleIProcPageObj_ReturnItems.WebEdit(reqNumbertextbox).Set  getReqNumber
								'Write the code here to enter the requisition number
								fnReportEvent "Pass","Enter Req Number Text box","Enter Req Number Text box found",False
										
										If OracleIProcPageObj_ReturnItems.WebButton(goButtonSearchReq).Exist Then
	 								 	OracleIProcPageObj_ReturnItems.WebElement(goButtonSearchReq).Click
	 									 fnReportEvent "Pass","Go Button","Clicked on Go button on receiving Page.",False
									
												If OracleIProcPageObj_ReturnItems.WebElement(returnQuantity).Exist Then
	 											OracleIProcPageObj_ReturnItems.WebEdit(returnQuantity).Set "1"
	 											 fnReportEvent "Pass","Return Quantity Edit box","Return Quantity Edit box is present on screen.",False
												  			
	    														If fn_Click(OracleIProcPageObj_ReturnItems.WebButton(FullyReceiveNextButton)) Then        
	       									 					fnReportEvent "Pass","Next Button","Succesfully Click on the Next Button",False
																
																If OracleIProcPageObj_ReturnItems.WebEdit(returnReason).Exist Then
																	fnReportEvent "Pass","Return Reason Page","Return Reason Page and Return Reason text box exists.",False
																	OracleIProcPageObj_ReturnItems.WebEdit(returnReason).Set "Damaged Product"
																Else 
																	fnReportEvent "Fail","Return Reason Page","Return Reason Page and Return Reason text box does not exist.",False
																End If

																	
	       									 						If fn_Click(OracleIProcPageObj_ReturnItems.WebButton(FullyReceiveNextNextButton)) Then        
																		fnReportEvent "Pass","Next Button","Succesfully Click on the Submit Button",False
																	Else
																		fnReportEvent "Fail","Next Button Button","Submit Button doesnt Exist",True    
	    															End  If
	    															
	    															If fn_Click(OracleIProcPageObj_ReturnItems.WebButton(reqSubmitBtn_xpath)) Then        
																		fnReportEvent "Pass","Submit Button","Succesfully Click on the Submit Button",False
																		returnconfirmation = OracleIProcPageObj_ReturnItems.WebTable("html id:=FwkErrorBeanId").GetROProperty("innertext")
																		
	       									 							If InStr(returnconfirmation,"Your returns have been submitted") >0 Then
																			fnReportEvent "Pass","Return Confirmation","Return is successfull",False
																			blnResultFlag = True
																		Else
																			fnReportEvent "Fail","Return Confirmation","Return is not successfull",True 	
																		End If
	       									 							
	    															Else
																		fnReportEvent "Fail","Submit Button","Submit Button doesnt Exist",True    
	    															End  If
																	
	    														Else
	       									 					fnReportEvent "Fail","Next Button","Next Button doesnt Exist",True    
	    														End  If				    						
												Else
													fnReportEvent "Fail","Receipt Quantity Edit box","Receipt Quantity Edit box is not present on screen.",False
												End If
								Else
								fnReportEvent "Fail","Go Button","Go Button is not present on screen.",False
								End If											
										
							Else
	 							fnReportEvent "Fail","Enter Req Number Text box","Enter Req Number Text box not found",True
							End If
			Else
	 			fnReportEvent "Fail","Receiving Page","Receiving Page/Return Items Link found",True
			End If
	Else
		fnReportEvent "Fail","Receiving tab","Receiving tab not found",True
	End If
	
	fn_ReturnRequisition = blnResultFlag
	
	 If Err.number <> 0 Then             
             fnReportEvent "Fail","Requisition Return","Failed to Return the Requisition" & Err.description,false
             fn_ReturnRequisition = false
             Exit function
      End If
	
End Function

Function fn_POAutoCreationStatusCheck()

On error resume next
blnResultFlag = false
		If fn_Click(OracleIProcPageObj.WebElement(requisitionTabXpath)) Then
			fnReportEvent "Pass","Requisitions Tab","Successfully navigated to Requisitions Tab",False 
	Else
			fnReportEvent "Fail","Requisitions Tab","Could not navigate to Requisitions Tab",False 	
	End If
	
	If OracleIProcPageObj_4.WebButton(GoButton).Exist Then
		fnReportEvent "Pass","Oracle iProcurement:Requisitions Page Navigation","Successfully navigated to Oracle iProcurement:Requisitions Page",False 
	Else
		fnReportEvent "Fail","Oracle iProcurement:Requisitions Page Navigation","Not able to navigate to Oracle iProcurement:Requisitions Page",True
	End If
	
	Req_Num = fn_getExecutionResultData("GSI.P2P.IP.SA.013","Requisition_No")
	If len(Req_Num)=0 or Req_Num="" or Isnull(Req_Num) Then
		fnReportEvent "Fail","Requisition Number","Failed to fetch the Requisition Number.",True
		Exit Function
	End If
	searchTextInColumn = "2"
	
	reqTableRows = OracleIProcPageObj_4.WebTable("name:=Requisition").GetROProperty("Rows")
	reqFound = False
	for Crow = 1 to reqTableRows
		
		If Req_Num = Trim(OracleIProcPageObj_4.WebTable("name:=Requisition").GetCellData(Crow,searchTextInColumn)) Then
			FoundRow = Crow
			reqFound = True
			Exit For 
		End if 
	Next
		
		If reqFound Then
		
		intLoopApprCnt = 0
						Do
								wait (3)
								OracleIProcPageObj_4.WebButton(GoButton).Click
								reqStatus = Trim(OracleIProcPageObj_4.WebTable("name:=Requisition").GetCellData(Crow,"6"))
								intLoopApprCnt = intLoopApprCnt +1
								If intLoopApprCnt = 60 Then
									Exit Do
								End If
						Loop While NOT (reqStatus="Approved")
		
		If reqStatus="Approved" Then
					fnReportEvent "Pass","The Requisition Approval Status","The requisition No."&Req_Num&" is approved.",False
	
					intLoopCnt = 0
						Do
								wait(3)
								OracleIProcPageObj_4.WebButton(GoButton).Click
'								crow = OracleIProcPageObj_4.WebTable("name:=Requisition").GetRowWithCellText(1234567890")
								generatedPONumber = Trim(OracleIProcPageObj_4.WebTable("name:=Requisition").GetCellData(Crow,"7"))
								intLoopCnt = intLoopCnt +1
								If intLoopCnt = 60 Then
									Exit Do
								End If
						Loop While NOT Len(generatedPONumber)>0
						
'			need to add the validation  for checking hyper link is getting generated
		If Len(generatedPONumber)>0 Then
			fnReportEvent "Pass","The Generated PO Number","The Generated PO number is :-"&generatedPONumber,False
			blnResultFlag=True
			AutoCreatedPONumber = generatedPONumber
			wait(60)
		Else
			fnReportEvent "Fail","The Generated PO Number","The PO Number is not generated",False
		End if
		
		Else
		 			fnReportEvent "Fail","The Requisition Approval Status","The requisition No."&Req_Num&" is not in approved status.",False
					Exit Function
		End If
		
	Else 
			fnReportEvent "Fail","The Requisition Number in Requisitions table","Couldn't find the requisitions number on Requisitions table",False
	End If
	
	If Err.number <> 0 Then             
             fnReportEvent "Fail","Automatic PO Creation Check","Failed to validate the automated PO creation." & Err.description,false
             fn_POAutoCreationRequisitionStatusCheck = false
             Exit function
      End If
      
      fn_POAutoCreationStatusCheck = blnResultFlag
	
End Function

Function fn_requisitionStatusCheck()
	blnResultFlag = false
	On error resume next
	
		If fn_Click(OracleIProcPageObj.WebElement(requisitionTabXpath)) Then
			fnReportEvent "Pass","Requisitions Tab","Successfully navigated to Requisitions Tab",False 
	Else
			fnReportEvent "Fail","Requisitions Tab","Could not navigate to Requisitions Tab",False 	
	End If
	
	If OracleIProcPageObj_4.WebButton(GoButton).Exist Then
		fnReportEvent "Pass","Oracle iProcurement:Requisitions Page Navigation","Successfully navigated to Oracle iProcurement:Requisitions Page",False 
	Else
		fnReportEvent "Fail","Oracle iProcurement:Requisitions Page Navigation","Not able to navigate to Oracle iProcurement:Requisitions Page",True
	End If
	
	'Req_Num = "74800000142"
	Req_Num = fn_getExecutionResultData("GSI.P2P.IP.SA.013","Requisition_No")
	
	If Req_Num=0 or Req_Num="" or Isnull(Req_Num) Then
		fnReportEvent "Fail","Requisition Number","Failed to fetch the Requisition Number.",True
		Exit Function
	End If
	searchTextInColumn = "2"
	
	reqTableRows = OracleIProcPageObj_4.WebTable("name:=Requisition").GetROProperty("Rows")
	reqFound = False
	for Crow = 1 to reqTableRows
		
		If Req_Num = Trim(OracleIProcPageObj_4.WebTable("name:=Requisition").GetCellData(Crow,searchTextInColumn)) Then
		FoundRow = Crow
		reqFound = True
		Exit For 
	End if 
	Next
		
		If reqFound Then
		
		intLoopApprCnt = 0
						Do
								wait (3)
								OracleIProcPageObj_4.WebButton(GoButton).Click
								reqStatus = Trim(OracleIProcPageObj_4.WebTable("name:=Requisition").GetCellData(Crow,"6"))
								intLoopApprCnt = intLoopApprCnt +1
								If intLoopApprCnt = 60 Then
									Exit Do
								End If
						Loop While NOT (reqStatus="Approved")
		
		If reqStatus="Approved" Then
					fnReportEvent "Pass","The Requisition Approval Status","The requisition No."&Req_Num&" is approved.",False
					blnResultFlag = True
		Else
		 			fnReportEvent "Fail","The Requisition Approval Status","The requisition No."&Req_Num&" is not in approved status.",False
					Exit Function
		End If
	
	Else 
			fnReportEvent "Fail","The Requisition Number in Requisitions table","Couldn't find the requisitions number on Requisitions table",False
	End If
	
	If Err.number <> 0 Then             
             fnReportEvent "Fail","Requisition Status Check","Failed to validate the Requisition status" & Err.description,false
             fn_requisitionStatusCheck = false
             Exit function
      End If
      
      fn_requisitionStatusCheck = blnResultFlag
	
End Function
