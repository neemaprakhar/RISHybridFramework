Public OracleIProcPageObj,OracleIProcPageObj_1,OracleIProcPageObj_2,OracleIProcPageObj_3
Set OracleIProcPageObj = Browser("name:=Oracle iProcurement: Shop").Page("title:=Oracle iProcurement: Shop")
Set OracleIProcPageObj_1 =  Browser("name:=Oracle iProcurement: Checkout").Page("title:=Oracle iProcurement: Checkout")
Set OracleIProcPageObj_2= Browser("name:=Approval Group").Page("title:=Approval Group")
Set OracleIProcPageObj_3 = Browser("name:=Confirmation").Page("title:=Confirmation")
Set objApproverDict = CreateObject("Scripting.Dictionary")
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
'Const chargeAccountNum = "html id:=N5___PoReqDistributionsVO___0:Charg"
Const chargeAccountNum = "title:=Charge Account"
Const approvalPagexpath = "xpath:=//h1[text()='Checkout: Requisition Information']" 

Const checkoutTable = "xpath:=//SPAN[@id='ItemTableRN']/TABLE[2]"
Const editlines_btn = "//BUTTON[@id='EditLines']"
'===============================================================



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
'		Call fn_selectBillableOrNonbillable
		call fn_vrfyapproversSubmitReq 
'		Call fn_verifyApprovers
'		Call fn_SubmitRequisition
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
'    If OracleIProcPageObj.WebButton(addToCart_xpath).Exist(2) Then        
'        OracleIProcPageObj.WebButton(addToCart_xpath).Click
'      
	call fn_Click_fieldname(OracleIProcPageObj.WebButton(addToCart_xpath),"AddToCart Button")
'   If fn_Click(OracleIProcPageObj.WebButton(addToCart_xpath)) Then
'	fnReportEvent "Pass","AddToCart Button","Successfully click on the AddToCart Button",False
'    Else
'        fnReportEvent "Fail","AddToCart Button","AddToCart Button doesnt Exist",True    
'    End  If
End Function


'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_enterValuesonNonCatlogRequestPage 
'description          will fill the requistion details required to create requistion 	
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

'Function fn_enterValuesonNonCatlogRequestPage()
'On Error Resume Next
'blnresultflag =false
'   
'if OracleIProcPageObj.WebEdit(itemDesc_xpath).Exist(10) then 
'	fn_Set OracleIProcPageObj.WebEdit(itemDesc_xpath),  gb_TestDataDic.item("Item_Description"),"ItemDescription"
'	fnReportEvent "Pass","Item Description","Successfully enter Item Description and value is ::" & gb_TestDataDic.item("Item_Description") ,False
'  
'  	fn_Set OracleIProcPageObj.WebEdit(category_xpath),gb_TestDataDic.item ("Category"),"Category"
'	fnReportEvent "Pass","Category Description Field","Successfully  enter the Category and value is" & gb_TestDataDic.item ("Category") ,False
'  
' 	 fn_Set OracleIProcPageObj.WebEdit(quantity_xpath),gb_TestDataDic.item ("Qty"),"Qty"
'	fnReportEvent "Pass","Quantity TextBox Value","Successfully  enter the Quantity  and value entered is::" &gb_TestDataDic.item ("Qty"),False
'	
'    	fn_Set OracleIProcPageObj.WebEdit(unitOfMeasure_xpath), gb_TestDataDic.item ("Unit Of Measure"),"Unit of measure"
'	fnReportEvent "Pass","Unit Of measurement Text Box","Successfully  enter the Unit Of measurement  and value entered is::" & gb_TestDataDic.item ("Unit Of Measure") ,False
'   
'   	fn_Set OracleIProcPageObj.WebEdit(unitPrice_xpath),gb_TestDataDic.item ("Unit Price"),"Unit Price"
'	fnReportEvent "Pass","UnitPrice field","Successfully  enter the  UnitPrice  and value entered is::" & gb_TestDataDic.item ("Unit Price") ,False
'
'	fn_Set OracleIProcPageObj.WebEdit(supplier_xpath), gb_TestDataDic.item ("Supplier Name"),"Supplier Name" 
'	fnReportEvent "Pass","Supplier field details","Successfully  enter the  Supplier detail  and value entered is::" & gb_TestDataDic.item ("Supplier Name") ,False
'    
'    	fn_Set OracleIProcPageObj.WebEdit(supplierSite_xpath),gb_TestDataDic.item ("Site"),"Supplier Site details"
'	fnReportEvent "Pass","Supplier Site details","Successfully  enter the  Supplier  site detail  and value entered is::" & gb_TestDataDic.item ("Site") ,False
'   
'       Call fn_AddtoCart()
'	
'		If OracleIProcPageObj.WebElement("innertext:=Error.*","class:=x5y").Exist(10) Then
'			fnReportEvent "Fail","Requistion creation","Fail to create requistion due to invalid test data  ",true
'			fn_enterValuesonNonCatlogRequestPage=blnresultflag
'			Exit Function
'		End If	
'			       
'		If  fn_ClickCheckOutButton Then
'			fn_enterValuesonNonCatlogRequestPage=true
'		End If	    
'     Else 
'      		fnReportEvent "Fail","Non Catalog page","Fail to load/navigate the Non Catalog page",True
'	End If     
'	If Err.number <> 0 Then             
'              fnReportEvent "Fail","Enter values on Non Catalogue request","Failed to Enter values on Non Catalogue request page" & Err.description,false
'             fn_enterValuesonNonCatlogRequestPage = false
'             Exit function
'      End If    
'End Function
'
Function fn_selectBillableOrNonbillable()
On Error Resume Next
    If  gb_TestDataDic("Business_Line") = "Mercer" Then
'	OracleIProcPageObj_1.Webtable(checkoutTable).GetROProperty("Rows")
'    	fnset  OracleIProcPageObj_1.WebEdit(projectSrc_xpath),gb_TestDataDic.item("Project_source"), "Project Source"
'    	OracleIProcPageObj_1.WebEdit(projectSrc_xpath).SendKeys"{TAB}"
'    	 fnReportEvent "Pass","Project Source Field","Succesfully entered Project Source Field and value is::" &gb_TestDataDic.item("Project_source") ,False
'    	
'    	fnset  OracleIProcPageObj_1.WebEdit(projectSrc_xpath),gb_TestDataDic.item("Project_code"), "Project code"
'    	OracleIProcPageObj_1.WebEdit(projectSrc_xpath)
'       	fnReportEvent "Pass","Project Source Field","Succesfully entered Project code Field and value is::" &gb_TestDataDic.item("Project_code") ,False
'       	
'Const checkoutTable1 = "xpath:=//SPAN[@id='ItemTableRN']/TABLE[2]"
'Set OracleIProcPageObj_1 =  Browser("name:=Oracle iProcurement: Checkout").Page("title:=Oracle iProcurement: Checkout")
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
	If Err.number <> 0 Then             
              fnReportEvent "Fail","billable/non billable/Checkout","Failed to check out " & Err.description,false
             fn_selectBillableOrNonbillable = false
             Exit function
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
    

Call get_ApproverList()
 
	call fn_Click_fieldname(OracleIProcPageObj_1.WebButton(nextBtn2_xpath),"Next button")

'    If fn_Click(OracleIProcPageObj_1.WebButton(nextBtn2_xpath), "Next Button") Then       
'         fnReportEvent "Pass","Next Button","Click on the Next Button ",False
'    Else
'        fnReportEvent "Fail","Next Button","Next Button doesnt Exist",True    
'    End  If
	If Err.number <> 0 Then             
              fnReportEvent "Fail","verify Approvers","Failed to verify approvers " & Err.description,false
             fn_verifyApprovers = false
             Exit function
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
		OracleIProcPageObj_1.WebButton(reqSubmitBtn_xpath).Click
		fnReportEvent "Pass","Submit Button","Succesfully submit the  requisition request",False
    Else           
		fnReportEvent "Fail","Submit Button","Submit Button doesnt Exist " & Err.description,false
		fn_SubmitRequisition = false
		Exit function
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
    strTest="has been submitted."
    var_Reqno=OracleIProcPageObj_3.WebElement(reqText_xpath).GetROProperty("innertext")
    ReqNo=split(var_Reqno," ")
    'print ReqNo(1)
    
    strQuery="UPDATE [ExecutionResult$] SET Requisition_No='"&ReqNo(1)&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"

    'print strQuery
    Call fn_updateQuery(strQuery)
    'strExpectedStmt= "Requisition no
    strReqConfirmStmt=OracleIProcPageObj_3.WebElement(confirmationText_xpath).GetROProperty("innertext")
    If InStr(strReqConfirmStmt,strTest) >0 Then
    print ("Confirmation found.Requisition no "& ReqNo(1) &" has been submitted")
    fnReportEvent "Pass","Confirmation","Confirmation found.Requisition no "& ReqNo(1) &" has been submitted",True
    fn_verifyConfirmation=true
    else
    fnReportEvent "Fail","Confirmation","Confirmation not found",True
        fn_verifyConfirmation=false
End If
End Function
'Function fn_getRequestorSupervisor()
'
'	OracleIProcPageObj_1.Link(requesterApproverLink_xpath).Click
'            If OracleIProcPageObj_2.Exist(10) Then
'            	var_Row=OracleIProcPageObj_2.WebTable(approverDetailTable_xpath).GetROProperty("rows")
'                For i=2 to var_Row
'	                var_ApproverName=OracleIProcPageObj_2.WebTable(approverDetailTable_xpath).GetCellData(i,1)
'	                print var_ApproverName
'	                strQuery="UPDATE [ExecutionResult$] SET Approver1='"&var_ApproverName&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"
'					print strQuery
'    				Call fn_updateQuery(strQuery)
'	                var_ApproverEmail=OracleIProcPageObj_2.WebTable(approverDetailTable_xpath).GetCellData(i,3)
'	                    If var_ApproverName <>"" And var_ApproverEmail <>"" Then
'	     		               fnReportEvent "Pass","Requester Supervisor","Requester Supervisor= "&var_ApproverName&".",False
'	     		               OracleIProcPageObj_2.WebButton(returnBtn_xpath).Click
'	                    End If                        
'	           Next
'              Else
'                    fnReportEvent "Fail","Requester Supervisor","Requester Supervisor not available",True
'                    Exit Function
'            End IF
'    fn_getRequestorSupervisor=var_ApproverName
'End Function
'Function fn_getCommodityApprover()
'OracleIProcPageObj_1.Link(commodityApproverLink_xpath).Click
'            If OracleIProcPageObj_2.Exist(10) Then
'            var_Row=OracleIProcPageObj_2.WebTable(approverDetailTable_xpath).GetROProperty("rows")
'	             For i=2 to var_Row
'	                var_ApproverName=OracleIProcPageObj_2.WebTable(approverDetailTable_xpath).GetCellData(i,1)
'	                print var_ApproverName
'	                strQuery="UPDATE [ExecutionResult$] SET Approver2='"&var_ApproverName&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"
'					print strQuery
'    				Call fn_updateQuery(strQuery)
'	                var_ApproverEmail=OracleIProcPageObj_2.WebTable(approverDetailTable_xpath).GetCellData(i,3)
'	                    If var_ApproverName <>"" And var_ApproverEmail <>"" Then
'	                    fnReportEvent "Pass","Commodity Approvers","Commodity Approvers name= "&var_ApproverName&".",False    
'	                    OracleIProcPageObj_2.WebButton(returnBtn_xpath).Click
'	                    End If                        
'	             Next
'            Else
'                    fnReportEvent "Fail","Commodity Approvers","Commodity Approvers not available",True
'                    ExitTest
'            End If
'    fn_getCommodityApprover=var_ApproverName
'End Function

Function fn_ClickCheckOutButton()
Set OracleIProcPageObj= Browser("name:=Oracle iProcurement.*").Page("title:=Oracle iProcurement.*")
	If OracleIProcPageObj.WebButton(checkoutBtn_xpath).Exist(3) Then        
            OracleIProcPageObj.WebButton(checkoutBtn_xpath).Click
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
			Call fnNonCatlogAddToCart()
							
				

End Function

Function get_ApproverList()
Dim ObjLinks, ObjChild,intCount
'Dim arrLinkName()
Set ObjLinks = Description.Create

ObjLinks("micclass").Value = "Link"
'ObjLinks("xpath").Value="//*[@id='ApproverListRN']//tr[4]/td/Table[1]//a"
ObjLinks("html tag").Value="A"
Set ObjChild = OracleIProcPageObj_1.WebTable("xpath:=.//*[@id='ApproverListRN']/TBODY/TR[4]/TD/TABLE[1]").ChildObjects(ObjLinks)
'Set ObjChild = OracleIProcPageObj_1.ChildObjects(ObjLinks)

print ObjChild.Count


				For intCount = 0 To ObjChild.Count

                arrLinkName= ObjChild(intCount).GetROProperty("innertext")
'                If arrLinkName="Requester's Supervisor" Then
'                	arrLinkName="Requester's Supervisor"
'                End  If
				print arrLinkName
				Call fn_getApproverName(arrLinkName,intCount+1)
                Next

Set get_ApproverList=arrLinkName
End function

Function fn_getApproverName(arrLinkName,vCounter)
Count=vCounter
If OracleIProcPageObj_1.Exist(5) Then
appXpath="xpath:=.//a[contains(text(),"&chr(34) &arrLinkName& chr(34)&")]"
OracleIProcPageObj_1.Link(appXpath).Click
			If OracleIProcPageObj_2.Exist(10) Then
			var_Row=OracleIProcPageObj_2.WebTable("xpath:=.//SPAN[@id=""GroupHeader.ApprovalGroupTable""]/TABLE[2]").GetROProperty("rows")
				For i=2 to var_Row
				var_ApproverName=OracleIProcPageObj_2.WebTable("xpath:=.//SPAN[@id=""GroupHeader.ApprovalGroupTable""]/TABLE[2]").GetCellData(i,1)
	                strQuery="UPDATE [ExecutionResult$] SET Approver"&Count&"='"&var_ApproverName&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"
					print strQuery
    				Call fn_updateQuery(strQuery)
				var_ApproverEmail=OracleIProcPageObj_2.WebTable("xpath:=.//SPAN[@id=""GroupHeader.ApprovalGroupTable""]/TABLE[2]").GetCellData(i,3)
					If var_ApproverName <>"" And var_ApproverEmail <>"" Then
					fnReportEvent "Pass",""&arrLinkName&"",""&arrLinkName&" exists",False	
					OracleIProcPageObj_2.WebButton("xpath:=.//*[@id='ReturnButton']").Click
					End If
										
					Next
				
					Else
					fnReportEvent "Fail",""&arrLinkName&"",""&arrLinkName&" not available",True
					Exit Function
					
		End If
End  If
End Function



'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_RequisitionCreationGlobalAccValidation
'description:             fn_RequisitionCreation :  will create a new requisition will vaildate the global account and submit
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
'=============================================================

Function fn_GlobalAccountValidation()
	On Error Resume Next
	'Dim dict_GAInfo
	Set dict_GAInfo = CreateObject("Scripting.Dictionary")
	Call fn_Click_fieldname(OracleIProcPageObj_1.WebElement(showDetails_xpath),"Show Details Button")
	
	'Call fn_getChargeAccountInformation
	 Set dict_GAInfo = fn_getChargeAccountInformation
	'msgbox dict_GAInfo("Global Account")
	firstDigitString = Cstr(dict_GAInfo.Item("Global Account"))
	firstdigit = Left(firstDigitString,1)
	
'    	If OracleIProcPageObj_1.WebElement("N5___PoReqDistributionsVO___0:Charg").Exist(5) Then
'    	retglbaccnum = OracleIProcPageObj_1.WebElement("title:=Charge Account","index:=0").GetROProperty("innertext")
'    	fnReportEvent "Pass","Global Account Number","Global Account Number is present on page",False
'    	Else
'        fnReportEvent "Fail","Global Account Number","Global Account Number is not present on page",True    
'    	End  If
'    	retglbaccnumString = CStr(retglbaccnum)
'    	retChargeAccNum = Mid(retglbaccnumString,7,5)
'    	firstdigit = Left(retChargeAccNum,1)
    	
    	threshold_value = gb_TestDataDic.item("Threshold Amount")
    	req_threshold =  gb_TestDataDic.item("Qty") * gb_TestDataDic.item("Unit Price")
    	req_threshold_Str = Cstr(req_threshold)
    	If req_threshold_Str > threshold_value Then
	    	If firstdigit = "1" Then
	    		fnReportEvent "Pass","Global Account Number","Global Account Number value is :="&dict_GAInfo("Global Account"),False
	    	Else
	      	 fnReportEvent "Fail","Global Account Number","Failed to validate the Global Account Number value is :="&dict_GAInfo("Global Account"),True   
	    	End  If
	ElseIf req_threshold_Str < threshold_value Then
		If firstdigit = "5" Then
	    		fnReportEvent "Pass","Global Account Number","Global Account Number value is :="&dict_GAInfo("Global Account"),False
	    	Else
	      	 	fnReportEvent "Fail","Global Account Number","Failed to validate the Global Account number. The value is :="&dict_GAInfo("Global Account"),True   
	    	End  If
    	End If
'    	If firstdigit <> "5" Then
'    		fnReportEvent "Fail","Charge Account Number","Charge Account Number does not Start with 5",False
'    	Else
'        	fnReportEvent "Pass","Global Account Number","Charge Account Number Start with 5",True    
'    	End  If
   
End Function

Function fn_validateApprovalSubmitReq()
		Call fn_verifyApprovers
		Call fn_SubmitRequisition
End Function

Function fn_enterAndCheckoutNCRequest()
blnResultFlag=false
 On error resume next
  
   OracleIProcPageObj.Sync
   Call fn_NonCatlogAddToCart()
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
'	fnReportEvent "Pass","Item Description","Successfully enter Item Description and value is ::" & gb_TestDataDic.item("Item_Description") ,False
  
  	fnSet_FieldName OracleIProcPageObj.WebEdit(category_xpath),gb_TestDataDic.item ("Category"),"Category"
'	fnReportEvent "Pass","Category Description Field","Successfully  enter the Category and value is" & gb_TestDataDic.item ("Category") ,False
  
 	 fnSet_FieldName OracleIProcPageObj.WebEdit(quantity_xpath),gb_TestDataDic.item ("Qty"),"Qty"
'	fnReportEvent "Pass","Quantity TextBox Value","Successfully  enter the Quantity  and value entered is::" &gb_TestDataDic.item ("Qty"),False
	
    	fnSet_FieldName OracleIProcPageObj.WebEdit(unitOfMeasure_xpath), gb_TestDataDic.item ("Unit Of Measure"),"Unit of measure"
'	fnReportEvent "Pass","Unit Of measurement Text Box","Successfully  enter the Unit Of measurement  and value entered is::" & gb_TestDataDic.item ("Unit Of Measure") ,False
   
   	fnSet_FieldName OracleIProcPageObj.WebEdit(unitPrice_xpath),gb_TestDataDic.item ("Unit Price"),"Unit Price"
'	fnReportEvent "Pass","UnitPrice field","Successfully  enter the  UnitPrice  and value entered is::" & gb_TestDataDic.item ("Unit Price") ,False

	fnSet_FieldName OracleIProcPageObj.WebEdit(supplier_xpath), gb_TestDataDic.item ("Supplier Name"),"Supplier Name" 
'	fnReportEvent "Pass","Supplier field details","Successfully  enter the  Supplier detail  and value entered is::" & gb_TestDataDic.item ("Supplier Name") ,False
    
    	fnSet_FieldName OracleIProcPageObj.WebEdit(supplierSite_xpath),gb_TestDataDic.item ("Site"),"Supplier Site details"
'	fnReportEvent "Pass","Supplier Site details","Successfully  enter the  Supplier  site detail  and value entered is::" & gb_TestDataDic.item ("Site") ,False
   
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
 	
'       If OracleIProcPageObj.WebEdit(itemDesc_xpath).Exist(10)  Then
'       		 fnReportEvent "Pass", "Oracle Iproc page navigation status","Successfully navigated to Oracle Iproc Home page ",false
'		'Call fn_enterValuesonNonCatlogRequestPage
'		
'		Call fn_selectBillableOrNonbillable
'		blnResultFlag=true
'	else
'	 	fnReportEvent "Fail", "Oracle Iproc page navigation status","Not able to  navigate to Oracle Iproc Home page ",True
'		Exit Function
'	 End  If
	  	fn_enterAndCheckoutNCRequest = blnResultFlag
End Function

Function fn_vrfyapproversSubmitReq()
blnResultFlag=false
 On error resume next
  If OracleIProcPageObj.WebElement(approvalPagexpath).Exist(5)  Then
       	fnReportEvent "Pass", "Oracle Iproc approval page navigation status","Successfully navigated to Oracle Iproc approval page ",false
	Call fn_verifyApprovers
	Call fn_SubmitRequisition
	blnResultFlag=true
	fn_vrfyapproversSubmitReq = blnResultFlag
Else
	fnReportEvent "Fail", "Oracle Iproc Approval page navigation status","Not able to navigate to Oracle Iproc Approval page ",True
		Exit Function
End  If
	fn_vrfyapproversSubmitReq = blnResultFlag

End Function

Function fn_glbaccvalidSubmitReq()
blnResultFlag=false
 On error resume next
 
  If OracleIProcPageObj.WebElement(approvalPagexpath).Exist(5)  Then
       	fnReportEvent "Pass", "Oracle Iproc approval page navigation status","Successfully navigated to Oracle Iproc approval page ",false
	Call fn_verifyApprovers
	Call fn_GlobalAccountValidation
	Call fn_SubmitRequisition
	
	blnResultFlag=true

Else
	fnReportEvent "Fail", "Oracle Iproc Approval page navigation status","Not able to navigate to Oracle Iproc Approval page ",True
	Exit Function
End  If
	fn_glbaccvalidSubmitReq = blnResultFlag
End Function

Function fn_rcvalidationsubmitReq()
	Call fn_verifyApprovers
	Call fn_RCCentreverification
	Call fn_SubmitRequisition
End Function

Function fn_verifyLERCGlobalSubAccount
On Error Resume Next
If fn_Click(OracleIProcPageObj_1.WebElement(showDetails_xpath)) Then        
        fnReportEvent "Pass","Next Button","Succesfully Click on the Show Details Button",False
Else
        fnReportEvent "Fail","Next Button","Show Details Button doesnt Exist",True    
End  If

Set ChargeAccountInformation = fn_getChargeAccountInformation

'
'If OracleIProcPageObj_1.WebElement("title:=Charge Account","index:=0").Exist(5) Then
'    	OracleIProcPageObj_1.WebElement("title:=Charge Account","index:=0").Highlight
'    	retChargeAccountNum = OracleIProcPageObj_1.WebElement("title:=Charge Account","index:=0").GetROProperty("innertext")
'    	fnReportEvent "Pass","Charge Account Number","Charge Account Number is present on page",False
'Else
'        fnReportEvent "Fail","Charge Account Number","Charge Account Number is not present on page",True    
'End  If
'
'	splitChargeAccount = split(retChargeAccountNum,".")
'	
'	For i=lbound(splitChargeAccount) to ubound(splitChargeAccount)
'       msgbox  splitChargeAccount(i)
'   	Next
   	
   If ChargeAccountInformation("Legal Entity") = gb_TestDataDic.item("Legal Entity") Then
  	 fnReportEvent "Pass","Legal Entity","Legal entity is showing the expected value",False	
   Else
  	 fnReportEvent "Fail","Legal Entity","Legal entity is not showing the expected value",False	
   End If
  
   If ChargeAccountInformation("Global Account") = gb_TestDataDic.item("Global Account") Then
   	fnReportEvent "Pass","Global Account","Successfuly validated the Global Account and the value as per the expected data is:-"&gb_TestDataDic.item("Global Account"),False	
   Else
   	fnReportEvent "Fail","Global Account","Failed to validateb the Global Account is not showing the expected value",False	
   End If	
   
   If ChargeAccountInformation("Sub Account") = gb_TestDataDic.item("SubAccount") Then
   	fnReportEvent "Pass","Sub Account","Sub Account is showing the expected value",False	
   Else
   	fnReportEvent "Fail","Sub Account","Sub Account is not showing the expected value",False	
   End If	
   
   If ChargeAccountInformation("RC")  = gb_TestDataDic.item("RC") Then
   fnReportEvent "Pass","Sub Account","Sub Account is showing the expected value",False	
   Else
   fnReportEvent "Fail","Sub Account","Sub Account is not showing the expected value",False	
   End If	

End Function

Function fn_LERCGlobalSubAccountValidSubmitReq()
blnResultFlag=false
 On error resume next
 
  If OracleIProcPageObj.WebElement(approvalPagexpath).Exist(5)  Then
       	fnReportEvent "Pass", "Oracle Iproc approval page navigation status","Successfully navigated to Oracle Iproc approval page ",false
	Call fn_verifyApprovers
	Call fn_verifyLERCGlobalSubAccount
	Call fn_SubmitRequisition
	
	blnResultFlag=true

Else
	fnReportEvent "Fail", "Oracle Iproc Approval page navigation status","Not able to navigate to Oracle Iproc Approval page ",True
	Exit Function
End  If
	fn_LERCGlobalSubAccountValidSubmitReq = blnResultFlag
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
'		
  
	  	fnSet_FieldName OracleIProcPageObj.WebEdit(category_xpath),Split (gb_TestDataDic.item("Category"),"|")(index),"Category"
'		
		  
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
             Exit function
End If 

End  Function


Function fn_getChargeAccountInformation()
On Error Resume Next

Set ChargeAccountInfo = CreateObject("Scripting.Dictionary")

If OracleIProcPageObj_1.WebElement("title:=Charge Account","index:=0").Exist(5) Then
    	OracleIProcPageObj_1.WebElement("title:=Charge Account","index:=0").Highlight
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
	Exit function
	
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

	set objprojectsource  = OracleIProcPageObj_1.Webtable(checkoutTable).ChildObjects(objWebEdit)
	objprojectsource(irow-1).set value
	fn_WSSendKeys ("TAB")
	wait(1)
	
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

