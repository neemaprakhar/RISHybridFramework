On error resume Next 
		
Call fn_intiatetest()

vTotalrowCount =  datatable.GetSheet("test_batch").GetRowCount
	
For irow = 1 To vTotalrowCount Step 1	
		Call fn_ReadTestScenario()
		
		Call fn_ExectueKeyword() @@ hightlight id_;_Browser("Oracle iProcurement: Shop").Page("Approval Group 2").WebButton("Return")_;_script infofile_;_ZIP::ssf12.xml_;_
		
		Call fn_StepResultTC()		
Next

DataTable.ExportSheet environment.Value("Str_testcase"), "test_batch","result"
fn_ExtentGenerateReport()



' OracleFormWindow("title:=Receipts.*").OracleButton("description:=Apply").Highlight
'fn_Click OracleFormWindow("title:=Receipts.*").OracleButton("description:=Apply")	

'label:=Apply
'Set orcreciptApplication = OracleFormWindow("title:=Applications.*")
'vtransactionNumber = "436050000075"
'If orcreciptApplication.GetROProperty("title") =  "Applications - QA8000"  Then
'	print " Able to navigaite to the Application"
'	orcreciptApplication.OracleTable("block name:=Table").Highlight
''	fn_EnterField orcreciptApplication.OracleTable("block name:=Table"),1,"Apply To",436050000075,"Apply To"
'	orcreciptApplication.OracleTable("block name:=Table").EnterField 1,"Apply To","436050000075"
'	
'End If


'Apply;Saved;Apply To;Installment;Apply Date;Amount Applied;Discount;Balance Due;Trans Currency;Customer Number;GL Date;Reversal GL Date;Allocated Receipt Amount;Cross Currency Rate;Exchange Gain/Loss;Activity;Application Reference Type;Application Reference Number;Application Reference Reason;Days Late;Line Number;Class;Type;[ ];( )

'
'
'
'
'
'
'Call get_ApproverList
'
'Set OracleIProcPageObj_1 = Browser("name:=Oracle iProcurement: Checkout").Page("title:=Oracle iProcurement: Checkout")
'
'Set ObjLinks = Description.Create
'
'ObjLinks("micclass").Value = "Link"
'ObjLinks("html tag").Value="A"
'
'Set ObjChild = OracleIProcPageObj_1.ChildObjects(ObjLinks)
'
'print ObjChild.Count
'
'
'				For intCount = 0 To ObjChild.Count
'				
'                			arrLinkName=ObjChild(intCount).GetROProperty("innertext")
'
'					print arrLinkName
''					Call fn_getApproverName(arrLinkName,intCount+1)
'					
'                Next
'
'Call fn_selectOU()
'
'Function fn_selectOU()
'	
'	On error Resume Next 
'	blnresultflag = false 	
'	 Set ObjCreateAddSiteCreation = Browser("name:=Create Address: Site Creation").Page("title:=Create Address: Site Creation") 	 
'   	Set ObjWebTable = Browser("name:=Create Address: Site Creation").Page("title:=Create Address: Site Creation").WebTable("xpath:=//span[@title='Operating Unit']/../../parent::tbody//parent::table[1]")    
'    	vstrEndRowCount = ObjWebTable.GetROProperty("rows")
'  	vstrOperatingUnit ="US OU"  'will fetch gb_dic object
'  	
'  	vasciinumber = Asc(letf(vstrOperatingUnit,1) )
'  	
'  	if vasciinumber <= 80 then
'  	 	print " will do the sorting on the operating unit--> click on the OU objetc" 
'  	 End If 
'Do 
'	For  RowIndex = 1 To vstrEndRowCount
'             vstrCellData = ObjWebTable.GetCellData(RowIndex,3)  
''		print vstrCellData             
'                 If trim(vstrCellData)=vstrOperatingUnit then
'                     Set ObjSelectCheckBox = ObjWebTable.ChildItem(RowIndex,0,"WebCheckBox",0)
'                     ObjSelectCheckBox.Set "ON" 
'                     fnReportEvent "Pass"," enabled the checkbox","Succefully enabled the checkbox and select checkbox value is =" & vstrOperatingUnit ,False
'			blnresultflag= true                     
'                     Exit do
'                 End If         
'             
'             if lastpagecounter then
'               	fnReportEvent "Fail"," Unable to select OU","Fail to select OU and  value is =" & vstrOperatingUnit ,true
'              	Exit do
'              End if
'        Next
'        
'  Set ObjWebTable = Browser("name:=Create Address: Site Creation").Page("title:=Create Address: Site Creation").WebTable("xpath:=//span[@title='Operating Unit']/../../parent::tbody//parent::table[1]")
''  ObjCreateAddSiteCreation.Image("alt:=Select to view next set","index:=0").Click
' 	If  ObjCreateAddSiteCreation.Image("alt:=Select to view next set","index:=0").GetROProperty("Visible") = true Then	     
'   		ObjCreateAddSiteCreation.Image("alt:=Select to view next set","index:=0").Click
'   		nextagecounter = true
'	ElseIf  ObjCreateAddSiteCreation.Image("alt:=Next functionality disabled","index:=0").Exist = true Then			
'   		vstrEndRowCount = ObjWebTable.GetROProperty("rows")
'   		lastpagecounter =true
'	End  If
'
'    
'Loop While ( nextagecounter = true or lastpagecounter =true)
'
'End Function
'    
' 
'
'vstrOperatingUnit ="UK  OU"  'will fetech gb_dic object
'  	
'  	vasciinumber = Asc(left(vstrOperatingUnit,1) )
'  	
'  	if vasciinumber >= 80 then
'  	 	print " will do the sorting on the operating unit--> click on the OU objetc" 
'  	 End If 
'   
''    3
'Set resultDic =  fn_getExecutionResultInDic(vgstrTestCaseExec_id)
'resultDic("Legal_entity")
'value1 = fn_getExecutionResultData("Transaction_No")
'print value1
'
'
'Function fn_getExecutionResultData(vgstrTestCaseExec_id,pfieldname)
'On Error Resume Next  
'Dim objCon,objRecSet
'vgstrTestCaseExec_id = "GSI.O2C.AR.SA.006"
'
'strQuery= "Select * from [ExecutionResult$] where TC_ID='" & vgstrTestCaseExec_id & "'"  & " Order by Start_Date DESC"
'print "query value is =" & strQuery
'
''strFileName =  mid(environment("TestDir"),1,InStrRev(environment("TestDir"),"\")) & "TestData\" &  environment("TestDataFileName")  & ".xls"
'Print strFileName
'strFileName =  "C:\Users\U1227650\OneDrive - MMC\Desktop\RIS_TestData_1.xls"
'    Set objCon=  fn_getDBconnection(strFileName)
'    Set objRecSet = fn_getRecordset(objCon,strQuery)
'    fn_getExecutionResultData=fn_getColValueFromRecorset(objRecSet,pfieldname)            
'    
'    If Err.number <> 0 Then             
'         print "check correct name exist in the recordset :fn_getExecutionResultData"
'         Exit function
'      End If
'End Function
'
'
'Function fn_getExecutionResultInDic()
'
'On Error Resume Next  
'Dim objCon,objRecSet
'gstrTestCaseExec_id = "GSI.O2C.AR.SA.006"
'
'strQuery= "Select * from [ExecutionResult$] where TC_ID='" & gstrTestCaseExec_id & "'"  & " Order by Start_Date DESC"
'print "query value is =" & strQuery
'
''strFileName =  mid(environment("TestDir"),1,InStrRev(environment("TestDir"),"\")) & "TestData\" &  environment("TestDataFileName")  & ".xls"
'Print strFileName
'strFileName =  "C:\Users\U1227650\OneDrive - MMC\Desktop\RIS_TestData_1.xls"
'    Set objCon=  fn_getDBconnection(strFileName)
'    Set objRecSet = fn_getRecordset(objCon,strQuery)
'
'Set objResultDic =  CreateObject("Scripting.Dictionary")
'Do  While Not objRecSet.EOF
'	    For I=0 To objRecSet.Fields.Count-1	
'	      		If Not IsNull(objRecSet.Fields(I).value) Then	      			
'	      			objResultDic.Add objRecSet.Fields(I).name,objRecSet.Fields.Item(i)
'	      		else	      			
'				objResultDic.Add objRecSet.Fields(I).name, "Null"	      		
'	      		End If
'	    Next	 
''	    fn_getExecutionResultInDic =objResultDic
'	Exit Do 
'    	objRecSet.MoveNext
'  Loop	
' 
'	
'	
'	If Err.number <> 0 Then             
'         print "check correct name exist in the recordset :fn_getExecutionResultInDic"
'         Exit function
'      End If
'
'
'End Function
'
'
'
'Function fn_getColValueFromRecorset(objRecSet,pfieldName)
'
'	If Not IsNull(objRecSet.Fields(pfieldName).value) Then		
'		print objRecSet.Fields(pfieldName).value
'		fn_getColValueFromRecorset = objRecSet.Fields(pfieldName).value
'	else
'		fn_getColValueFromRecorset = "Null"	      		
'	End If
'
'    
'End Function
'
'

'	counter = 1
'	Do
''		wait(1)
'		 counter=counter+1
'		 print counter
''		If objParent.Exist(1) Then		
'''			fn_exist = true
''		Exit do 
''		Else 
''			fn_exist = false
''		End If
'	Loop while counter<20


'Set OracleFormFuncationsList = OracleFormWindow("title:=Navigator.*").OracleTabbedRegion("label:=Functions").OracleList("description:=Function List")
'Set list = OracleFormFuncationsList.GetAllROProperties
'print list(1)
' print OracleFormWindow("title:=Navigator.*").OracleTabbedRegion("label:=Functions").OracleList("description:=Function List").ToString
'Print OracleFormFuncationsList(1).toString
'OracleFormFuncationsList.Select("+  Transactions")
'OracleFormFuncationsList.Activate("+  Transactions")
'OracleFormFuncationsList.Highlight
'For i = 1 To 1 Step 1
'	
'	OracleFormFuncationsList.
'Next
'OracleFormFuncationsList.Activate(4)
'Call fn_NavigatorOraclePage("Transactions","Transactions")
'
'Call fn_NavigateOraclePage("Transactions","transactions-->transaction")
'Set OracleNavigatorForm = OracleFormWindow("title:=Navigator.*").OracleTabbedRegion("label:=Functions")




'OracleFormWindow("title:=Navigator.*").OracleTabbedRegion("label:=Functions").OracleButton("description:=Collapse All").highlight
'OracleFormWindow("title:=Navigator.*").OracleTabbedRegion("label:=Functions").OracleButton("description:=Collapse All").click
'
'OracleTabbedRegion("label:=Functions").

'call fn_switchResponsibility("IDN AR Transaction Approver")


'
'Function fn_switchResponsibility(respName)
'On error resume next
'fn_switchResponsibility = false       
'    If (fn_exist (OracleFormWindow("title:=Navigator.*"))) Then
''        OracleFormWindow("title:=Navigator.*").SelectMenu "File->Switch Responsibility..." 	
'        Call fn_SelectMenu(OracleFormWindow("title:=Navigator.*"),"SwitchResponsibility")
'	If OracleListOfValues("title:=Responsibilities").Exist(5) Then
'		OracleListOfValues("title:=Responsibilities").Select respName	 	      
'         Else 
'        	fnReportEvent "Fail", "Switch Responsibility","Fail to naivigate to Switch Responsibility",true        	
'    	End If    
'    End  if 	
''    validate if responsibilty is selected correctly 
'title =  OracleFormWindow("title:=Navigator.*").GetROProperty("title")
'	If Instr(1,title,respName) > 1 Then
'		fnReportEvent "Pass", "Navigator Page Status","Navigator Page is displaying and User is able to switch the Responsibility to "&respName,false
'		fn_switchResponsibility = true
'	else
'		fnReportEvent "Fail", "Navigator Page Status","Navigator Page is not displaying or Responsibility is not present for that user "& respName ,true
'	End If
'        
'        If Err.number <> 0 Then             
'              fnReportEvent "Fail", "Navigator Page Status","Navigator Page is not displaying or Responsibility is not present for that user "& respName ,true
'              fn_switchResponsibility = false             	
'            Exit function
'        End If
'End function



'msgbox OracleListOfValues("title:=Responsibilities").Select("IDN AR Transaction Approver11")

'fn_switchResponsibility = fn_Navigator("Complete Transaction","")
'print fn_switchResponsibility


