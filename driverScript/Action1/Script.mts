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
'************************************************************************************************
'Function Name:- fnRandomNumber
'Description:- Function to Create a Random Number of Any Length
'Input Parameters:- LengthOfRandomNumber
'Output Parameters:- 
'Created By:- 	MMC team
'************************************************************************************************
 Function fnRandomNumber(LengthOfRandomNumber)

Dim sMaxVal : sMaxVal = ""
Dim iLength : iLength = LengthOfRandomNumber

'Find the maximum value for the given number of digits
For iL = 1 to iLength
 sMaxVal = sMaxVal & "9"
Next
 sMaxVal = Int(sMaxVal)

'Find Random Value
Randomize
 iTmp = Int((sMaxVal * Rnd) + 1)
'Add Trailing Zeros if required
 iLen = Len(iTmp)
 fnRandomNumber = iTmp * (10 ^(iLength - iLen))

 End Function


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
