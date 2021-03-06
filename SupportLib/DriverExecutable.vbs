Public blnStartFlag,blnEndTestcase,blnSkipKeyword

blnStartFlag = true

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_intiatetest 
'description: 			fn_intiatetest : It will intitate all the pre-requiste required for test case
'Parameter				
'Function call ::		
'Return Type -null
'*************************************************************************
'=============================================================

Function fn_intiatetest()

	On error resume Next 

'Will load the .ini file 
	strpath = mid(Environment("TestDir"),1,instrRev(Environment("TestDir"),"\")) & "EnvironmentVar\Enviornment.ini"
	Environment.LoadFromFile strpath

	If err.number<>0 Then
		msgbox "Failed to load the Environment.ini file"
		ExitTest
	End If
	
	testdatafilename = "testdata"
	vtestScenarioFileName = environment("TestScenarioFileName")
	environment.Value("Str_testcase") = mid(environment("TestDir"),1,InStrRev(environment("TestDir"),"\")) & "Test_Scenario\" & vtestScenarioFileName & ".xlsx"

	'print environment.Value("Str_testcase")
'''''''Report Intalization
	Call fn_CreateLogfolder()
	call fn_ExtentReporterHeader()
'''''''Importing the test case sheet into the  dataTabe 
	datatable.AddSheet "test_batch"
	DataTable.ImportSheet environment.Value("Str_testcase"), "test_batch","test_batch"

        If err.number <> 0 Then
		msgbox err.Description & "Failed to load all  the  pre-requiste required for the test case"
		Exittest	
	End If

End Function

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_ReadTestScenario 
'description: 			fn_ReadTestScenario : it will used to read the test case and will update the global variable to execute test case
'Parameter				globalvariable used :gstrTestCaseExec_id,gstrTestcasename,gstrStep_name,gstrkeyword,gstrIdentifer1,gstrIdentifer2,
'						gstrIdentifer3,gstrStepResult
'Function call ::		
'Return Type -null
'*************************************************************************
'=============================================================

Function fn_ReadTestScenario()

 On error resume next 


	totalTC_Count =  datatable.GetSheet("test_batch").GetRowCount
	gstrStepResult  = datatable.Value("Result","test_batch")	
'	 start  reading test scenario file from  from here
	If blnStartflag   Then	
'	add the condition to run specific test case		
				datatable.SetCurrentRow (1)
				gstrTestCaseExec_id =  datatable.Value("TestCase_Id","test_batch")
				gstrPrevTCID =gstrTestCaseExec_id
				gstrTestcasename =  datatable.Value("TestCase_Name","test_batch")
				blnStartflag = false		   			
	
	ElseIf lcase(gstrStepResult)="fail" or gstrStepResult= "No run"  Then	
			datatable.GetSheet("test_batch").SetNextRow

		If (gstrTestCaseExec_id  =  datatable.Value("TestCase_Id","test_batch"))  or (datatable.Value("TestCase_Id","test_batch") ="") Then			
				blnEndTestcase = true	
				Exit Function					
			else
				blnEndTestcase = false   ' add condition when will execute the next test case	
				gstrTestCaseExec_id = datatable.Value("TestCase_Id","test_batch") ''	Re-assign value when moving from failed scenario to next  TC 
				gstrPrevTCID = gstrTestCaseExec_id
				 gstrTestcasename =  datatable.Value("TestCase_Name","test_batch")				 
			End If
			
			If totalTC_Count = datatable.GetSheet("test_batch").GetCurrentRow Then
				blnEndofAlltestCase =true								
			End If			
	else
				datatable.GetSheet("test_batch").SetNextRow
'			adding below condition for the passed scenario 
			If ( gstrTestCaseExec_id <> datatable.Value("TestCase_Id","test_batch") ) and  (datatable.Value("TestCase_Id","test_batch") <>"")Then				
				gstrTestCaseExec_id = datatable.Value("TestCase_Id","test_batch")
				gstrPrevTCID = datatable.Value("TestCase_Id","test_batch") 
				 gstrTestcasename =  datatable.Value("TestCase_Name","test_batch")
			End If
			
			If TotalTC_rowcount = datatable.GetSheet("test_batch").GetCurrentRow Then
				blnEndofAlltestCase = true
			End If					
	End If
		
	gstrStep_name =  datatable.Value("Steps","test_batch")
	gstrkeyword =  datatable.Value("Keyword","test_batch")
	
	'print "Test case id:=" & gstrTestCaseExec_id  &  "  Test case name := " & gstrTestcasename & "  Keyword :=" & gstrkeyword
 	
	gstrTdIdentifer1 =  datatable.Value("Testdata_Identifer1","test_batch")
	gstrTdIdentifer2  = datatable.Value("Testdata_Identifer2","test_batch")
	gstrStepResult =  datatable.Value("Result","test_batch")	
	
'if keyword are unable to read from excel due to some issue this flag will return true and skip execution
	If gstrkeyword =  "" Then
		blnSkipKeyword = true
	End If
	
	If err.number<>0 Then
		'print  err.Description & "Fn :fn_readtestscenario"
		On Error GoTo 0
		ExitTest
	End If
	
End Function

'=============================================================
'*************************************************************************
'Created By - MMC team	
'Creation Time & Date:    09/02/2021
'Name - 				fn_ExectueKeyword 
'description: 			fn_ExectueKeyword : It will exectue the keyword. if blnEndTestcase flag status is false then it  will exit the function 		
'Parameter::				
'Function call ::		
'Return Type -      		true or false
'*************************************************************************
'=============================================================
Function fn_ExectueKeyword()

On error resume next

Dim ptr_keyword
''	if  blnEndTestcase & blnSkipKeyword --> 1. Will be true when test step is failed  2.In keyword column blank value is passed
		If (blnSkipKeyword ) or (blnEndTestcase ) Then				
			Exit function
		End If
		
	'print  "will execute keyword::" & gstrkeyword
	gstrkeyword = Ucase(gstrkeyword)	
	Set ptr_keyword =  GetRef(gstrkeyword)
	
		If err.number <> 0 Then				
				gblnStepExecutionResult = false
				On Error GoTo 0
		Else
' it will call the function which is assigned to it and will return the value true or false		
				gblnStepExecutionResult = ptr_keyword						
		End if 
	gstrkeyword =""
End Function

'=============================================================
'*************************************************************************
'Created By - 		
'Creation Time & Date:  09/02/2021
'Name - 				fn_StepResultTC 
'description: 			fn_StepResultTC : will update test result in the data sheet
'Parameter			
'Function call ::			
'Return Type -null
'*************************************************************************
'=============================================================

Function fn_StepResultTC()
		
	If (gblnStepExecutionResult)  Then
		datatable.Value("Result","test_batch") = "Pass"
	Else 
		If blnEndTestcase  Then  'or datatable.Value("Result","test_batch") = "No run")
			datatable.Value("Result","test_batch") = "No run"	
		else
			datatable.Value("Result","test_batch") = "Fail"		
		End  if 	
	End If

	blnEndTestcase= false
	blnSkipKeyword =false	
			
End Function

Function fn_batchrunStatus()

On Error resume Next 
	strFileName =  mid(environment("TestDir"),1,InStrRev(environment("TestDir"),"\")) & "TestData\" &  environment("TestDataFileName")  & ".xls"
	vsheetname = environment("TestDataSheetName")
	 vtcIdentifier = gstrTdIdentifer1  &"|" & Split(environment("Legal_entity"),"|")(0)
	
	'	creating dictionary object 	
	Set batchRunStatus =  fn_createDataDictionary(strFileName,vsheetname,vtcIdentifier)
	fn_batchrunStatus = ucase(batchRunStatus("Run_Status")) 
	If err.number<>0 Then
		'print  err.Description & "Fn :fn_batchrunStatus"
		On Error GoTo 0
		
	End If
	
End Function

