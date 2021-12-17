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


