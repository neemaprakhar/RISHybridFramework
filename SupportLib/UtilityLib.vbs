
'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fn_createDataDictionary 
'description: 			It will create the dictionary object
'Parameter				strFileName,vsheetname,vtcIdentifier
'Function call ::		       
'Return Type -           	dictionary object
'*************************************************************************
'=============================================================

Function fn_createDataDictionary(strFileName,vsheetname,vtcIdentifier)

Dim objCon, objRecSet

	Set objCon=  fn_getDBconnection(strFileName)
	strQuery = fn_getQuery(vsheetname,vtcIdentifier)
	Set objRecSet = fn_getRecordset(objCon,strQuery)	
	Set objDict = CreateObject("Scripting.Dictionary")
	
	fieldCount = objRecSet.Fields.Count
		While objRecSet.EOF = False
		    For i = 0 To fieldCount - 1
		        vkey= objRecSet(i).name 
		         vitem= objRecSet.Fields.Item(i)
		          objDict.Add trim(vkey),trim(vitem)
		    Next
		    objRecSet.moveNext
		Wend
	
		Set objRecSet= Nothing
		objCon.Close
		Set objCon = Nothing

If Err.number <> 0 Then
	fn_createDataDictionary = false
     print "failed to create the data dictionary object"
     Exit Function
End If

Set fn_createDataDictionary=objDict

End Function


'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fn_getDBconnection 
'description: 			It will create the database connection & will treat excel as database
'Parameter				strFileName
'Function call ::		       
'Return Type -           	object = db connection
'*************************************************************************
'=============================================================

Public Function fn_getDBconnection(strFileName)
Dim objCon
Set objCon= CreateObject("ADODB.Connection")
objCon.connectionstring= "PROVIDER=MSDASQL;DRIVER={Microsoft Excel Driver (*.xls)};DBQ="& strFileName &";ReadOnly=False"
objCon.Open
if objCon.State<>1 Then 
print "Connection not established"
End  if 
Set fn_getDBconnection=objCon
    
End Function


'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fn_getRecordset 
'description: 			It will retun the database record set 
'Parameter				strFileName
'Function call ::		       
'Return Type -           	object = recordset
'*************************************************************************
'=============================================================

Public Function fn_getRecordset(objCon,strQuery)
 
 On Error Resume Next 
 
Const adOpenStatic = 3
Const adLockOptimistic = 3    
Set objRecSet = CreateObject("ADODB.Recordset")  
objRecSet.CursorType=3
objRecSet.Open strQuery, objCon,adOpenStatic,adLockOptimistic
'print"Record Count" &  objRecSet.RecordCount
Set fn_getRecordset=objRecSet


If Err.number <> 0 Then
     msgbox "Query used in the connection is incorrect "  & err.description      
End If

End Function


'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fn_getQuery 
'description: 			It will retun the query
'Parameter				strFileName
'Function call ::		       
'Return Type -           	String = Sql Query for the record set
'*************************************************************************
'=============================================================

Function fn_getQuery(vsheetname,vtcIdentifier1)
'strQuery = "Select * from ["&vsheetname& "$] where TC_ID='" &vtcIdentifier& "'"
	arrtcIdentifier= Split(vtcIdentifier1,"|")	 
	If  UBound(arrtcIdentifier) = 0   Then
		strQuery = "Select * from ["&vsheetname& "$] where TC_ID='" &arrtcIdentifier(0)& "'"
	ElseIf UBound(arrtcIdentifier) =1  Then
		strQuery = "Select * from ["&vsheetname& "$] where TC_ID='" &arrtcIdentifier(0) & "' and Legal_entity ='" &	arrtcIdentifier(1) & "'"			
	End If		
	fn_getQuery = strQuery 
End Function

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fn_updateQuery 
'description: 			fn_updateQuery :  will update date in the execution result table
'Parameter				will update date in the execution result table
'Return Type -           	NA
'*************************************************************************
'=============================================================

Public Function fn_updateQuery(strQuery)
Dim objCon
On Error Resume Next	
strFileName =  mid(environment("TestDir"),1,InStrRev(environment("TestDir"),"\")) & "TestData\" &  environment("TestDataFileName")  & ".xls"
'Print strFileName
	Set objCon=  fn_getDBconnection(strFileName)
	objCon.Execute strQuery
	If Err.number <> 0 Then             
             print "Unable to execute query"
            
             Exit function
      End If
	Set objCon=Nothing 
End Function
'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  20/10/2021
'Name - 				fn_getExecutionResultData 
'description: 			fn_getExecutionResultData :  will fetech the single field from execution result table
'Parameter				will update date in the execution result table
'Return Type -           	NA
'*************************************************************************
'=============================================================

Function fn_getExecutionResultData(vgstrTestCaseExec_id,pfieldname)
On Error Resume Next  
Dim objCon,objRecSet


strQuery= "Select * from [ExecutionResult$] where TC_ID='" & vgstrTestCaseExec_id & "'"  & " Order by Start_Date DESC"
'print "query value is =" & strQuery

strFileName =  mid(environment("TestDir"),1,InStrRev(environment("TestDir"),"\")) & "TestData\" &  environment("TestDataFileName")  & ".xls"

    Set objCon=  fn_getDBconnection(strFileName)
    Set objRecSet = fn_getRecordset(objCon,strQuery)
    fn_getExecutionResultData=fn_getColValueFromRecorset(objRecSet,pfieldname)            
    
    If Err.number <> 0 Then             
         print "check correct name exist in the recordset :fn_getExecutionResultData"
         Exit function
      End If
End Function

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  20/10/2021
'Name - 				fn_getExecutionResultData 
'description: 			fn_getExecutionResultData :  will update date in the execution result table
'Parameter				will update date in the execution result table
'Return Type -           	NA
'*************************************************************************
'=============================================================

Function fn_getColValueFromRecorset(objRecSet,pfieldName)

	If Not IsNull(objRecSet.Fields(pfieldName).value) Then		
'		print objRecSet.Fields(pfieldName).value
		fn_getColValueFromRecorset = objRecSet.Fields(pfieldName).value
	else
		fn_getColValueFromRecorset = "Null"	      		
	End If

    
End Function


'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  20/10/2021
'Name - 				fn_getExecutionResultInDic 
'description: 			fn_getExecutionResultInDic :  will return the dic object from execution result sheet
'Parameter				will update date in the execution result table
'Return Type -           	dic object 
'*************************************************************************
'=============================================================


Function fn_getExecutionResultInDic(vgstrTestCaseExec_id)

On Error Resume Next  
Dim objCon,objRecSet
'gstrTestCaseExec_id = "GSI.O2C.AR.SA.006"

strQuery= "Select * from [ExecutionResult$] where TC_ID='" & vgstrTestCaseExec_id & "'"  & " Order by Start_Date DESC"
'print "query value is =" & strQuery

strFileName =  mid(environment("TestDir"),1,InStrRev(environment("TestDir"),"\")) & "TestData\" &  environment("TestDataFileName")  & ".xls"
Print strFileName
'strFileName =  "C:\Users\U1227650\OneDrive - MMC\Desktop\RIS_TestData_1.xls"
    Set objCon=  fn_getDBconnection(strFileName)
    Set objRecSet = fn_getRecordset(objCon,strQuery)

Set objResultDic =  CreateObject("Scripting.Dictionary")
Do  While Not objRecSet.EOF
	    For I=0 To objRecSet.Fields.Count-1	
	      		If Not IsNull(objRecSet.Fields(I).value) Then	      			
	      			objResultDic.Add objRecSet.Fields(I).name,objRecSet.Fields.Item(i)
	      		else	      			
				objResultDic.Add objRecSet.Fields(I).name, "Null"	      		
	      		End If
	    Next	 
'	    fn_getExecutionResultInDic =objResultDic
	Exit Do 
    	objRecSet.MoveNext
  Loop	
 
	
	
	If Err.number <> 0 Then             
         print "check correct name exist in the recordset :fn_getExecutionResultInDic"
         Exit function
      End If


End Function



'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fn_GetROPropertyValue 
'description: 			fn_GetROPropertyValue :  will return the property value to the function
'Parameter				
'Function call ::		       
'Return Type -null           String =will  return  property value
'*************************************************************************
'=============================================================

Public Function fn_GetROPropertyValue(objParent)

	Reporter.Filter = 3
	If objParent.getroproperty("innertext") <> "" Then
		fn_GetROPropertyValue = objParent.getroproperty("innertext")		
	ElseIf objParent.getroproperty("name") <> "" Then
		fn_GetROPropertyValue = objParent.getroproperty("name")		
	ElseIf objParent.getroproperty("value") <> "" Then
		fn_GetROPropertyValue = objParent.getroproperty("value")	
	ElseIf objParent.getroproperty("text") <> "" Then
		fn_GetROPropertyValue = objParent.getroproperty("text")
	ElseIf objParent.getroproperty("title") <> "" Then
		fn_GetROPropertyValue = objParent.getroproperty("title")
	ElseIf objParent.getroproperty("html id") <> "" Then
	    fn_GetROPropertyValue = objParent.getroproperty("html id")
	Else
		fn_GetROPropertyValue = objParent.getroproperty("class")		
	End If
	'Reporter.Filter = 1
End Function
'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fn_Click_fieldname
'description: 			fn_Click :  click on the object  and event will be capture in the report
'Parameter			       objParent,strFieldName
'Return Type -null           
'*************************************************************************
'=============================================================

Function fn_Click_fieldname(objParent,vFieldName)
	
	If objParent.Exist(10) Then
'	   	objParent.highlight	
		objParent.Click	
		fn_Click_fieldname = true
		Call fnReportEvent ("Pass",vFieldName & " object","Succesfully click on the " & vFieldName ,true)
	Else			
		Call fnReportEvent ("FAIL","Fail to click on :: "& vFieldName &" Object", "Function ::fn_Click = " & objParent.getROproperty("micclass") & " does not exists, Please check",true)
		fn_Click_fieldname = false
	End If	
	
End Function

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fn_Click
'description: 			fn_Click :  click on the object 
'Parameter			       objParent
'Return Type -null           
'*************************************************************************
'=============================================================
Function fn_Click(objParent)
	
	If objParent.Exist(10) Then	   	
		objParent.Click	
		fn_Click = true		
	Else		
		Call fnReportEvent ("FAIL","Fail to click on "& objParent.getROproperty("micclass") &" Object", "Function ::fn_Click = " & objParent.getROproperty("micclass") & " does not exists, Please check",true)
		fn_Click = false
	End If	
	
End Function
'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fnSet 
'description: 			fnSet :  Set the text in a web object
'Parameter				objParent,strValue
'Function call ::		       
'Return Type -           	NA
'*************************************************************************
'=============================================================
Public Function fn_Set(objParent,strValue)

'	vPropertyValue = ucase(fn_GetROPropertyValue(objParent))
'	vClass =Ucase(objParent.getROproperty("micclass"))
	If IsNull(strValue) Then
		strValue = ""
	End If
	If objParent.Exist(10) Then
	    	objParent.Set strValue				
'		fnReportEvent "PASS","Set Value in "& vPropertyValue & " object", vPropertyValue& " is exist and expected Value is entered:- " & strValue, true
	Else
		Call fnReportEvent ("FAIL","Set Value in "& strFieldName &" Object", "Function ::fn_Set = " & strFieldName& " does not exists, Please check",false)
	End If
'	In futute might need to add the validation code
End Function




'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fnSet_FieldName
'description: 			fnSet_FieldName :  Set the text in a web object
'Parameter				objParent,strValue,strFieldName
'Function call ::		       
'Return Type -           	NA
'*************************************************************************
'=============================================================
Public Function fnSet_FieldName(objParent,strValue,strFieldName)

'	vPropertyValue = ucase(fn_GetROPropertyValue(objParent))
'	vClass =Ucase(objParent.getROproperty("micclass"))
	If IsNull(strValue) Then
		strValue = ""
	End If
	If objParent.Exist(10) Then
	    	objParent.Set strValue	   	
		fnReportEvent "PASS",strFieldName& "FieldName","Successfully enter the " &  strFieldName & "and  Value is:= "& strValue , false
	Else
		fnReportEvent "FAIL",strFieldName& "FieldName","Fail to enter  " &  strFieldName & "and  Value is:= "& strValue , true
	End If
'	In futute might need to add the validation code
End Function
'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  29/09/2021
'Name - 				fn_Select
'description: 			fn_Select :  Select Value from Oracle Table Weblist
'Parameter			       objParent,strValue,strFieldName
'Return Type -null           
'*************************************************************************
'=============================================================
Public Function fn_Select(objParent,strValue,strFieldName)
	
	vPropertyValue = ucase(fn_GetROPropertyValue(objParent))
	vClass =Ucase(objParent.getROproperty("micclass"))
	If IsNull(strValue) Then
		'strValue = ""
		Exit Function
	End If
	If objParent.Exist(10) Then	
	    	objParent.select  strValue				
		fnReportEvent "Pass","Selected Value for the  "& strFieldName & " list", "Succesfully selected the value from the list and value selected is := " & strValue,false
		fn_Select = true
	Else
		Call fnReportEvent ("FAIL","Selected Value for the  " & strFieldName & " list", "Function ::fn_Select = " & strFieldName & " does not exists, Please check",true)
		fn_Select = false		
	End If
'	In futute might need to add the validation code
End Function
Function fn_SelectWeblist(objParent,strValue,strFieldName)

	If IsNull(strValue) or strValue="" Then
		fnReportEvent "FAIL","Value passed in blank ", "Function ::fn_Select = " & strFieldName & " does not exists, Please check",false
		Exit Function
	End  If
	If objParent.Exist(2) Then
		fn_Click objParent
	    	objParent.select  strValue				
		fnReportEvent "PASS","Selected Value for the  "& strFieldName & " list", "Succesfully selected the value from the list and value selected is := " & strValue,false
		fn_SelectWeblist = true
	Else
		Call fnReportEvent ("FAIL","Selected Value for the  " & strFieldName & " list", "Function ::fn_Select = " & strFieldName & " does not exists, Please check",true)
		fn_SelectWeblist = false		
	End If
End Function

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  29/09/2021
'Name - 				fn_SelectWeblist
'description: 			fn_SelectWeblist :  Select Value from Browser dropdown
'Parameter			       objParent,strValue,strFieldName
'Return Type -null           
'*************************************************************************
'=============================================================
Function fn_SelectWeblist(objParent,strValue,strFieldName)

	If IsNull(strValue) or strValue="" Then
		fnReportEvent "FAIL","Value passed in blank ", "Function ::fn_Select = " & strFieldName & " does not exists, Please check",false
		Exit Function
	End  If
	If objParent.Exist(2) Then
		fn_Click objParent
	    	objParent.select  strValue				
		fnReportEvent "PASS","Selected Value for the  "& strFieldName & " list", "Succesfully selected the value from the list and value selected is := " & strValue,false
		fn_SelectWeblist = true
	Else
		Call fnReportEvent ("FAIL","Selected Value for the  " & strFieldName & " list", "Function ::fn_Select = " & strFieldName & " does not exists, Please check",true)
		fn_SelectWeblist = false		
	End If
'	In futute might need to add the validation code
	
End Function

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fn_exist 
'description: 			fn_exist :  will check the check existence of the object 
'Parameter				objParent
'Function call ::		       
'Return Type -           	true or false 
'*************************************************************************
'=============================================================

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  29/09/2021
'Name - 				fn_exist
'description: 			fn_exist : Check if object exists
'Parameter			       objParent
'Return Type -null           
'*************************************************************************
'=============================================================
Public Function fn_exist(objParent)
	counter = 1
	Do
		wait(1)
		counter=counter+1
		If objParent.Exist(1) Then		
			fn_exist = true
			Exit do 
		Else 
			fn_exist = false
		End If
	Loop while counter<90
End  Function
'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fn_Highlight 
'description: 			fnSet :  Set the text in a web object
'Parameter				highlight the objparent
'Return Type -           	NA
'*************************************************************************
'=============================================================

Public Function fn_Highlight(objParent)

	If  objParent.Exist(1) Then
		objParent.highlight		
'		fnReportEvent "PASS","Highlight "&vPropertyValue&" Object of class "&vClass,vPropertyValue&" Object of class "&vClass&" Exist and Highlighted Successfully","NO"
	Else
		fnReportEvent "FAIL","Highlight "&vPropertyValue&" Object of class "&vClass,vPropertyValue&" Object of class "&vClass&" Does not Exist, Please Check.","YES"		
	End If

End Function

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  29/09/2021
'Name - 				fn_Enter
'description: 			fn_Enter :  Enter Value in Oracle Form Field & Display in Report
'Parameter			       objParent,strValue,strFieldName
'Return Type -null           
'*************************************************************************
'=============================================================
Public Function fn_ReportEnter(objParent,strValue,strFieldName)

	If IsNull(strValue) Then
		strValue = ""
	End If
	If objParent.Exist(5) Then
	    	objParent.Enter strValue
		Call fnReportEvent ("Pass",strFieldName& " : FieldName ","Successfully entered value " & strValue &  " in " & strFieldName & " Field",false)		
	Else
		Call fnReportEvent ("Fail",strFieldName& " : FieldName","Unable to enter value " & strValue &  " in " & strFieldName & " Field. Please check if field exists",true)
	End If
End Function

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  29/09/2021
'Name - 				fn_Enter
'description: 			fn_Enter :  Enter Value in Oracle Form Field
'Parameter			       objParent,strValue
'Return Type -null           
'*************************************************************************
'=============================================================
Public Function fn_Enter(objParent,strValue)


	If IsNull(strValue) Then
		strValue = ""
	End If
	If objParent.Exist(5) Then
	    	objParent.Enter strValue
	    	fn_Enter = true
		'Call fnReportEvent ("Pass","Entered Value in "& strFieldName & "Field", "Function :: fn_Enter = " & strValue & " exists",false)	    	
	Else
		Call fnReportEvent ("Fail","Enter Value in "& vPropertyValue & "Field", "Function :: fn_Enter = " & strValue & " does not exists, Please check",true)
		fn_Enter = false
	End If
End Function



'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  29/09/2021
'Name - 				fn_EnterField
'description: 			fn_EnterField :  Enter Value in Oracle Form Table
'Parameter			       objParent,intRecordNo,strColumn,strValue,strFieldNames
'Return Type -null           
'*************************************************************************
'=============================================================
Public Function fn_EnterField(objParent,intRecordNo,strColumn,strValue,strFieldName)
	If IsNull(strValue) or len(strValue)=0 Then
		'fnReportEvent "Pass","Function :: fn_EnterField","Entered Blank Value " & strValue &  " in " & strFieldName & " Field",true
		Exit Function
	End If
	If objParent.Exist(10) Then
	    	objParent.EnterField intRecordNo,strColumn,strValue	
		fnReportEvent "Pass","Function :: fn_EnterField","Entered Value " & strValue &  " in " & strFieldName & " Field",false	    	
	Else
		fnReportEvent "Fail","Function :: fn_EnterField","Unable to enter value " & strValue &  " in " & strFieldName & " Field. Please check if field exists",true	
	End If
End Function


'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  29/09/2021
'Name - 				fn_WSSendKeys
'description: 			fn_WSSendKeys :  For using keywordevent
'Parameter			       objParent,strValue,strFieldName
'Return Type -null           
'*************************************************************************
'=============================================================

Function fn_WSSendKeys(strkeywordevent)
	Set mySendKeys = CreateObject("WScript.shell")
	Select Case Ucase(strkeywordevent)
		Case "TAB"
		   mySendKeys.SendKeys("{TAB}")
		Case "ENTER"
		   mySendKeys.SendKeys("{ENTER}")
		  Case "TAB3"
		   mySendKeys.SendKeys("{TAB 3}")
	End Select
	Set mySendKeys=nothing
End Function



'************************************************************************************************
'Function Name:- fn_RandomNumber
'Description:- Function to Create a Random Number of Any Length
'Input Parameters:- LengthOfRandomNumber
'Output Parameters:- 
'Created By:- 	MMC team
'************************************************************************************************
 Function fn_RandomNumber(LengthOfRandomNumber)

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
 fn_RandomNumber = iTmp * (10 ^(iLength - iLen))

 End Function
 
'Function f_GetNumericValueFromString(str)
'					Dim c,a
'					for x=1 to len(str)
'							c=mid(str,x,1)
'							If isnumeric(c) OR c="," then
'							     a=a&c
'							End If 
'					next
'					f_GetNumericValueFromString = a
'End Function

Public Function fn_GetROPropertyValueByPropName(objParent,propName)

	If objParent.Exist(1) Then
		fn_GetROPropertyValueByPropName = objParent.getroproperty(propName)	
	Else
		fn_GetROPropertyValueByPropName=""	
	End If

End Function



'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  02/11/2021
'Name -                 fn_getSysdateFormat 
'description:         fn_getSysdateFormat : Gets the system Date in respective format 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_getSysdateFormat(pDateFormat)
v_date=Day(Date())
v_MonthName=Month(Date())
v_year=Year(Date())

Select Case pDateFormat

Case "DD-MMM-YYYY"
v_MonthName=MonthName(Month(Date()),true)
fn_getSysdateFormat=v_date&"-"&v_MonthName&"-"&v_year

Case "MMDDYY"
fn_getSysdateFormat=v_date&v_MonthName&v_year
End Select

End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  02/11/2021
'Name -                 fn_FileExist 
'description:         fn_FileExist : Check the file existence 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_FileExist(fileLocation)
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Wait(2)
    If objFso.FileExists(fileLocation) Then
'        fnReportEvent "Pass","Download File Location","Succefully Downloaded the file at:" &fileLocation ,false
        fn_FileExist = true
    Else 
       fnReportEvent "Fail","Download File Location","Failed to Download the file at:" &fileLocation ,false
	fn_FileExist = false	           
    End If
End Function
