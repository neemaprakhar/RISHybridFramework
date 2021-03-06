
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
print"Record Count" &  objRecSet.RecordCount


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
'Name - 				fn_GetToPropertyValue 
'description: 			fn_GetToPropertyValue :  will return the property value to the function
'Parameter				
'Function call ::		       
'Return Type -null           String =will  return  property value
'*************************************************************************
'=============================================================

Public Function fn_GetToPropertyValue(objParent)

	Reporter.Filter = 3
	If objParent.gettoproperty("innertext") <> "" Then
		fnGetToPropertyValue = objParent.getroproperty("innertext")		
	ElseIf objParent.gettoproperty("name") <> "" Then
		fnGetToPropertyValue = objParent.getroproperty("name")		
	ElseIf objParent.gettoproperty("value") <> "" Then
		fnGetToPropertyValue = objParent.getroproperty("value")	
	ElseIf objParent.gettoproperty("text") <> "" Then
		fnGetToPropertyValue = objParent.getroproperty("text")
	ElseIf objParent.gettoproperty("title") <> "" Then
		fnGetToPropertyValue = objParent.getroproperty("title")
	ElseIf objParent.gettoproperty("html id") <> "" Then
	    fnGetToPropertyValue = objParent.getroproperty("html id")
	Else
		fnGetToPropertyValue = objParent.getroproperty("class")		
	End If

End Function
'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  15/09/2021
'Name - 				fn_Click
'description: 			fn_Click :  click on the object 
'Parameter			       objParent,strFieldName
'Return Type -null           
'*************************************************************************
'=============================================================

Function fn_Click_fieldname(objParent,vFieldName)
	
	If objParent.Exist(10) Then
	   	objParent.highlight	
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
	   	objParent.highlight	
		objParent.Click	
		fn_Click = true
		'fnReportEvent "PASS","Click on a "& vPropertyValue &" Object ",vPropertyValue&" is exist and clicked successfully","NO"		
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
'Parameter				objParent,strValue,strFieldName
'Function call ::		       
'Return Type -           	NA
'*************************************************************************
'=============================================================
Public Function fn_Set(objParent,strValue,strFieldName)

	vPropertyValue = ucase(fn_GetToPropertyValue(objParent))
	vClass =Ucase(objParent.gettoproperty("micclass"))
	If IsNull(strValue) Then
		strValue = ""
	End If
	If objParent.Exist(10) Then
		fn_Highlight objParent
	    	objParent.Set strValue				
'		fnReportEvent "PASS","Set Value in "& vPropertyValue & " object", vPropertyValue& " is exist and expected Value is entered:- " & strValue, true
	Else
		Call fnReportEvent ("FAIL","Set Value in "& strFieldName &" Object", "Function ::fn_Set = " & vPropertyValue& " does not exists, Please check",false)
	End If
'	In futute might need to add the validation code
End Function

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
'		objParent.highlight		
'		fnReportEvent "PASS","Highlight "&vPropertyValue&" Object of class "&vClass,vPropertyValue&" Object of class "&vClass&" Exist and Highlighted Successfully","NO"
	Else
		fnReportEvent "FAIL","Highlight "&vPropertyValue&" Object of class "&vClass,vPropertyValue&" Object of class "&vClass&" Does not Exist, Please Check.","YES"		
	End If

End Function

Public Function fn_updateQuery(strQuery)
Dim objCon
On Error Resume Next	
strFileName =  mid(environment("TestDir"),1,InStrRev(environment("TestDir"),"\")) & "TestData\" &  environment("TestDataFileName")  & ".xls"
Print strFileName
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
'Creation Time & Date:  15/09/2021
'Name - 				fn_Select 
'description: 			fn_Select :  will select the value from the web list 
'Parameter				objParent,strValue,strFieldName(Used for reporting 
'Function call ::		       
'Return Type -           	true or false 
'*************************************************************************
'=============================================================
Public Function fn_Select(objParent,strValue,strFieldName)
	
	vPropertyValue = ucase(fn_GetToPropertyValue(objParent))
	vClass =Ucase(objParent.gettoproperty("micclass"))
	If IsNull(strValue) Then
		strValue = ""
	End If
	If objParent.Exist(10) Then	
	    	objParent.select  strValue				
		fnReportEvent "PASS","Selected Value for the  "& strFieldName & " list", "Succesfully selected the value from the list and value selected is := " & strValue,false
		fn_Select = true
	Else
		Call fnReportEvent ("FAIL","Selected Value for the  " & strFieldName & " list", "Function ::fn_Select = " & strFieldName & " does not exists, Please check",true)
		fn_Select = false		
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

Function fn_exist(objParent)
	counter = 1 
	intflagcounter = 1
	Do
		wait (1)
		counter = counter+1 
		If objParent.Exist(1) Then		    							
			fn_exist = true
			intflagcounter = 0
			Exit do 				
		End If
	Loop Until counter< 15
		
		
		If intflagcounter = 0 Then
			print "object exist"
		Else 
			Print "objectdoesnot exist"
		End If
End Function
