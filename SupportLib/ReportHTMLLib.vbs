Function fn_CreateLogfolder()
	
On error resume Next
	Set objFso = CreateObject("Scripting.FileSystemObject")
	
	''will create the log folder 
	strLogFolderPath = mid(Environment("TestDir"),1,instrRev(Environment("TestDir"),"\")) & "Log"
	
	If Not(objFso.FolderExists(strLogFolderPath)) Then
		objFso.CreateFolder(strLogFolderPath)
	End If
	'will create ScreenShot folder
	strScreenShotfolder = strLogFolderPath & "\ScreenShot"
	environment.value("vScreenShotfolder") = strScreenShotfolder
	If Not(objFso.FolderExists(strScreenShotfolder)) Then
		objFso.CreateFolder(strScreenShotfolder)
	End If
	
	'will create execution report folder
	strExecutionReportFolder =  strLogFolderPath & "\Execution Reports"
	If Not(objFso.FolderExists(strExecutionReportFolder)) Then
		objFso.CreateFolder(strExecutionReportFolder)
	End If	
	
	'will create datefolder 
	strDateFolder = strExecutionReportFolder &"\" &  replace(Date,"/","")	
	environment.value("vLogDatefolder") =  strDateFolder
	If Not(objFso.FolderExists(strDateFolder)) Then
		objFso.CreateFolder(strDateFolder)
	End If
	
'	will create the extent report 
	strExtentReport = strDateFolder &"\RIS_Extent_Report" 
environment.value(	"vExtentReport") =  strExtentReport
	If Not(objFso.FolderExists(strExtentReport)) Then
		objFso.CreateFolder(strExtentReport)
	End If
	
	'will create the UFT report 
	strUFTReport = strDateFolder &"\RIS_UFT_Report" 
	environment.value("vUFTReport") =  strUFTReport
	If Not(objFso.FolderExists(strUFTReport)) Then
		objFso.CreateFolder(strUFTReport)
	End If
	
	If err.number<>0 Then
			Call fnReportEvent ("FAIL","Log folder creation", "Function ::fn_CreateLogfolder = fail to create log folder and error description is ::"& err.description,false)	
	End If
		
End Function




'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_ExtentReporterHeader 
'description: 			fn_ExtentReporterHeader :It will create the Extent report header and file location where log will be saved
'Parameter				
'Function call ::		
'Return Type -null
'*************************************************************************
'=============================================================
Function fn_ExtentReporterHeader()
On error Resume Next 

temp_logfolder =environment.value("vExtentReport") & "\" & Replace(Replace(Replace(now,"/","")," ",""),":","") &  ".html"
'print  "Log folder =" & temp_logfolder

  Set htmlReporter = CreateObject("UFT_Extent_Reports.HTMLReporter")
  htmlReporter.InitializeReport(temp_logfolder )
  htmlReporter.AddReportName("RIS MMC TEST AUTOMATION REPORT")
  htmlReporter.AddDocumentTitle("Business Acceptance Report")

If err.number<>0 Then
	print err.description
End If
End Function 

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_ExtentReportCreTCNode 
'description: 			fn_ExtentReportCreTCNode :It will create the test case node in the  Extent report
'Parameter				strTestCaseName : will be passed from  fn_ReadTestScenario
'Function call ::		
'Return Type -null
'*************************************************************************
'=============================================================

Function fn_ExtentReportCreTCNode(strTestCaseName)
	
	call  htmlReporter.CreateTest(strTestCaseName)
	call htmlReporter.AssignAuthorToTest(Environment.Value("UserName"))
       call htmlReporter.AssignCategoryToTest("Regression Testing")
End Function


'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_ExtentReportlogStepLevel 
'description: 			fn_ExtentReportlogStepLevel :It will create the step level log in the  Extent report
'Parameter				strStatus, strDescription,vimgPath
'Function call in ::		fnReportEvent
'Return Type -null
'*************************************************************************
'=============================================================
Function fn_ExtentReportlogStepLevel(strStatus, strDescription,vimgPath)
	
	On Error Resume Next
		
	imgflag =false
	If Len(vimgPath) > 0 Then
		imgflag = true
	End If
	
	Select Case UCase(strStatus)
		
		Case "PASS"
		If imgflag Then			
			call htmlReporter.AddPassLog(strDescription & "<br><a href = " &  chr(34) & vimgPath  & chr(34)  &">ScreenShot</a>" )
		else
			call htmlReporter.AddPassLog(strDescription)
		End If
		
		Case "FAIL"
			If imgflag Then			
'				 call htmlReporter.AddFailLog(strDescription,vimgPath)
				call htmlReporter.AddFailLog(strDescription & "<br><a href = " &  chr(34) & vimgPath  & chr(34)  &">ScreenShot</a>")
			else
				 call htmlReporter.AddFailLog(strDescription)
			End If
			
		Case "INFO"
			call htmlReporter.AddInfoLog(strDescription)
			
		Case "ERROR"
			call htmlReporter.AddErrorLog(strDescription)
		
	End Select
	
	If err.number<>0 Then
		print err.description
	End If
End  Function 


'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_ExtentReportlogStepLevel 
'description: 			fn_ExtentReportlogStepLevel :It will create the step level log in the  Extent report
'Parameter				strStatus, strDescription,vimgPath
'Function call in ::		fnReportEvent
'Return Type -null
'*************************************************************************
'=============================================================
	
	

Function fn_ExtentGenerateReport()
 '''Change Theme to dark
    htmlReporter.ChangeToDarkTheme
  
    ''Generate the html reports
	    htmlReporter.GenerateReport()

indexHtml =environment.value("vExtentReport") & "\index.html"
dashboardHtml = environment.value("vExtentReport") & "\dashboard.html"
taghtml = environment.value("vExtentReport") & "\tag.html"

'renaming the report file 
Set objFso = CreateObject("Scripting.FileSystemObject")

If fn_FileExist(indexHtml) Then 
	objFso.MoveFile environment.value("vExtentReport") & "\index.html", environment.value("vExtentReport") & "\index"&Replace(Replace(Replace(now,"/","")," ",""),":","") &  ".html"
End If
If fn_FileExist(dashboardHtml) Then 
	objFso.MoveFile  environment.value("vExtentReport") & "\dashboard.html", environment.value("vExtentReport") & "\dashboard"&Replace(Replace(Replace(now,"/","")," ",""),":","") &  ".html"
End If
	
If fn_FileExist(taghtml) Then 
	objFso.MoveFile environment.value("vExtentReport") & "\tag.html", environment.value("vExtentReport") & "\tag"&Replace(Replace(Replace(now,"/","")," ",""),":","") &  ".html"
End If	
End Function



'******************************************* fnReportEvent()**************************************
'Name                     :   fnReportEvent
'Description          :   To generate the HTML report 
'Created By           :   MMC team 
'Date		               :   13 Aug 2021
'Input Parameters     :	  Yes 	
'Return Value         :   NA 
'******************************************************************************************************

Public Function fnReportEvent(pStatus,pStepName,pStepDesc,ImgFlag)

	if ImgFlag = true then
		vFilePath = environment.value("vScreenShotfolder") &"\"& Replace(Replace(Replace(now,"/","")," ",""),":","") &".png"	
		Desktop.CaptureBitmap vFilePath,True
	End If 	
	Select Case Ucase(pStatus)		
	Case "PASS"
	   Reporter.Filter = 0
		If ImgFlag Then	   	  
			Reporter.ReportEvent micPass,pStepName,pStepDesc,vFilePath	
			call fn_ExtentReportlogStepLevel(pStatus,pStepDesc,vFilePath)
		Else
		  	Reporter.ReportEvent micPass,pStepName,pStepDesc
			call fn_ExtentReportlogStepLevel(pStatus,pStepDesc,"")		  	
		End If
	   Reporter.Filter = 1   
	Case "FAIL"
	   Reporter.Filter = 0
	   	If ImgFlag Then	   	  
			  Reporter.ReportEvent micFail,pStepName,pStepDesc,vFilePath
	  		 call fn_ExtentReportlogStepLevel(pStatus,pStepDesc,vFilePath)
		Else
		  	Reporter.ReportEvent micFail,pStepName,pStepDesc
			call fn_ExtentReportlogStepLevel(pStatus,pStepDesc,"")		  	
		End If
	 
	   Reporter.Filter = 1
	Case "WARNING"
	   Reporter.Filter = 0
		If ImgFlag Then	   	  
			 Reporter.ReportEvent micWarning,pStepName,pStepDesc,vFilePath	 
			 call fn_ExtentReportlogStepLevel(pStatus,pStepDesc,vFilePath)			 
		Else
		  	Reporter.ReportEvent micWarning,pStepName,pStepDesc
			call fn_ExtentReportlogStepLevel(pStatus,pStepDesc,"")			  	
		End If
	   Reporter.Filter = 1
	Case "DONE"
		   Reporter.Filter = 0
		   Reporter.ReportEvent micDone,pStepName,pStepDesc	
'		   call fn_ExtentReportlogStepLevel(pStatus,pStepDesc,"")	
		   Reporter.Filter = 1
	End Select	

		
End Function
	
	

	

