Public TstExecStart
TstExecStart=Now()
'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_Initialization 
'description: 			fn_Initialization :  will fetch the test data corresponding  to test case and data will be stored in the dictionary object
'Parameter				
'Function call ::		       fn_createDataDictionary
'Return Type -null           gb_TestDataDic
'*************************************************************************
'=============================================================

Function fn_Initialization()
	print "Intialization function "
	On error resume Next

	strFileName =  mid(environment("TestDir"),1,InStrRev(environment("TestDir"),"\")) & "TestData\" &  environment("TestDataFileName")  & ".xls"
	print  "Test data  file name : " & strFileName
	vsheetname = environment("TestDataSheetName")
	 vtcIdentifier = gstrTdIdentifer1  &"|" & Split(environment("Legal_entity"),"|")(0)
	
	'creating global dictionary object 
	
	Set gb_TestDataDic = CreateObject("Scripting.Dictionary")	
	Set gb_TestDataDic =  fn_createDataDictionary(strFileName,vsheetname,vtcIdentifier)
'	In this condition will not run the test case
	if ucase(gb_TestDataDic("Run_Flag")) =  UCase("N") then
		blnEndTestcase =true	
		Exit Function
	End  If 

	'''''''Report TC Node Creation	
	call fn_ExtentReportCreTCNode( gstrTestCaseExec_id & ":"  & gstrTestcasename)
	call fnReportEvent("DONE"," ******TC Started :" & gstrTestcasename &" *************", "Test case Name :" & gstrTestCaseExec_id & ":"  & gstrTestcasename,false)
		
	If (gb_TestDataDic("TC_ID")="" or len(gb_TestDataDic("TC_ID")) = 0  or Isempty(gb_TestDataDic("TC_ID"))=true )Then
			call fnReportEvent("Fail","Intialization_fn", "Fail to  fetch the test data from the test data file for test case :" & gstrTestcasename, false)	
			fn_Initialization = false
		 	SystemUtil.CloseProcessByName  "iexplore.exe"

	else
		call fnReportEvent("Pass","Intialization_fn", "Successfully fetch the test data from the file for test case :"& gstrTestcasename ,false)	
		fn_Initialization = true
'		strQuery="Insert into [ExecutionResult$] Values ('"&gstrTestCaseExec_id&"','"&TstExecStart&"','"&gb_TestDataDic("ModuleName")&"','"&gb_TestDataDic("Legal_entity")&"','','','','','')"
'	strQuery="Insert into [ExecutionResult$] Values ('"&gstrTestCaseExec_id&"','"&TstExecStart&"','"&gb_TestDataDic("ModuleName")&"','"&gb_TestDataDic("Legal_entity")&"','','','','','')"
'		print strQuery
'		call fn_updateQuery(strQuery)
	End If
	
	
	If err.number <> 0 Then
		call fnReportEvent("Fail", "fn_Initialization", "Fail to fetech the test data from the file",false)
		fn_Initialization = false
		 SystemUtil.CloseProcessByName  "iexplore.exe"
	End If

End Function

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_BrowserSelect 
'description: 			fn_BrowserSelect :  will Select the browser based on the ini file
'Parameter				
'Function call ::		       
'Return Type -null :          
'*************************************************************************
'=============================================================

Function fn_BrowserSelect(strBrowserName,strURL)
On Error resume Next

Select Case Ucase(strBrowserName)

              Case "IE"
                             SystemUtil.CloseProcessByName  "iexplore.exe"
                             SystemUtil.Run "C:\Program Files (x86)\Internet Explorer\iexplore.exe" , strUrl,,,3
'                         
        Case "CHROME" 
                		SystemUtil.CloseProcessByName  "chrome.exe"
                            SystemUtil.Run "C:\PFilesrogram  (x86)\Google\Chrome\Application\chrome.exe" , strUrl,,,3
                            
               Case "EDGE"
                           SystemUtil.CloseProcessByName  "msedge.exe"
                           SystemUtil.Run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" , strUrl,,,3
                           

'              Case "Firefox"
'                            SystemUtil.CloseProcessByName  "firefox.exe"
'                             SystemUtil.Run "firefox.exe" , strUrl
'                              hWnd = Browser("temp_Browsername").GetROProperty("hwnd")
'                             Window("hwnd:=" & hWnd).Maximize
        
              Case Else
                             msgbox "Please check the Browser Name."
               
End Select 

                 If Err.number <> 0 Then
                                 Print Err.description
                                 ExitTest
                 End If
                   
End Function
'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_LoginSSO 
'description: 			fn_LoginSSO :  will login into Oracle EBiz application 
'Parameter				
'Function call ::		       
'Return Type -null          true or false
'*************************************************************************
'=============================================================

Public Function fn_LoginSSO()

On Error Resume Next

strUrl = 	environment("URL_SSO" )									'"https://test.risebs.mmc.com/OA_HTML/AppsLogin"
strUsername =environment("SSO_Username")
strPassword = environment("SSO_Password")
Call fn_BrowserSelect(environment("Browser_name"),strUrl)
                    
                Set   brSSOLoginObject = Browser("name:=Secure Login - Marsh & McLennan Companies.*").Page("title:=Secure Login - Marsh & McLennan Companies.*")

                If brSSOLoginObject.Exist(45) Then
                    fnReportEvent "Pass", "Login Page Status","Login Page Loaded Successfully",true
                Else
                   	 fnReportEvent "Fail", "Login Page Status","Login Page Not Found, check URL/Page properties again. Exiting Test",true        
                      ExitTest                                    
                End If
                
                brSSOLoginObject.WebEdit("xpath:=//INPUT[@id='username']").Set strUsername
                brSSOLoginObject.WebEdit("xpath:=//INPUT[@id='password']").Set strPassword

                If  brSSOLoginObject.WebButton("xpath:=//INPUT[@value='Login']").Exist(10) Then
                    brSSOLoginObject.WebButton("xpath:=//INPUT[@value='Login']").Click                    
                End If
                
            Set    MMCPageObj =  Browser("title:=MMC CIS Portal.*").Page("title:=MMC CIS Portal.*")
            If MMCPageObj.Exist(45)Then
            	  fn_LoginSSO =true            
                fnReportEvent "Pass", "SSO Login - Home Page Status","Expected Oracle EBS Home Page is loaded successfully creds are" & strUsername ,true
            Else
                fnReportEvent "Fail", "SSO Login - Home Page Status","Failed to Load Oracle EBS Home Page.Using Test Creds value is = " &strUsername,true
		  fn_LoginSSO =   false       			  
               ExitTest
            End If  

          
          If err.number <> 0  Then
          	  fnReportEvent "Fail", "SSO Login - Home Page Status",err.description,true
          	  fn_LoginSSO =   false
          	  ExitTest
          End If
End Function
'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_Login_Iproc 
'description: 			fn_Login_Iproc :  will login into Oracle iproc and iexpense  application 
'Parameter				
'Function call ::		       
'Return Type -null          true or false
'*************************************************************************
'=============================================================

Function fn_Login_Iproc()
    
    On Error Resume Next
    
    blnResultflag=false
    strUrl = 	environment("URL_iproc")									'"https://test.risebs.mmc.com/OA_HTML/AppsLogin"
    strUsername =environment(gb_TestDataDic("Legal_entity") & "_" &"Username")
    strPassword = environment(gb_TestDataDic("Legal_entity") & "_" & "Password")
    
    Call  fn_BrowserSelect(environment("Browser_name"),strUrl)
            
            Set userdefobj = Browser("name:=Login").Page("title:=Login")
            If userdefobj.Exist(45) Then
'                    fnReportEvent "Pass", "Login Page Status","Login Page Loaded Successfully",False
            Else
                    fnReportEvent "Fail", "Login Page Status","Login Page Not Found, check URL/Page properties again. Exiting Test",true        
                    ExitTest    
            End If
    
    
        userdefobj.WebEdit("xpath:=//INPUT[@id='unamebean']").Set strUsername
        userdefobj.WebEdit("xpath:=//INPUT[@id='pwdbean']").Set strPassword        
        
        If  userdefobj.WebButton("xpath:=//*[@id='SubmitButton']").Exist(2) Then 
            userdefobj.WebButton("xpath:=//*[@id='SubmitButton']").Click
            
                Set    OracleAppPageObj =  Browser("name:=Oracle Applications Home Page").Page("title:=Oracle Applications Home Page")
                                    OracleAppPageObj.Sync
                                    If OracleAppPageObj.Exist(20)Then
                                        fnReportEvent "Pass", "Oracle Application Page Status","Successfully login to application and user id = " &strUsername,false
                                        blnResultflag=true
                                    Else
                                        fnReportEvent "Fail","Oracle Application Page Status","Expected Page is not loaded successfully",true
                                        ExitTest
                                    End If            
        End If
            fn_Login_Iproc=blnResultflag
    
End Function


'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_NavigateResponsibility 
'description: 			fn_NavigateResponsibility : NavigateResponsibility page:  will click on the responsibility based on the test data
'Parameter				
'Function call ::		       
'Return Type -null          true or false
'*************************************************************************
'=============================================================

Function fn_NavigateResponsibility()
    blnResultFlag=false
    counter = 0
    On error resume next
    If gstrTdIdentifer2<>"" Then
    	arrResponsibility = Split(gb_TestDataDic(gstrTdIdentifer2),"|")
    	
	    	For iRespIndex = 0 To ubound(arrResponsibility) 
		   	Set OracleAppObj = Browser("name:=Oracle.*").Page("title:=Oracle.*")			
			    If  Not fn_Click(OracleAppObj.Link("text:="&arrResponsibility(iRespIndex),"index:=0")) Then    			           				                         
			    		counter = counter +1
			    End If			    			   
	  	Next
	
		If counter = 0 Then
			blnResultFlag =true
			if ubound(arrResponsibility) = 1 then 
				fnReportEvent "Pass", "Oracle  Navigation responsibility Status","Responsibility Found.Successfully clicked on Responsibility name ="&arrResponsibility(0) & "----->" & arrResponsibility(1)   ,false 	
			else
			 	fnReportEvent "Pass", "Oracle  Navigation responsibility Status","Responsibility Found.Successfully clicked on Responsibility name ="&arrResponsibility(0)   ,false 	
			End If
		else
			fnReportEvent "Fail","Oracle Navigation responsibility Status","Responsibility details is not present corresponding to the user= " & arrResponsibility(0) & "----->" & arrResponsibility(1),true
			Exit Function
 		End If 		  
	Else 
		 fnReportEvent "Fail","Oracle Navigation responsibility Status","Responsibility details is not present in the test data file" ,true
	End If 
	

     	If Err.number <> 0 Then             
              fnReportEvent "Fail","Oracle IProc Navigation Status","Fail to Navigate the responsibilty " & Err.description,true
             fn_NavigateResponsibility = false
             Exit function
      End If
                    
       fn_NavigateResponsibility=blnResultFlag             
End Function


'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  01/09/2021
'Name - 				fn_NavigateMenu 
'description: 			fn_NavigateMenu : will navigate to the Oracle application from EBiz  Page 
'Parameter				
'Function call ::		       
'Return Type -null          true or false
'*************************************************************************
'=============================================================

Function fn_NavigateMenu()
    On error resume next
                 Set MMCPageObj_1 = Browser("name:=MMC CIS Portal").Page("title:=MMC CIS Portal")
                 Set OraclePageObj_1 = Browser("name:=Oracle.*").Page("title:=Oracle.*").WebElement("innertext:=Oracle.*","html tag:=H1")
                      'vstrselectapp = gb_TestDataDic.item("SelectApplication")
                      MMCPageObj_1.Sync
                                   If MMCPageObj_1.Exist(10)Then
                                        fnReportEvent "Pass", "Home Page Status","Expected Home Page is loaded successfully",false
                                      MMCPageObj_1.Link("text:=E-Business Suite").Highlight
                                       MMCPageObj_1.Link("text:=E-Business Suite").Click
                                    'MMCPageObj_1.WebElement("innertext:="&vstrselectapp&"").Click
                                    'MMCPageObj_1.WebElement("innertext:=MMC Oracle E-Business Suite (EBS) - OLTT81").Highlight
                                    'MMCPageObj_1.WebElement("innertext:=MMC Oracle E-Business Suite (EBS) - OLTT81").Click
                                    Browser("MMC CIS Portal").Page("MMC CIS Portal").WebElement("MMC Oracle E-Business").Highlight
                                    Browser("MMC CIS Portal").Page("MMC CIS Portal").WebElement("MMC Oracle E-Business").Click				 
                          
						If OraclePageObj_1.Exist(15) Then	
							OraclePageObj_1.Highlight
							 fnReportEvent "Pass", "Oracle Application Page Status","Expected Oracle Application Page is loaded successfully",false
							 fn_NavigateMenu=true
						Else 
							fnReportEvent "Fail", "Oracle Application Page Status","Expected Oracle Application Page is not loaded successfully",false
							fn_NavigateMenu=false
						End If
                                   Else
                                        fnReportEvent "Fail", "Home Page Status","Expected Home Page is not loaded",true
                                        fn_NavigateMenu = false
                                   End If    
                          
	Browser("creationtime:=0").close

		     If Err.number <> 0 Then
		             Print Err.description
		             fn_NavigateMenu=false
		             Exit Function            
		       End If
                    
End Function

'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  29/09/2021
'Name - 				fn_logout
'description: 			fn_logout :It  will  close all the browser
'Parameter			       objParent
'Return Type -null           
'*************************************************************************
'=============================================================

Function fn_logout()
	print "logout function "
	 SystemUtil.CloseProcessByName  "iexplore.exe"
	call fnReportEvent("Pass","logout Step", "Succesfully logout from the application",false)

	Set gb_TestDataDic = Null
	fn_logout = true
	
End Function
 
'=============================================================
'*************************************************************************
'Created By - 	MMC team	
'Creation Time & Date:  29/09/2021
'Name - 				fn_CloseWindow
'description: 			fn_CloseWindow : Close Oracle forms
'Parameter			       objParent
'Return Type -null           
'*************************************************************************
'=============================================================

Public Function fn_CloseWindow(objParent)
	If objParent.Exist(10) Then
	    	objParent.CloseWindow
	    	wait 2 
		Call fnReportEvent ("Pass","Close Window Status","Closed Window ",false)	    	
	Else
		Call fnReportEvent ("Fail","Close Window Status","Unable to Close Window ",true)
	End If
	
End Function


