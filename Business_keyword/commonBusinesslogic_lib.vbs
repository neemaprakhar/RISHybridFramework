Public TstExecStart
TstExecStart = Now()

'=============================================================
'*************************************************************************
'Login Page Objects
'=============================================================
'*************************************************************************
Const vIprocLoginUname_xpath = "xpath:=//INPUT[@id='unamebean']"
Const vIprocLoginPasswd_xpath = "xpath:=//INPUT[@id='pwdbean']"
Const vIprocLoginSubmitBtn_xpath = "xpath:=//*[@id='SubmitButton']"
Const vSSOLoginUname_xpath = "xpath:=//INPUT[@id='username']"
Const vSSOLoginPasswd_xpath = "xpath:=//INPUT[@id='password']"
Const vSSOLoginSubmitBtn_xpath = "xpath:=//INPUT[@value='Login']"
'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_Initialization 
'description:             fn_Initialization :  will fetch the test data corresponding  to test case and data will be stored in the dictionary object
'Parameter                
'Function call ::               fn_createDataDictionary
'Return Type -null           gb_TestDataDic
'*************************************************************************
'=============================================================

Function fn_Initialization()
    
    On Error Resume Next
    
    strFileName = Mid(environment("TestDir"),1,InStrRev(environment("TestDir"),"\")) & "TestData\" & environment("TestDataFileName") & ".xls"
'    print  "Test data  file name : " & strFileName
    vsheetname = environment("TestDataSheetName")
    vtcIdentifier = gstrTdIdentifer1 & "|" & Split(environment("Legal_entity"),"|")(0)
    
    'creating global dictionary object 
    
    Set gb_TestDataDic = CreateObject("Scripting.Dictionary")
    Set gb_TestDataDic = fn_createDataDictionary(strFileName,vsheetname,vtcIdentifier)
    '    In this condition will not run the test case
    If UCase(gb_TestDataDic("Run_Flag")) = UCase("N") Then
        blnEndTestcase = True
        Exit Function
    End If
    
    '''''''Report TC Node Creation    
    Call fn_ExtentReportCreTCNode( gstrTestCaseExec_id & ":" & gstrTestcasename)
    Call fnReportEvent("DONE"," ******TC Started :" & gstrTestcasename & " *************", "Test case Name :" & gstrTestCaseExec_id & ":" & gstrTestcasename,False)
    
    If (gb_TestDataDic("TC_ID") = "" Or Len(gb_TestDataDic("TC_ID")) = 0  Or IsEmpty(gb_TestDataDic("TC_ID")) = True )Then
        Call fnReportEvent("Fail","Intialization_fn", "Fail to  fetch the test data from the test data file for test case :" & gstrTestcasename, False)
        fn_Initialization = False
        SystemUtil.CloseProcessByName  "iexplore.exe"
        
    Else
        Call fnReportEvent("Pass","Intialization_fn", "Successfully fetch the test data from the file for test case :" & gstrTestcasename ,False)
        fn_Initialization = True
'        strQuery = "Insert into [ExecutionResult$] Values ('" & gstrTestCaseExec_id & "','" & TstExecStart & "','" & gb_TestDataDic("ModuleName") & "','" & gb_TestDataDic("Legal_entity") & "','','','','','','','','','")"  
	'strQuery = "Insert into [ExecutionResult$] Values ('" & gstrTestCaseExec_id & "','" & TstExecStart & "','" & gb_TestDataDic("ModuleName") & "','" & gb_TestDataDic("Legal_entity") & "','','','','','','','','','','"  & "')"
	strQuery = "Insert into [ExecutionResult$] Values ('" & gstrTestCaseExec_id & "','" & TstExecStart & "','" & gb_TestDataDic("ModuleName") & "','" & gb_TestDataDic("Legal_entity") & "','','','','','','','','','','','','','','"  & "')"         
'        print strQuery
        Call fn_updateQuery(strQuery)
    End If
    
    
    If err.number <> 0 Then
        Call fnReportEvent("Fail", "fn_Initialization", "Fail to fetech the test data from the file",False)
        fn_Initialization = False
        SystemUtil.CloseProcessByName  "iexplore.exe"
    End If
    
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_BrowserSelect 
'description:             fn_BrowserSelect :  will Select the browser based on the ini file
'Parameter                
'Function call ::               
'Return Type -null :          
'*************************************************************************
'=============================================================

Function fn_BrowserSelect(strBrowserName,strURL)
    On Error Resume Next
    
    Select Case UCase(strBrowserName)
        
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
        MsgBox "Please check the Browser Name."
        
    End Select
    
    If Err.number <> 0 Then
        Print Err.description
        ExitTest
    End If
    
End Function
'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_LoginSSO 
'description:             fn_LoginSSO :  will login into Oracle EBiz application 
'Parameter                
'Function call ::               
'Return Type -null          true or false
'*************************************************************************
'=============================================================

Public Function fn_LoginSSO()
    
    On Error Resume Next
    
    strUrl = environment("URL_SSO" )                                    '"https://test.risebs.mmc.com/OA_HTML/AppsLogin"
    strUsername = environment("SSO_Username")
    strPassword = environment("SSO_Password")
    Call fn_BrowserSelect(environment("Browser_name"),strUrl)
    
    Set   brSSOLoginObject = Browser("name:=Secure Login - Marsh & McLennan Companies.*").Page("title:=Secure Login - Marsh & McLennan Companies.*")
    
    If fn_exist(brSSOLoginObject) = True Then
        fnReportEvent "Pass", "Login Page Status","Login Page Loaded Successfully",True
        
        brSSOLoginObject.WebEdit(vSSOLoginUname_xpath).Set strUsername
        fnReportEvent "Pass","Username Field","Successfully entered Username and value is ::" & strUsername,False
        brSSOLoginObject.WebEdit(vSSOLoginPasswd_xpath).Set strPassword
        fnReportEvent "Pass","Password Field","Successfully entered Password and value is ::" & strPassword,False
        
        'fn_Set brSSOLoginObject.WebEdit(vSSOLoginUname_xpath),strUsername,"SSO Login - UserName"
        'fn_Set brSSOLoginObject.WebEdit(vSSOLoginPasswd_xpath),strPassword,"SSO Login - Password"
        
        If  fn_exist(brSSOLoginObject.WebButton(vSSOLoginSubmitBtn_xpath)) = True Then
            fn_Click brSSOLoginObject.WebButton(vSSOLoginSubmitBtn_xpath)
        End If
    Else
        fnReportEvent "Fail", "Login Page Status","Login Page Not Found, check URL/Page properties again. Exiting Test",True
        ExitTest
    End If
    
    'brSSOLoginObject.WebEdit(vSSOLoginUname_xpath).Set strUsername
    'brSSOLoginObject.WebEdit(vSSOLoginPasswd_xpath).Set strPassword
    
    
    Set    MMCPageObj = Browser("title:=MMC CIS Portal.*").Page("title:=MMC CIS Portal.*")
    If MMCPageObj.Exist(45)Then
        fn_LoginSSO = True
        fnReportEvent "Pass", "SSO Login - Home Page Status","Expected Oracle EBS Home Page is loaded successfully creds are" & strUsername ,True
    Else
        fnReportEvent "Fail", "SSO Login - Home Page Status","Failed to Load Oracle EBS Home Page.Using Test Creds value is = " & strUsername,True
        fn_LoginSSO = False
        ExitTest
    End If
    
    If err.number <> 0  Then
        fnReportEvent "Fail", "SSO Login - Home Page Status",err.description,True
        fn_LoginSSO = False
        ExitTest
    End If
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_Login_Iproc 
'description:             fn_Login_Iproc :  will login into Oracle iproc and iexpense  application 
'Parameter                
'Function call ::               
'Return Type -null          true or false
'*************************************************************************
'=============================================================

Function fn_Login_Iproc()
    
    On Error Resume Next
    
    blnResultflag = False
    strUrl = environment("URL_iproc")
    If gstrApproverLogin Then
        strUsername = "ASTRID-SURYAPRANATA"
        strPassword = "oracle02"
        gstrApproverLogin = False
    Else
        '"https://test.risebs.mmc.com/OA_HTML/AppsLogin"
        strUsername = environment(gb_TestDataDic("Legal_entity") & "_" & "Username")
        strPassword = environment(gb_TestDataDic("Legal_entity") & "_" & "Password")
    End If
    Call  fn_BrowserSelect(environment("Browser_name"),strUrl)
    
    Set userdefobj = Browser("name:=Login").Page("title:=Login")
    If fn_exist(userdefobj) = True Then
        fnReportEvent "Pass", "Login Page Status","Login Page Loaded Successfully",False
    Else
        fnReportEvent "Fail", "Login Page Status","Login Page Not Found, check URL/Page properties again. Exiting Test",True
        ExitTest
    End If
    
    userdefobj.WebEdit(vIprocLoginUname_xpath).Set strUsername
    fnReportEvent "Pass","Username Field","Successfully entered Username and value is ::" & strUsername,False
    userdefobj.WebEdit(vIprocLoginPasswd_xpath).Set strPassword
    fnReportEvent "Pass","Password Field","Successfully entered Password and value is ::" & strPassword,False
    
    'fn_Set userdefobj.WebEdit(vIprocLoginUname_xpath),strUsername,"Iproc Login - UserName"
    'fn_Set userdefobj.WebEdit(vIprocLoginPasswd_xpath),strPassword,"Iproc Login - Password"
    
    If  fn_exist(userdefobj.WebButton(vIprocLoginSubmitBtn_xpath)) = True Then
        fn_Click userdefobj.WebButton(vIprocLoginSubmitBtn_xpath)
        
        Set    OracleAppPageObj = Browser("name:=Oracle Applications Home Page").Page("title:=Oracle Applications Home Page")
        OracleAppPageObj.Sync
        If fn_exist(OracleAppPageObj) = True Then
            fnReportEvent "Pass", "Oracle Application Page Status","Successfully login to application and user id = " & strUsername,False
            blnResultflag = True
        Else
            fnReportEvent "Fail","Oracle Application Page Status","Expected Page is not loaded successfully",True
            ExitTest
        End If
    End If
    fn_Login_Iproc = blnResultflag
    
End Function


'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_NavigateResponsibility 
'description:             fn_NavigateResponsibility : NavigateResponsibility page:  will click on the responsibility based on the test data
'Parameter                
'Function call ::               
'Return Type -null          true or false
'*************************************************************************
'=============================================================

Function fn_NavigateResponsibility()
  On Error Resume Next
    blnResultFlag = False
    counter = 0
  
    If gstrTdIdentifer2 <> "" Then
        arrResponsibility = Split(gb_TestDataDic(gstrTdIdentifer2),"|")
        
        For iRespIndex = 0 To UBound(arrResponsibility)
            Set OracleAppObj = Browser("name:=Oracle.*").Page("title:=Oracle.*")
            OracleAppObj.Exist(15)
            If  Not fn_Click(OracleAppObj.Link("text:=" & arrResponsibility(iRespIndex),"index:=0")) Then
                counter = counter + 1
            End If
        Next
        
        If counter = 0 Then
            blnResultFlag = True
            If UBound(arrResponsibility) = 1 Then
                fnReportEvent "Pass", "Oracle  Navigation responsibility Status","Responsibility Found.Successfully clicked on Responsibility name =" & arrResponsibility(0) & "----->" & arrResponsibility(1)   ,False
            Else
                fnReportEvent "Pass", "Oracle  Navigation responsibility Status","Responsibility Found.Successfully clicked on Responsibility name =" & arrResponsibility(0)   ,False
            End If
        Else
            fnReportEvent "Fail","Oracle Navigation responsibility Status","Responsibility details is not present corresponding to the user= " & arrResponsibility(0) & "----->" & arrResponsibility(1),True
            Exit Function
        End If
    Else
        fnReportEvent "Fail","Oracle Navigation responsibility Status","Responsibility details is not present in the test data file" ,True
    End If
    
    
    If Err.number <> 0 Then
        fnReportEvent "Fail","Oracle IProc Navigation Status","Fail to Navigate the responsibilty " & Err.description,True
        fn_NavigateResponsibility = False
        Exit Function
    End If
    
    fn_NavigateResponsibility = blnResultFlag
End Function


'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  01/09/2021
'Name -                 fn_NavigateMenu 
'description:             fn_NavigateMenu : will navigate to the Oracle application from EBiz  Page 
'Parameter                
'Function call ::               
'Return Type -null          true or false
'*************************************************************************
'=============================================================
Function fn_NavigateMenu()
    On Error Resume Next
    Set MMCPageObj_1 = Browser("name:=MMC CIS Portal").Page("title:=MMC CIS Portal")
    Set OraclePageObj_1 = Browser("name:=Oracle.*").Page("title:=Oracle.*").WebElement("innertext:=Oracle.*","html tag:=H1")
    'vstrselectapp = gb_TestDataDic.item("SelectApplication")
    MMCPageObj_1.Sync
    If MMCPageObj_1.Exist(10)Then
        fnReportEvent "Pass", "Home Page Status","Expected Home Page is loaded successfully",False
'        MMCPageObj_1.Link("text:=E-Business Suite").Highlight
		fn_Click MMCPageObj_1.Link("text:=E-Business Suite")
'        MMCPageObj_1.Link("text:=E-Business Suite").Click
     
'        Browser("MMC CIS Portal").Page("MMC CIS Portal").WebElement("MMC Oracle E-Business").Highlight
        Browser("MMC CIS Portal").Page("MMC CIS Portal").WebElement("MMC Oracle E-Business").Click
        
        If OraclePageObj_1.Exist(15) Then
            OraclePageObj_1.Highlight
            fnReportEvent "Pass", "Oracle Application Page Status","Expected Oracle Application Page is loaded successfully",False
            fn_NavigateMenu = True
        Else
            fnReportEvent "Fail", "Oracle Application Page Status","Expected Oracle Application Page is not loaded successfully",False
            fn_NavigateMenu = False
        End If
    Else
        fnReportEvent "Fail", "Home Page Status","Expected Home Page is not loaded",True
        fn_NavigateMenu = False
    End If
    
    Browser("creationtime:=0").close
    
    If Err.number <> 0 Then
        Print Err.description
        fn_NavigateMenu = False
        Exit Function
    End If
    
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  29/09/2021
'Name -                 fn_logout
'description:             fn_logout :It  will  close all the browser
'Parameter                   objParent
'Return Type -null           
'*************************************************************************
'=============================================================

Function fn_logout()
'    print "logout function "
    SystemUtil.CloseProcessByName  "iexplore.exe"
    Call fnReportEvent("Pass","logout Step", "Succesfully logout from the application",False)
    
    Set gb_TestDataDic = Null
    fn_logout = True
    
End Function

'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  29/09/2021
'Name -                 fn_CloseWindow
'description:             fn_CloseWindow : Close Oracle forms
'Parameter                   objParent
'Return Type -null           
'*************************************************************************
'=============================================================

Public Function fn_CloseWindow(objParent)
    If objParent.Exist(10) Then
        objParent.CloseWindow
        wait 2
        'Call fnReportEvent ("Pass","Close Window Status","Closed Window",false)            
    Else
        strValue = fn_GetROPropertyValue(objParent)
        Call fnReportEvent ("Fail","Close Window Status",strValue & "  : Unable to Close Window ",True)
    End If
End Function

Function f_GetNumericValueFromString(str)
    Dim c,a
    For x = 1 To Len(str)
        c = Mid(str,x,1)
        If IsNumeric(c) Or c = "," Then
            a = a & c
        End If
    Next
    f_GetNumericValueFromString = a
End Function

Function f_CheckRequestStatus(strReqId)
    
    OracleFormWindow("Navigator").SelectMenu "View->Requests"
    If strReqId <> "" Then
        OracleFormWindow("title:=Find Requests").OracleRadioGroup("selected index:=3").Select "Specific Requests"
        OracleFormWindow("Find Requests").OracleTextField("Request ID").Enter strReqId
    End If
    OracleFormWindow("Find Requests").OracleButton("Find").Click
    intLoopCnt = 1
    Do
        OracleFormWindow("Requests_2").OracleButton("Refresh Data").Click
        strPhaseStatus = OracleFormWindow("Requests").OracleTable("Table").GetFieldValue(1,4)
        If intLoopCnt = 100 Then
            Exit Do
        End If
        intLoopCnt = intLoopCnt + 1
    Loop While strPhaseStatus <> "Completed"
    f_CheckRequestStatus = strPhaseStatus
End Function
'=========================================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  29/09/2021
'Name -                	fn_getSysdateFormat
'description:             fn_getSysdateFormat : get System date in different formats
'Parameter                   objParent
'Return Type -null           
'*************************************************************************
'=========================================================================
Function fn_getSysdateFormat(pDateFormat)
    v_date = Day(Date())
    v_MonthName = Month(Date())
    v_year = Year(Date())
    Select Case pDateFormat
        Case "DD-MMM-YYYY"
        v_MonthName = MonthName(Month(Date()),True)
        fn_getSysdateFormat = v_date & "-" & v_MonthName & "-" & v_year
        Case "MMDDYY"
        fn_getSysdateFormat = v_date & v_MonthName & v_year
    End Select
End Function
