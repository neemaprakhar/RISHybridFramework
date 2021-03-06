Public objEntrySupplierLink,objSupplierPage,objManageSitePage,objUpdateInfo,objAddressBook
Public ObjCreateSupplier,ObjSelectValues,ObjCreateAddress,ObjSelectSearch,ObjCreateAddSiteCreation,ObjBankAccount,ObjCreateBankAccount,ObjWebTable, mySendKeys,ObjSelectCheckBox,ObjWebTable_OU

Set objEntrySupplierLink = Browser("name:=Oracle Applications Home Page").Page("title:=Oracle Applications Home Page")
Set objSupplierPage = Browser("name:=Suppliers").Page("title:=Suppliers")
Set objManageSitePage = Browser("name:=Manage Sites").Page("title:=Manage Sites")
Set objUpdateInfo = Browser("name:=Update.*").Page("title:=Update.*")
Set objAddressBook = Browser("name:=Address Book").Page("title:=Address Book")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Gauravis Objects for vendor creation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Set ObjSupplierPage = Browser("name:=Suppliers").Page("title:=Suppliers")
Set ObjCreateSupplier = Browser("name:=Create Supplier").Page("title:=Create Supplier")
'Set ObjUpdateInfo = Browser("name:=Update.*").Page("title:=Update.*")
Set ObjSelectValues = Browser("name:=Search and Select List of Values").Page("title:=Search and Select List of Values").Frame("title:=Search and Select List of Values")
'Set ObjAddressBook = Browser("name:=Address Book").Page("title:=Address Book")
Set ObjCreateAddress = Browser("name:=Create/Update Address").Page("title:=Create/Update Address")
Set ObjSelectSearch = Browser("name:=Search and Select List of Values").Page("title:=Search and Select List of Values")
Set ObjCreateAddSiteCreation = Browser("name:=Create Address: Site Creation").Page("title:=Create Address: Site Creation")
Set ObjBankAccount = Browser("name:=Bank Accounts").Page("title:=Bank Accounts")
Set ObjCreateBankAccount = Browser("name:=Create Bank Account").Page("title:=Create Bank Account")
Set ObjWebTable_OU = Browser("name:=Create Address: Site Creation").Page("title:=Create Address: Site Creation").Webtable("html id:=allSitesExtnd.tableRN-nb")
Set ObjWebTable_OU1 = Browser("name:=Create Address: Site Creation").Page("title:=Create Address: Site Creation").WebTable("xpath:=//SPAN[@id='allSitesExtnd.tableRN']/TABLE[1]/TBODY[1]/TR[3]/TD[1]/TABLE[1]")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const searchSuppName_xpath = "xpath:=//input[@id='SearchSuppName']"
Const entrySupplierLink_xpath = "xpath:=//a[@id='N55']"
Const clickGoButton_xpath = "xpath:=//button[@id='GoButton']"
Const supplierInactiveDate_xpath = "xpath:=//input[@id='InactiveOn']"
Const btnSave_xpath = "xpath:=//button[@id='btnSave']"
Const manageSite_xpath = "xpath:=//a[@id='N32:mngSites:0']//img"
Const manageSiteInactiveDate_xpath = "xpath:=//input[@id='N11:EditDateEnabled:0']"
Const manageSiteSaveButton_xpath = "xpath:=//button[@id='applyBtn']"
Const addressBookLink_xpath = "xpath:=//a[@id='POS_HT_SP_B_ADDR_BK']"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Gauravis constant for vendor creation
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Const CreateSupplierButton_xpath = "xpath:=//button[@id='supCreatBtn']"
Const OrgName_xpath = "xpath:=//input[@id='organization_name']"
Const CreateSupplierApply_xpath = "xpath:=//button[@id='applyBtn_uixr']"
'Const AddressBookLink_xpath = "xpath:=//a[@id='POS_HT_SP_B_ADDR_BK']"
Const CreateAddressButton_xpath = "xpath:=//button[@id='Create']"
Const SearchCountry_xpath = "xpath:=//*[@id='HzFlexCountry__xc_0']/a/img"
Const CategoryChoice_xpath = "xpath:=//*[@id='categoryChoice']"
Const CountryCode_xpath = "xpath:=//*[@title='Search Term']"
Const VCGoBtn_xpath = "xpath:=//button[text()='Go']"
Const QuickSelectBtn_xpath = "xpath:=//img[@title='Quick Select']"
Const RdnBtn_xpath = "xpath:=//*[@id='N1:N8:0']"
Const SelectBtn_xpath = "xpath:=(//button[text()='Select'])[1]"
Const AddressLineOne_xpath = "xpath:=//input[@id='HzAddressStyleFlex1']"
Const City_xpath = "xpath:=//input[@id='HzAddressStyleFlex5']"
Const AddressName_xpath = "xpath:=//input[@id='hzPartySiteName']"
Const County_xpath="xpath:=//input[@id='HzAddressStyleFlex6']"
Const State_xpath = "xpath:=//select[@id='HzAddressStyleFlex7']"
Const PostalCode_xpath = "xpath:=//input[@id='HzAddressStyleFlex8']"
Const PurchasingSiteCheckBox_xpath = "xpath:=//input[@id='purSite']"
Const PaymentCheckBox_xpath = "xpath:=//input[@id='paySite']"
Const ContinueBtn_xpath ="xpath:=//button[@id='nextBtn']"
Const VCApplyBtn_xpath = "xpath:=//button[@id='applyBtn']"
Const BankingDetailsLink_xpath = "xpath:=//*[@id='POS_SBD_BUYER_MAIN']"
Const CreateBankingDetailsButton_xpath = "xpath:=//button[@id='CreateBankAccount111']" 
Const BankName_xpath = "xpath:=//input[@id='BankNameSelect']"
Const BranchName_xpath = "xpath:=//input[@id='BranchNameSelect']"
Const ApplyBankingDetailsButton_xpath = "xpath:=//button[@id='Apply']"
Const AccountNumber_xpath = "xpath:=//input[@id='AcctNumber']"
Const SaveButton_xpath = "xpath:=//button[@id='apply']"
Const SelectOUCheckBox_xpath = "xpath:=//*[@id='allSitesExtnd.tableRN']/table/tbody/tr[3]/td/table/tbody/tr[2]/td[1]/input"
Const Country_xpath = "xpath:=//input[@id='HzFlexCountry']"
Const Confirmation_xpath = "xpath:=//*[@id='FwkErrorBeanId']/tbody/tr/td/table/tbody/tr[2]/td[2]/div[1]/div/table/tbody/tr/td[3]/table/tbody/tr/td/h1"
 Const Op_UnitLink_xpath = "xpath:=//*[text()='Operating Unit']"
 Const BankConfirmation_xpath = "xpath:=//TABLE[@id='FwkErrorBeanId']/TBODY[1]/TR[1]/TD[1]/TABLE[1]/TBODY[1]/TR[2]/TD[2]/DIV[1]/DIV[1]/TABLE[1]/TBODY[1]/TR[1]/TD[3]/TABLE[1]/TBODY[1]/TR[1]/TD[1]/H1[1]"


'=============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  24/11/2021
'Name -                 fn_vendorValidation 
'description:         GSI.P2P.AP.SA.025 : to validate the vendor created
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================

Function fn_vendorValidation()
    blnResultflag=false
    On Error Resume Next    
    
  strSupplierName = fn_getExecutionResultData("GSI.P2P.AP.SA.024","Supplier_Name")
  
  If len (strSupplierName)= 0 or isNull(strSupplierName) Then
  	fnReportEvent "Fail","Supplier Name status","Supplier name is not created in TC GSI.P2P.AP.SA.024" ,false
  	Exit function
  End If
 	   
	If ObjSupplierPage.WebEdit(searchSuppName_xpath).Exist(3) Then
		fn_Set ObjSupplierPage.WebEdit(searchSuppName_xpath), strSupplierName				
		fn_Click_fieldname ObjSupplierPage.WebButton(clickGoButton_xpath),"Go"  
		fnReportEvent "Pass","Supplier Page Name","Successfully entered Supplier Name and value is ::" & strSupplierName ,False
	Else 
		fnReportEvent "Fail","Supplier Page Name","Not able to enter Supplier Name="  & strSupplierName ,true  
	End If
	
    fn_Set ObjUpdateInfo.WebEdit(supplierInactiveDate_xpath),""
    fn_Click_fieldname ObjUpdateInfo.WebButton(btnSave_xpath),"Save"            
    vstrStatus = ObjUpdateInfo.WebEdit(inactiveDate).GetROProperty("innertext")
	If vstrStatus = "" Then
		fnReportEvent "Pass", "Update Supplier Page Status","Able to update the supplier, Inactive date is removed,saved & clicked on Address book link",false        
	Else 
		fnReportEvent "Fail", "Update Supplier Page Status","Not able to update the supplier , Check the update supplier page",true
	Exit function    
	End If    
	
	fn_Click ObjUpdateInfo.Link(AddressBookLink_xpath)
	fn_Click_fieldname ObjAddressBook.Image(manageSite_xpath),"Manage site"
	fn_Set ObjManageSitePage.WebEdit(manageSiteInactiveDate_xpath),""
        vstrUpdatedStatus = ObjManageSitePage.WebEdit(manageSiteInactiveDate).GetROProperty("innertext")
        
		If vstrUpdatedStatus = "" Then
			fnReportEvent "Pass", "Manage Sites Page Status","Able to remove the Inactive date & save",false
			blnResultflag=true        
		Else 
			fnReportEvent "Fail", "Manage Sites Page Status","Unable to remove the Inactive date & save",true
			Exit function    
		End If      
		
	fn_Click_fieldname ObjManageSitePage.WebButton(manageSiteSaveButton_xpath),"Save"
	Browser("name:=Address Book").Close
	fn_vendorValidation = blnResultflag
	
            If Err.number <> 0 Then             
                fnReportEvent "Fail","Vendor Validation","Failed to Validate Vendor " & Err.description,false
              Exit function
              End If    
    End Function
    
    '===============================================================
'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  29/09/2021
'Name -                 fn_VendorCreation 
'description:             fn_VendorCreation :  will create supplier with Address & Banking Details
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_VendorCreation()               'Responsibility used - GLB Supplier Maintenance

On error resume next
blnresult=false
ObjSupplierPage.Sync
Organization_Name = "ORG" & fn_RandomNumber(4)
vstrRowCount = ObjWebTable_OU1.GetROProperty("rows")

Call fn_Click_fieldname(ObjSupplierPage.WebButton(CreateSupplierButton_xpath),"Create Supplier")

'Code to enter Organizaton Name    
fn_Set ObjCreateSupplier.WebEdit(OrgName_xpath),Organization_Name
strQuery="UPDATE [ExecutionResult$] SET Supplier_Name='"&Organization_Name&"' where TC_ID='"&gstrTestCaseExec_id&"' and Start_Date='"&TstExecStart&"'"
Call fn_updateQuery(strQuery)

fnReportEvent "Pass","Address Line One","Successfully enter Organization Name and value is ::" & Organization_Names,False    
Call fn_Click_fieldname(ObjCreateSupplier.WebButton(CreateSupplierApply_xpath),"Create Supplier Apply Button")
fn_Click ObjUpdateInfo.Link(AddressBookLink_xpath)       
Call fn_Click_fieldname(ObjAddressBook.WebButton(CreateAddressButton_xpath),"Create Address Button")
Call fn_AddressDetails
'Call fn_SelectOU
blnresult = fn_SelectOU
If blnresult= false Then
    fnReportEvent "Fail","Operating Unit Selection","Unable to Select Operating Unit",false
    Exit Function
End If
    'Code to check and confirm if address is entered successfully         
    If ObjAddressBook.WebElement(Confirmation_xpath).GetROProperty("innertext")="Confirmation" Then        
        'If ObjAddressBook.WebElement("html id:=FwkErrorBeanId").GetROProperty("innertext")="Confirmation" Then
        fnReportEvent "Pass","Address Details Save Status","Confirmation message displayed & Address Details saved successfully",False    
    Else 
        fnReportEvent "Fail","Address Details Save Status","Confirmation message not displayed. Unable to save Address Details",True
    End If
    
'    Call fn_BankingDetails
    blnresult = fn_BankingDetails
       fnReportEvent "Pass","Vendor Creation","Successfully Created Vendor",False
    fn_VendorCreation = blnresult
    If Err.number <> 0 Then             
        fnReportEvent "Fail","Vendor Creation","Failed to Create Vendor " & Err.description,true
        Exit function
    End If
    
End Function

'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  29/09/2021
'Name -                 fn_SelectOU 
'description:             fn_SelectOU :  will select specific Operating Unit from List 
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_SelectOU()
On error Resume Next 
    blnresultflag = false     
     Set ObjCreateAddSiteCreation = Browser("name:=Create Address: Site Creation").Page("title:=Create Address: Site Creation")      
       Set ObjWebTable = Browser("name:=Create Address: Site Creation").Page("title:=Create Address: Site Creation").WebTable("xpath:=//span[@title='Operating Unit']/../../parent::tbody//parent::table[1]")    
      vstrEndRowCount = ObjWebTable.GetROProperty("rows")
      vstrOperatingUnit = gb_TestDataDic.item("Select_Operating_Unit")  'will fetch gb_dic object - US OU
      
      vasciinumber = Asc(letf(vstrOperatingUnit,1) )
      
      if vasciinumber <= 90 then
    print " will do the sorting on the operating unit--> click on the OU object links" 
    fn_Click ObjCreateAddSiteCreation.Link(Op_UnitLink_xpath)
    fn_Click ObjCreateAddSiteCreation.Link(Op_UnitLink_xpath)
    End If 
Do 
    For  RowIndex = 1 To vstrEndRowCount
             vstrCellData = ObjWebTable.GetCellData(RowIndex,3)  
'        print vstrCellData             
                 If trim(vstrCellData)=vstrOperatingUnit then
                     Set ObjSelectCheckBox = ObjWebTable.ChildItem(RowIndex,0,"WebCheckBox",0)
                     ObjSelectCheckBox.Set "ON" 
                     fnReportEvent "Pass"," enabled the checkbox","Successfully enabled the checkbox and select checkbox value is = " & vstrOperatingUnit ,False
            blnresultflag= true                     
                     Exit do
                 End If         
             
             if lastpagecounter then
                   fnReportEvent "Fail"," Unable to select OU","Fail to select OU and  value is =" & vstrOperatingUnit ,true
                  Exit do
              End if
        Next
        
  Set ObjWebTable = Browser("name:=Create Address: Site Creation").Page("title:=Create Address: Site Creation").WebTable("xpath:=//span[@title='Operating Unit']/../../parent::tbody//parent::table[1]")
'  ObjCreateAddSiteCreation.Image("alt:=Select to view next set","index:=0").Click
     If  ObjCreateAddSiteCreation.Image("alt:=Select to view next set","index:=0").GetROProperty("Visible") = true Then         
           ObjCreateAddSiteCreation.Image("alt:=Select to view next set","index:=0").Click
           nextagecounter = true
    ElseIf  ObjCreateAddSiteCreation.Image("alt:=Next functionality disabled","index:=0").Exist = true Then            
           vstrEndRowCount = ObjWebTable.GetROProperty("rows")
           lastpagecounter =true
    End  If

    
Loop While ( nextagecounter = true or lastpagecounter =true)
Call fn_Click_fieldname(ObjCreateAddSiteCreation.WebButton(VCApplyBtn_xpath),"Apply Button")
fn_SelectOU = blnresultflag

If Err.number <> 0 Then             
         fnReportEvent "Fail","Operating Unit Selection","Failed to Select Operating Unit " & Err.description,true
         Exit function
End If
End Function

'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  29/09/2021
'Name -                 fn_AddressDetails 
'description:             fn_AddressDetails :  will enter Address Details and confirm the data is saved successfully
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_AddressDetails()
On error resume next
fn_AddressDetails = false

'Code to add Address Details 
fn_Click ObjCreateAddress.Image(SearchCountry_xpath)

    If fn_exist(ObjSelectValues.WebList(CategoryChoice_xpath))=true Then
        ' ObjSelectValues.WebList(CategoryChoice_xpath).Highlight
        ObjSelectValues.WebList(CategoryChoice_xpath).Select gb_TestDataDic.item("Country_Choice")          
         fn_Set ObjSelectValues.WebEdit(CountryCode_xpath),gb_TestDataDic.item("Country_Code")                            
        fnReportEvent "Pass","Country Code","Successfully entered Country Code and value is ::" & gb_TestDataDic.item("Country_Code") ,False   
        fn_Click ObjSelectValues.WebButton(VCGoBtn_xpath)
        ObjSelectValues.WebRadioGroup(RdnBtn_xpath).Select "0"                      
        fn_Click ObjSelectValues.WebButton(SelectBtn_xpath)
    Else 
         fnReportEvent "Fail","Country Code","Unable to enter Country Code",False
         Exit Function
    End  If 
    Browser("name:=Create/Update Address").Sync
    If fn_exist(ObjCreateAddress.WebEdit(AddressLineOne_xpath))=true Then
        fn_Set ObjCreateAddress.WebEdit(AddressLineOne_xpath),gb_TestDataDic.item("Address_Line_One")
        fnReportEvent "Pass","Address Line One","Successfully entered Address Line One and value is :" & gb_TestDataDic.item("Address_Line_One"),False
    Else 
        fnReportEvent "Fail","Address Line One","Not able to enter Address Line Ones" & gb_TestDataDic.item("Address_Line_One"),true
        Exit Function
    End If

    fn_Set ObjCreateAddress.WebEdit(City_xpath),gb_TestDataDic.item("City_Name")
    fnReportEvent "Pass","City Name","Successfully entered City Name and value is ::" & gb_TestDataDic.item("City_Name") ,False
    ObjCreateAddress.WebList(State_xpath).Select gb_TestDataDic.item("State_Name")                                
    fn_Set ObjCreateAddress.WebEdit(PostalCode_xpath),gb_TestDataDic.item("Postal_Code")                           'pass postal cofe from test data sheet - Example : 95101
    fnReportEvent "Pass","Postal Code","Successfully entered Postal Code and value is ::" & gb_TestDataDic.item("Postal_Code") ,False    
    fn_Set ObjCreateAddress.WebEdit(AddressName_xpath),gb_TestDataDic.item("Address_Name")                    'pass Address Name from test data sheet - Example : Home 
    fnReportEvent "Pass","Address Name","Successfully entered City Name and value is ::" & gb_TestDataDic.item("Address_Name") ,False
    'ObjCreateAddress.WebEdit(County_xpath).Set vstrCounty                                'pass County Name from test data sheet - Example : Santa Clara
    
    ObjCreateAddress.WebCheckBox(PurchasingSiteCheckBox_xpath).Set "ON"                'select purchasing site checkbox
    ObjCreateAddress.WebCheckBox(PaymentCheckBox_xpath).Set "ON"                    'select payment check box 
    
    Call fn_Click_fieldname(ObjCreateAddress.WebButton(ContinueBtn_xpath),"Continue Button")
    fn_AddressDetails=true

    If Err.Number<>0 Then
        fnReportEvent "Fail","Address Details","Failed to enter Address Details" & Err.description,True
        Exit Function
    End If
End Function

'*************************************************************************
'Created By -     MMC team    
'Creation Time & Date:  29/09/2021
'Name -                 fn_BankingDetails 
'description:             fn_BankingDetails :  will enter Banking Details and confirm the data is saved successfully
'Parameter                
'Function call ::               
'Return Type -null          
'*************************************************************************
'=============================================================
Function fn_BankingDetails()
On error resume next
fn_BankingDetails=false
Account_No = "ACC" & fn_RandomNumber(5)
fn_Click ObjAddressBook.Link(BankingDetailsLink_xpath)

Call fn_Click_fieldname( ObjBankAccount.WebButton(CreateBankingDetailsButton_xpath),"Create Banking Details Button")

    If fn_exist (ObjCreateBankAccount.WebEdit(BankName_xpath))=true Then
        fn_Set  ObjCreateBankAccount.WebEdit(BankName_xpath),gb_TestDataDic.item("Bank_Name")                              
        fnReportEvent "Pass","Bank Name","Successfully entered Bank Name and value is : " & gb_TestDataDic.item("Bank_Name") ,False       
        fn_WSSendKeys TAB2
        fn_Set ObjCreateBankAccount.WebEdit(BranchName_xpath),gb_TestDataDic.item("Branch_Name")                               
        fnReportEvent "Pass","Branch Name","Successfully entered Branch Name and value is : " & gb_TestDataDic.item("Branch_Name") ,False        
        fn_WSSendKeys TAB2
        If fn_exist (ObjCreateBankAccount.WebEdit(AccountNumber_xpath))=true Then
            fn_Set ObjCreateBankAccount.WebEdit(AccountNumber_xpath),Account_No                                    
            fnReportEvent "Pass","Account No","Successfully entered Account No and value is : " & Account_No ,False 
        End If
               
        fn_Click ObjCreateBankAccount.WebButton(ApplyBankingDetailsButton_xpath)
        Call fn_Click_fieldname(ObjBankAccount.WebButton(SaveButton_xpath),"Create Banking Details Button")
        fnReportEvent "Pass","Bank Details","Successfully entered Bank Details",False  
        fn_BankingDetails = true
    Else 
        fnReportEvent "Fail","Bank Name","Unable to enter Bank Details",False  
        Exit Function
    End If

'Confirmation of Bank Details 
    If ObjBankAccount.WebElement(BankConfirmation_xpath).GetROProperty("innertext")="Confirmation" Then
'    If ObjBankAccount.WebElement("html id:=FwkErrorBeanId").GetROProperty("innertext")="Confirmation" Then
        fnReportEvent "Pass","Bank Details Save Status","Confirmation message displayed & Bank Details saved successfully",False
        fn_BankingDetails = true
    Else 
        fnReportEvent "Fail","Bank Details Save Status","Confirmation message not displayed. Unable to save Bank Details",True
    End  If

    If Err.Number<>0 Then
        fnReportEvent "Fail","Bank Details","Failed to enter Bank Details" & Err.description,True
        Exit Function
    End If
End Function

