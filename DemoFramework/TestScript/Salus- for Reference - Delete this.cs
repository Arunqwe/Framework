//using System;
//using System.Collections.Generic;
//using System.Text.RegularExpressions;
//using System.Windows.Input;
//using System.Windows.Forms;
//using System.Drawing;
//using Microsoft.VisualStudio.TestTools.UITesting;
//using Microsoft.VisualStudio.TestTools.UnitTesting;
//using Microsoft.VisualStudio.TestTools.UITest.Extension;
//using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
//using Salus.GenericFunctions;


//namespace Salus..TestScript
//{
//    /// <summary>
//    /// Summary description for CodedUITest1
//    /// </summary>
//    [CodedUITest]
//    public class Salus
//    {
//        // Object declaration
//        public string screenshotName;
//        HTMLControls objHTMLControl = new HTMLControls();
//        UIMaps.Salus_UIMapClasses.Salus_UIMap objSalus = new UIMaps.Salus_UIMapClasses.Salus_UIMap();
//        public Salus()
//        {
//        }

//        // To create an user record and deactivate the same after verification
//        [DataSource("System.Data.Odbc", "Dsn=Excel Files;dbq=C:\\Salus_Solution\\Salus\\TestScript\\Salus.xlsx;defaultdir=.;driverid=1046;maxbuffersize=2048;pagetimeout=5", "UserCreation$", DataAccessMethod.Sequential), TestMethod]
//        public void Salus_UserCreation()
//        {
//            try
//            {
//                BrowserWindow.ClearCache();
//                BrowserWindow.ClearCookies();
//                // Launch the application
//                objSalus.Launch_SalusApp();

//                // Login to application
//                if (objSalus.UISignintoyouraccountIWindow.UISignintoyouraccountDocument.UIUse_another_accountTable.Exists)
//                {
//                    try
//                    {
//                        Mouse.Click(objSalus.UISignintoyouraccountIWindow.UISignintoyouraccountDocument.UIUse_another_accountTable);
//                    }
//                    catch
//                    { }
//                }
//                objHTMLControl.HtmlEdit2Prop("Username", TestContext.DataRow["Username"].ToString());
//                objHTMLControl.HtmlEdit2Prop("Password", TestContext.DataRow["Password"].ToString());
//                objHTMLControl.HtmlButton2PropCustom("SignIn");
//                Playback.Wait(5000);

//                // Login verification
//                if (objSalus.UISignintoyouraccountIWindow.UIDashboardsSalusCliniDocument1.UINavTabLogoTextIdPane.UIMicrosoftDynamicsCRMImage.Exists)
//                    Console.WriteLine("Login Successful");


//                // Navigate to user creation module
//                objHTMLControl.ClickHTMLHyperlink2prop("HomeTab");
//                objHTMLControl.ClickHTMLHyperlink("Salus");
//                objHTMLControl.ClickHTMLHyperlink("Patients");

//                // Click on New button
//                objHTMLControl.HtmlCustom2PropContains("New");
//                Playback.Wait(5000);

//                // Enter Name
//                //Mouse.Click(objSalus.UIContactNewContactIntWindow.UIContactNewContactDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UITitleLabel);
//                objHTMLControl.HtmlLabel2PropContains("Title");
//                objHTMLControl.HtmlListItem2Prop("Mr");
//                objHTMLControl.HtmlLabel2PropContains("FullName");
//                objHTMLControl.HtmlEdit2Prop("FirstNameEdit", TestContext.DataRow["FirstName"].ToString());
//                objHTMLControl.HtmlLabel2PropContains("LastName");
//                objHTMLControl.HtmlEdit2Prop("LastNameEdit", TestContext.DataRow["LastName"].ToString());
//                objHTMLControl.HtmlButton2PropCustom("Done");

//                // Enter Gender
//                objHTMLControl.HtmlLabel2PropContains("Gender");
//                objHTMLControl.HtmlListItem2Prop("Male");

//                // Enter Birthday
//                objHTMLControl.HtmlLabel2PropContains("Birthday");
//                objHTMLControl.HtmlEdit2Prop("BirthdayEdit", TestContext.DataRow["BirthDay"].ToString());

//                // Enter Preferred Language
//                objHTMLControl.HtmlSpan("PreferredLanguage");
//                objHTMLControl.HtmlImage1Prop("PrefLangLookup");
//                Mouse.Hover(objSalus.UIContactsPatientsMicrWindow.UIContactNewContactDocument1.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIAkanAshanti001Custom.UIAkanAshantiCustom);
//                Mouse.MoveScrollWheel(-4);
//                Mouse.Click(objSalus.UIContactsPatientsMicrWindow.UIContactNewContactDocument2.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UILookUpMoreRecordsCustom.UILookUpMoreRecordsCustom1);
//                Playback.Wait(5000);
//                Keyboard.SendKeys(TestContext.DataRow["PreferredLanguage"].ToString());
//                //Mouse.Click(objSalus.UIContactNewContactIntWindow1.UIContactNewContactDocument.UIInlineDialog_IframeFrame.UILookUpRecordDocument.UISearchforrecordsLabel);
//                //objSalus.UIContactsPatientsMicrWindow.UIContactNewContactDocument2.UIInlineDialog_IframeFrame.UILookUpRecordDocument.UISearchforrecordsEdit.Text = TestContext.DataRow["PreferredLanguage"].ToString();
//                Mouse.Click(objSalus.UIContactsPatientsMicrWindow.UIContactNewContactDocument2.UIInlineDialog_IframeFrame.UILookUpRecordDocument.UIStartsearchImage);
//                Mouse.Click(objSalus.UIContactsPatientsMicrWindow.UIContactNewContactDocument2.UIInlineDialog_IframeFrame.UILookUpRecordDocument.UIAddButton);

//                // Save the created user
//                objHTMLControl.HtmlCustom2PropContains("Save");
//                objHTMLControl.HtmlButton1Prop("SaveButton");

//                // Create New Encounter
//                Mouse.Click(objSalus.UIContactTestAutomatioWindow.UIContactTestAutomatioDocument.UITestAutomationPane.UITestAutomationImage);
//                Mouse.Click(objSalus.UIContactTestAutomatioWindow.UIContactTestAutomatioDocument.UIEncountersPatientHyperlink);
//                Mouse.Click(objSalus.UIContactTestAutomatioWindow.UIContactTestAutomatioDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIArea_cap_contact_capFrame.UIHttpssaluscare2crm4dDocument.UIAddNewEncounterAddarCustom.UI_imgsimagestripstranImage);
//                objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIIEdit.Text = TestContext.DataRow["CallbackNumber"].ToString();
//                Mouse.Click(objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIComplaintLabel);
//                objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIIEdit1.Text = TestContext.DataRow["Complaint"].ToString();
//                Mouse.Click(objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UISymptomsStartedLabel);
//                objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIIDateInputEdit.Text = TestContext.DataRow["SymptomsStartedDate"].ToString();
//                Mouse.Click(objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UISaveImage);

//                // Verify Encounter Creation
//                if (objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIENC24Pane.InnerText.Contains("ENC"))
//                {
//                    Console.WriteLine("Encounter Creation Successful");
//                }

//                // Click Next Stage button
//                objHTMLControl.ClickHtmlDivOneProp("NextStage");
//                Playback.Wait(2000);
//                objHTMLControl.ClickHtmlDivOneProp("NextStage");
//                Mouse.Click(objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UICreatePane.UICreateImage);
//                objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UISearchTextEdit.Text = "an";
//                Mouse.Click(objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIDmsscriptsCustom.UIAnxietyCustom);
//                Mouse.Click(objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIActualresultsPane.UIYesButton);

//                // Verify Accept and Decline
//                if (objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIAcceptButton.Exists
//                    && objSalus.UIEncounterNewEncounteWindow1.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIDeclineButton.Exists)
//                {
//                    Console.WriteLine("Assessment Dispositions Verified");
//                    // Close browser instance
//                    objHTMLControl.KillBrowser();
//                }
//                else
//                {
//                    // Close browser instance
//                    objHTMLControl.KillBrowser();
//                    Assert.Fail("Assessment Dispositions Verification Failed");
//                }
//            }
//            catch (Exception e)
//            {
//                screenshotName = objSalus.captureScreen();
//                Console.WriteLine("Screenshot:" + screenshotName);

//                // Close browser instance
//                objHTMLControl.KillBrowser();
//                Console.WriteLine(e.Message.ToString());
//                Console.WriteLine("\n\n\n\n\n\n");
//                Console.WriteLine(e.StackTrace.ToString());
//                Assert.Fail("Exception Occured");
//            }
//        }

//        [DataSource("System.Data.Odbc", "Dsn=Excel Files;dbq=C:\\Salus_Solution\\Salus\\TestScript\\Salus.xlsx;defaultdir=.;driverid=1046;maxbuffersize=2048;pagetimeout=5", "EncounterCreation$", DataAccessMethod.Sequential), TestMethod]
//        public void Salus_EncounterCreation()
//        {
//            try
//            {
//                BrowserWindow.ClearCache();
//                BrowserWindow.ClearCookies();
//                // Launch the application
//                objSalus.Launch_SalusApp();

//                BrowserWindow test = new BrowserWindow();
//                test.Maximized = true;

//                // Login to application
//                if (objSalus.UISignintoyouraccountIWindow.UISignintoyouraccountDocument.UIUse_another_accountTable.Exists)
//                {
//                    try
//                    {
//                        Mouse.Click(objSalus.UISignintoyouraccountIWindow.UISignintoyouraccountDocument.UIUse_another_accountTable);
//                    }
//                    catch
//                    { }
//                }
//                objHTMLControl.HtmlEdit2Prop("Username", TestContext.DataRow["Username"].ToString());
//                objHTMLControl.HtmlEdit2PropPassword("Password", TestContext.DataRow["Password"].ToString());
//                objHTMLControl.HtmlButton2PropCustom("SignIn");
//                Playback.Wait(8000);

//                // Login verification
//                if (objSalus.UISignintoyouraccountIWindow.UIDashboardsSalusCliniDocument1.UINavTabLogoTextIdPane.UIMicrosoftDynamicsCRMImage.Exists)
//                    Console.WriteLine("Login Successful");

//                // Navigate to Encounters module
//                objHTMLControl.ClickHTMLHyperlink2prop("HomeTab");
//                objHTMLControl.ClickHTMLHyperlink("Salus");
//                objHTMLControl.ClickHTMLHyperlink("Encounters");

//                // Verify Encounter Screen
//                if (objSalus.UISignintoyouraccountIWindow.UIDashboardsSalusCliniDocument1.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIActiveEncountersHyperlink.Exists)
//                    Console.WriteLine("Encounter Screen is displayed");

//                // Click on New button
//                objHTMLControl.HtmlCustom2PropContains("New");
//                Playback.Wait(5000);

//                // Verify New Encounter Screen
//                if (objSalus.UIEncounterNewEncounteWindow6.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UINewEncounterPane.Exists)
//                    Console.WriteLine("New Encounter Screen is displayed");

//                //Enter Callback Number
//                objHTMLControl.HtmlEdit2PropSalusContains("CallbackNumber", TestContext.DataRow["CallbackNumber"].ToString());

//                // Enter Patient Name
//                //Mouse.Hover(objSalus.UIEncounterNewEncounteWindow16.UIEncounterNewEncounteDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIPatientclicktoenterPane);
//                objHTMLControl.HoverHtmlDivOneProp("Patient");
//                objHTMLControl.HtmlImage1Prop("PatientLookup");
//                objHTMLControl.HtmlCustom2PropContains("ClickPatient");

//                // Enter Complaint
//                objHTMLControl.ClickHtmlDivOneProp("Complaint");
//                objHTMLControl.HtmlTextAreaTwoProp("ComplaintEdit", TestContext.DataRow["Complaint"].ToString());
//                Keyboard.SendKeys("{TAB}");

//                // Enter Symptoms Duration
//                Keyboard.SendKeys("{TAB}");
//                objHTMLControl.ClickHtmlDivOneProp("SymptomsDuration");
//                objHTMLControl.HtmlEdit2Prop("SymptomsDurationEdit", TestContext.DataRow["SymptomsDuration"].ToString());

//                // Enter Time Unit
//                objHTMLControl.ClickHtmlDivOneProp("SymptomsStratedUnit");
//                objSalus.UIEncounterNewEncounteWindow.UIEncounterNewEncounteDocument1.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UITheunitoftimeforList.UIMinutesListItem.SearchProperties["InnerText"] = TestContext.DataRow["SymptomsStartedUnit"].ToString();
//                Mouse.Click(objSalus.UIEncounterNewEncounteWindow.UIEncounterNewEncounteDocument1.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UITheunitoftimeforList.UIMinutesListItem);

//                // Click Save button
//                objHTMLControl.HtmlCustom2PropContains("Save");
//                Playback.Wait(5000);

//                // Click Next Stage button
//                objHTMLControl.ClickHtmlDivOneProp("NextStage");

//                // Verify Collect Medical History screen
//                if (objSalus.UISignintoyouraccountIWindow.UIDashboardsSalusCliniDocument1.UIContentAreaFrame1.UIHttpssaluscare2crm4dDocument.UIIllnessesAllergiesPane.Exists)
//                    Console.WriteLine("Collect Medical History screen is displayed");

//                // Fill Medical History
//                objHTMLControl.ClickHtmlDivOneProp("OngoingIllness");
//                objHTMLControl.HtmlTextAreaTwoProp("OngoingIllnessEdit", TestContext.DataRow["OngoingIllness"].ToString());
//                Keyboard.SendKeys("{TAB}");
//                objHTMLControl.HtmlTextAreaTwoProp("RegularMedication", TestContext.DataRow["RegularMedication"].ToString());
//                Keyboard.SendKeys("{TAB}");

//                // Add Family History
//                objHTMLControl.HtmlImage1Prop("AddFamilyHistory");
//                Mouse.Click(objSalus.UIFamilyHistoryNewFamiWindow.UIFamilyHistoryNewFamiDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIThefamilialrelationList.UIGrandfatherListItem);
//                Keyboard.SendKeys("{TAB}");
//                Playback.Wait(2000);
//                objHTMLControl.HoverHtmlDivOneProp("FamilyCondition");
//                objHTMLControl.HtmlImage1Prop("FamilyConditionLookup");
//                objHTMLControl.HtmlCustom2PropContains(TestContext.DataRow["FamilyCondition"].ToString());
//                objHTMLControl.HtmlCustom2PropContains("Save & Close");


//                // Add Assessment
//                Mouse.Click(objSalus.UIEncounterENC4InterneWindow.UIEncounterENC4Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UINextStagePane.UINextStagePane1);

//                // Verify Select Assessment Dialogue
//                if (objSalus.UIEncounterENC4InterneWindow.UIEncounterENC4Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UICreatePane.UICreateImage.Exists)
//                    Console.WriteLine("Select Assessment Dialogue is displayed");

//                // Create New Assessment
//                Mouse.Click(objSalus.UIEncounterENC27InternWindow.UIEncounterENC27Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UICreatePane.UICreatePane1);
//                Playback.Wait(5000);

//                // Verify New Assessment
//                if (objSalus.UISignintoyouraccountIWindow.UIDashboardsSalusCliniDocument1.UIContentAreaFrame1.UIHttpssaluscare2crm4dDocument.UINewAssessmentPane.Exists)
//                    Console.WriteLine("New Assessment Screen is displayed");

//                // Enter search script
//                objHTMLControl.HtmlEdit2Prop("AssessmentSearchText", "an");
//                Playback.Wait(2000);
//                objHTMLControl.HtmlCustom2PropContains("ClickAssessment");
//                Playback.Wait(2000);
//                Mouse.MoveScrollWheel(-2);
//                objHTMLControl.HtmlButton1Prop("Yes");

//                // Verify Assessment Close Out Screen
//                if (objSalus.UISignintoyouraccountIWindow.UIDashboardsSalusCliniDocument1.UIAssessmentCloseOutClCustom.UIItemCustom.Exists)
//                    Console.WriteLine("Assessment Close Out Screen is displayed");

//                // Check Care Instructions
//                objHTMLControl.ClickHtmlCheckbox("check-0-0");
//                objHTMLControl.ClickHtmlCheckbox("check-0-1");
//                objHTMLControl.ClickHtmlCheckbox("check-0-2");

//                objSalus.UIEncounterENC4InterneWindow.UIAssessmentAMT3Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIKeepthepatientuprighCheckBox.Checked = true;
//                Mouse.MoveScrollWheel(-3);
//                objSalus.UIEncounterENC4InterneWindow.UIAssessmentAMT3Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UILoosenanytightclothiCheckBox.Checked = true;
//                objSalus.UIEncounterENC4InterneWindow.UIAssessmentAMT3Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIIftheyhavebeenprescrCheckBox.Checked = true;

//                Mouse.MoveScrollWheel(-3);
//                objSalus.UIEncounterENC4InterneWindow.UIAssessmentAMT3Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIKeeptheindividualcalCheckBox.Checked = true;
//                objSalus.UIEncounterENC4InterneWindow.UIAssessmentAMT3Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIIftheirlegsorbackareCheckBox.Checked = true;
//                objSalus.UIEncounterENC4InterneWindow.UIAssessmentAMT3Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UILoosenanyrestrictiveCheckBox.Checked = true;
//                objSalus.UIEncounterENC4InterneWindow.UIAssessmentAMT3Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UICovertheindividualwiCheckBox.Checked = true;
//                objSalus.UIEncounterENC4InterneWindow.UIAssessmentAMT3Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIDonotgiveanythingtoeCheckBox.Checked = true;

//                Mouse.MoveScrollWheel(-2);
//                objSalus.UIEncounterENC4InterneWindow.UIAssessmentAMT3Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIIfsymptomsofshockdevCheckBox.Checked = true;
//                objSalus.UIEncounterENC4InterneWindow.UIAssessmentAMT3Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIIfthereisanychangeinCheckBox.Checked = true;

//                objHTMLControl.HtmlTextAreaTwoProp("AdditionalAdvice", TestContext.DataRow["AdditionalAdvice"].ToString());
//                Mouse.MoveScrollWheel(-2);
//                //objHTMLControl.HtmlButton1Prop("Accept");
//                Mouse.Click(objSalus.UIAssessmentAMT23InterWindow.UIAssessmentAMT23Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIWebResource_DLMFrame.UIHttpssaluscare2crm4dDocument.UIAcceptButton);
//                Playback.Wait(5000);

//                // Close Out
//                objHTMLControl.HtmlCustom2PropContains("Close Out");
//                Playback.Wait(2000);
//                objSalus.UIEncounterENC4InterneWindow.UIHttpssaluscare2crm4dDocument.UIStatusReasonComboBox.SelectedItem = "Complete";
//                Mouse.Click(objSalus.UIEncounterENC4InterneWindow.UIHttpssaluscare2crm4dDocument.UIOKButton);
//                Playback.Wait(2000);
//                Keyboard.SendKeys("{F5}");
//                Playback.Wait(2000);

//                // Verification
//                if (objSalus.UIEncounterENC4InterneWindow.UIAssessmentAMT3Document.UIContentAreaFrame1.UIHttpssaluscare2crm4dDocument.UIStatusReasonCompleteLabel.Exists)
//                {
//                    Console.WriteLine("Encounter Created Successfully");

//                    // Close browser instance
//                    objHTMLControl.KillBrowser();
//                }
//                else
//                {
//                    // Close browser instance
//                    objHTMLControl.KillBrowser();
//                    Assert.Fail("Encounter Creation Failed");
//                }
//            }
//            catch (Exception e)
//            {
//                screenshotName = objSalus.captureScreen();
//                Console.WriteLine("Screenshot: " + screenshotName);

//                // Close browser instance
//                objHTMLControl.KillBrowser();

//                Console.WriteLine(e.Message.ToString());
//                Console.WriteLine("\n\n\n\n\n\n");
//                Console.WriteLine(e.StackTrace.ToString());
//                Assert.Fail("Exception Occured");
//            }
//        }

//        [DataSource("System.Data.Odbc", "Dsn=Excel Files;dbq=C:\\Salus_Solution\\Salus\\TestScript\\Salus.xlsx;defaultdir=.;driverid=1046;maxbuffersize=2048;pagetimeout=5", "SNote$", DataAccessMethod.Sequential), TestMethod]
//        public void Create_Snote()
//        {
//            try
//            {
//                // Launch the application
//                objSalus.Launch_SalusApp();

//                //Login to application
//                if (objSalus.UISignintoyouraccountIWindow.UISignintoyouraccountDocument.UIUse_another_accountTable.Exists)
//                {
//                    Mouse.Click(objSalus.UISignintoyouraccountIWindow.UISignintoyouraccountDocument.UIUse_another_accountTable);
//                }
//                objHTMLControl.HtmlEdit2Prop("Username", TestContext.DataRow["Username"].ToString());
//                objHTMLControl.HtmlEdit2Prop("Password", TestContext.DataRow["Password"].ToString());
//                objHTMLControl.HtmlButton2PropCustom("SignIn");

//                Playback.Wait(5000);

//                // To choose Patients
//                // Navigate to user creation module
//                objSalus.UISignintoyouraccountIWindow.UIDashboardsSalusCliniDocument.UINavTabLogoTextIdPane.UIMicrosoftDynamicsCRMImage.WaitForControlExist(1400);
//                objHTMLControl.ClickHTMLHyperlink2prop("HomeTabLink");
//                objHTMLControl.ClickHTMLHyperlink("Salus");
//                Mouse.Click(objSalus.UISignintoyouraccountIWindow.UIDashboardsSalusCliniDocument1.UIPatientsHyperlink);

//                //To select one patient
//                Mouse.Click(objSalus.UIContactsPatientsMicrWindow.UIContactsPatientsMicrDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIJaneDoeHyperlink);

//                //To select Special note hyperlink
//                Mouse.Click(objSalus.UIContactsPatientsMicrWindow.UIContactsPatientsMicrDocument.UIJaneDoePane.UIJaneDoeImage);
//                Mouse.Click(objSalus.UIContactsPatientsMicrWindow.UIContactsPatientsMicrDocument.UISpecialNotesHyperlink);

//                //To add special note
//                Mouse.Click(objSalus.UIContactsPatientsMicrWindow.UIContactsPatientsMicrDocument.UIContentAreaFrame1.UIHttpssaluscare2crm4dDocument.UIArea_cap_contact_capFrame.UIHttpssaluscare2crm4dDocument.UIAddNewSpecialNoteAddCustom.UIItemCustom);

//                //To enter a Special note
//                Mouse.Click(objSalus.UIContactsPatientsMicrWindow.UISpecialNoteNewSpeciaDocument.UINavBarGloablQuickCreFrame.UIHttpssaluscare2crm4dDocument.UICap_notePane.UINotePane);
//                Keyboard.SendKeys(TestContext.DataRow["SpecialNote"].ToString());

//                //To create a special note click on Save button
//                Mouse.Click(objSalus.UIContactsPatientsMicrWindow.UIContactsPatientsMicrDocument.UISaveButton);

//                //To verify Special note created or not
//                string date = DateTime.Now.ToString("dd/MM/yyyy");
//                Verify("23/11/2016");
//                // Close browser instance
//                objHTMLControl.KillBrowser();
//            }
//            catch
//            {
//                screenshotName = objSalus.captureScreen();
//                Console.WriteLine("Screenshot:" + screenshotName);

//                // Close browser instance
//                objHTMLControl.KillBrowser();

//                Assert.Fail("Exception Occured");
//            }
//        }

//        [DataSource("System.Data.Odbc", "Dsn=Excel Files;dbq=C:\\Salus_Solution\\Salus\\TestScript\\Salus.xlsx;defaultdir=.;driverid=1046;maxbuffersize=2048;pagetimeout=5", "Sheet1$", DataAccessMethod.Sequential), TestMethod]
//        public void Conditions()
//        {
//            try
//            {
//                UIMaps.Salus_Conditions_UIMapClasses.Salus_Conditions_UIMap ObjCon = new UIMaps.Salus_Conditions_UIMapClasses.Salus_Conditions_UIMap();
//                // Launch the application
//                objSalus.Launch_SalusApp();

//                //Login to application
//                if (objSalus.UISignintoyouraccountIWindow.UISignintoyouraccountDocument.UIUse_another_accountTable.Exists)
//                {
//                    Mouse.Click(objSalus.UISignintoyouraccountIWindow.UISignintoyouraccountDocument.UIUse_another_accountTable);
//                }
//                objHTMLControl.HtmlEdit2Prop("Username", "CDNurse@capitahealthcloud.onmicrosoft.com");
//                objHTMLControl.HtmlEdit2Prop("Password", "Wobo7193");
//                objHTMLControl.HtmlButton2PropCustom("SignIn");

//                Playback.Wait(10000);

//                //To choose Patients
//                // Navigate to user creation module
//                objSalus.UISignintoyouraccountIWindow.UIDashboardsSalusCliniDocument.UINavTabLogoTextIdPane.UIMicrosoftDynamicsCRMImage.WaitForControlExist(1400);

//                Playback.Wait(5000);
//                objHTMLControl.ClickHTMLHyperlink2prop("HomeTabLink");
//                objHTMLControl.ClickHTMLHyperlink("Salus");
//                objHTMLControl.ClickHTMLHyperlink("PatientHyperlink");


//                //To select one patient
//                objHTMLControl.ClickHTMLHyperlink2prop("ClickPatient");

//                //To open condition
//                objHTMLControl.HtmlImage2Prop("JaneDoeImg");
//                objHTMLControl.ClickHTMLHyperlink2prop("ConditionHyperLink");

//                //To add new condition
//                objHTMLControl.HtmlCustom2PropContains("AddCondition");
//                if (ObjCon.UIConditionNewConditioWindow.UIConditionNewConditioDocument.Exists)
//                {
//                    Console.Write("New condition page is visible");
//                }

//                //To verify whether the patient name is prefetched or not

//                if (ObjCon.UIConditionNewConditioWindow.UIConditionNewConditioDocument1.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UICap_contactPane.UIPatientJaneDoeJaneDoPane.Exists)
//                {
//                    Console.WriteLine("Patient Name is prefetched");
//                }
//                else
//                {
//                    Console.WriteLine("Patient Name is not prefetched");
//                }

//                //To enter value for condition
//                objHTMLControl.HtmlImage2Prop("ConditionImg");
//                objHTMLControl.HtmlCustom2PropContains("CustomCondition");

//                //To enter a name for condition
//                Mouse.Click(ObjCon.UIConditionNewConditioWindow.UIConditionNewConditioDocument1.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UICap_namePane.UINamePane);
//                Keyboard.SendKeys("Test");

//                //To enter Start Date
//                Mouse.Click(ObjCon.UIConditionNewConditioWindow.UIConditionNewConditioDocument1.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIStartDatePane.UIStartDatePane1);
//                Keyboard.SendKeys("24/11/2016");

//                //To enter End Date
//                Mouse.Click(ObjCon.UIConditionNewConditioWindow.UIConditionNewConditioDocument1.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIEndDatePane.UIEndDatePane1);
//                Keyboard.SendKeys("24/11/2016");

//                //To click on Save button
//                Mouse.Click(ObjCon.UIConditionNewConditioWindow.UIConditionNewConditioDocument1.UISaveSavethisConditioCustom.UIItemCustom);

//                Mouse.Click(ObjCon.UIConditionTestInterneWindow.UIConditionTestDocument1.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UITdAreasPane.UIItemScrollBar);
//                Mouse.StartDragging();
//                Mouse.StopDragging(new Point(1183, 589));
//                //To add treatment
//                objHTMLControl.HtmlImage2Prop("AddTreatmentImg");

//                if (ObjCon.UIConditionTestInterneWindow.UIConditionTestDocument.UISaveCancelPane.Exists)
//                {
//                    Console.WriteLine("New treatment page is opened");
//                }
//                else
//                {
//                    Console.WriteLine("New treatment page is not opened");
//                }

//                //To verify whether patient name and condition is prefetched or not
//                if (ObjCon.UIConditionTestInterneWindow.UITreatmentNewTreatmenDocument.UINavBarGloablQuickCreFrame.UIHttpssaluscare2crm4dDocument.UIJaneDoePane1.Title == "")
//                {
//                    Console.WriteLine("Patient name is not prefetched");
//                }
//                else
//                {
//                    Console.WriteLine("Patient name is prefetched");
//                }
//                if (ObjCon.UIConditionTestInterneWindow.UITreatmentNewTreatmenDocument.UINavBarGloablQuickCreFrame.UIHttpssaluscare2crm4dDocument.UITestPane.Title == "")
//                {
//                    Console.WriteLine("Condition name is not prefetched");
//                }
//                else
//                {
//                    Console.WriteLine("Condition name is prefetched");
//                }

//                //To enter treatment value
//                Mouse.Click(ObjCon.UIConditionTestInterneWindow.UITreatmentNewTreatmenDocument1.UINavBarGloablQuickCreFrame.UIHttpssaluscare2crm4dDocument.UICap_name_iEdit);
//                Keyboard.SendKeys("TestTreatment");

//                //To click on save button
//                objHTMLControl.HtmlButton("Id", "globalquickcreate_save_button_NavBarGloablQuickCreate", "TagName", "Button", "Type", "submit", "InnerText", "Save");

//                //To verify whether Treatment created or not

//                int row = ObjCon.UIConditionTestInterneWindow.UITreatmentNewTreatmenDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIGridBodyTableTable.RowCount;
//                int Flag = 0;
//                for (int i = 1; i < row; i++)
//                {
//                    if (ObjCon.UIConditionTestInterneWindow.UITreatmentNewTreatmenDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIGridBodyTableTable.GetCell(i, 1).FriendlyName.Contains("TestTreatment"))
//                    {
//                        Flag = 1;
//                        break;
//                    }
//                    if (Flag == 1)
//                    {
//                        Console.WriteLine("Treatment Created");
//                    }
//                }
//                if (Flag == 0)
//                {
//                    Console.WriteLine("Treatment not Created");
//                }


//                //To check whether Condition created or not
//                Mouse.Click(ObjCon.UIConditionTestInterneWindow.UITreatmentTestTreatmeTitleBar.UICloseButton);
//                int row1 = ObjCon.UIContactJaneDoeInternWindow.UIContactJaneDoeDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIArea_cap_contact_capFrame.UIHttpssaluscare2crm4dDocument.UIGridBodyTableTable.RowCount;
//                int flag = 0;
//                for (int j = 1; j < row1; j++)
//                {

//                    if (ObjCon.UIContactJaneDoeInternWindow.UIContactJaneDoeDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIArea_cap_contact_capFrame.UIHttpssaluscare2crm4dDocument.UIGridBodyTableTable.GetCell(j, 1).FriendlyName.Contains("Test"))
//                    {
//                        flag = 1;
//                        break;
//                    }
//                    if (flag == 1)
//                    {
//                        Console.WriteLine("Condition Created");
//                    }
//                }
//                if (flag == 0)
//                {
//                    // Close browser instance
//                    objHTMLControl.KillBrowser();

//                    Assert.Fail("Condition not Created");
//                }
//            }
//            catch
//            {
//                screenshotName = objSalus.captureScreen();
//                Console.WriteLine("Screenshot:" + screenshotName);

//                // Close browser instance
//                objHTMLControl.KillBrowser();

//                Assert.Fail("Exception Occured");
//            }
//        }




//        // Additional Functions
//        public void Verify(string d)
//        {
//            bool flag = false;
//            int row = objSalus.UIContactsPatientsMicrWindow.UIContactJaneDoeDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIArea_cap_contact_capFrame.UIHttpssaluscare2crm4dDocument.UIGridBodyTableTable.RowCount;
//            for (int i = 1; i < row; i++)
//            {
//                if (objSalus.UIContactsPatientsMicrWindow.UIContactJaneDoeDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIArea_cap_contact_capFrame.UIHttpssaluscare2crm4dDocument.UIGridBodyTableTable.GetCell(i, 4).FriendlyName.Contains(d) && objSalus.UIContactsPatientsMicrWindow.UIContactJaneDoeDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIArea_cap_contact_capFrame.UIHttpssaluscare2crm4dDocument.UIGridBodyTableTable.GetCell(i, 7).FriendlyName.Contains("Test Note"))
//                {
//                    Mouse.DoubleClick(objSalus.UIContactsPatientsMicrWindow.UIContactJaneDoeDocument.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UIArea_cap_contact_capFrame.UIHttpssaluscare2crm4dDocument.UIGridBodyTableTable.GetCell(i, 1));
//                    if (objSalus.UITreatmentNewTreatmenWindow.UISpecialNoteSN10Document.UIContentAreaFrame.UIHttpssaluscare2crm4dDocument.UICap_notePane.UINoteSpecialTestNoteSPane.InnerText.Contains(TestContext.DataRow["SpecialNote"].ToString()))
//                    {
//                        Console.WriteLine("Special Note Creation Successful");
//                        flag = true;
//                    }
//                    break;
//                }
//            }
//            if (flag == false)
//            {
//                Assert.Fail("Special Note Creation Failed");
//            }
//        }


//        #region Additional test attributes

//        // You can use the following additional attributes as you write your tests:

//        ////Use TestInitialize to run code before running each test 
//        //[TestInitialize()]
//        //public void MyTestInitialize()
//        //{        
//        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
//        //}

//        ////Use TestCleanup to run code after each test has run
//        [TestCleanup()]
//        public void MyTestCleanup()
//        {
//            objSalus.ExcelReport(this.TestContext.TestName, this.TestContext.CurrentTestOutcome.ToString(), TestContext.DataRow.Table.Rows.IndexOf(TestContext.DataRow), screenshotName);
//        }

//        #endregion

//        /// <summary>
//        ///Gets or sets the test context which provides
//        ///information about and functionality for the current test run.
//        ///</summary>
//        public TestContext TestContext
//        {
//            get
//            {
//                return testContextInstance;
//            }
//            set
//            {
//                testContextInstance = value;
//            }
//        }
//        private TestContext testContextInstance;
//    }
//}
