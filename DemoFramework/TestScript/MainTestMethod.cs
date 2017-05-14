using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using RelevantCodes.ExtentReports;
using System.Configuration;
using DemoFramework.Reports;
using DemoFramework.GenericFunctions;

namespace DemoFramework.TestScript
{
    /// <summary>
    /// Summary description for CodedUITest1
    /// </summary>
    [CodedUITest]
    public class MainTestMethod
    {
        Reports.Report objReport = new Reports.Report();
        GenericFunctions.HTMLControls objHTMLControl = new GenericFunctions.HTMLControls();
        GenericFunctions.Excel objExcel = new GenericFunctions.Excel();
        DemoFramework.ObjectRepository.UIMaps.UIMapClasses.UIMap objUIMap = new ObjectRepository.UIMaps.UIMapClasses.UIMap();

        public static ExtentReports extent1;
        public ExtentTest test;
        public static int _maxTestRunsCount = 2;
        public static int _maxTestRuns = _maxTestRunsCount;


        public static int iteration;
        Generic ObjGen = new Generic();

        public MainTestMethod()
        {
        }

        [ClassInitialize]
        public static void Classinitialize(TestContext test)
        {
            string solutionPath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.IndexOf("TestResults"));
            string ReportPath = solutionPath + ConfigurationManager.AppSettings.Get("ExtentPath");

            extent1 = new ExtentReports(ReportPath, true);
            extent1
            .AddSystemInfo("Host Name", "Vinuchandran")
            .AddSystemInfo("Environment", "QA")
            .AddSystemInfo("User Name", "Vinuchandran");
            extent1.LoadConfig(ConfigurationManager.AppSettings.Get("ExtentXML") + "Extent-Report.xml");
        }

        [TestInitialize]
        public void TestInitialize()
        {
            test = extent1.StartTest(testContextInstance.TestName+"_"+TestContext.DataRow.Table.Rows.IndexOf(TestContext.DataRow).ToString());
            _maxTestRuns = _maxTestRunsCount;
        }

        [TestMethod]
        public void CodedUITestMethod()
        {
            try
            {
                objHTMLControl.launchBrowser("http://www.testhouse.net/");
                objHTMLControl.clickHTMLHyperlinkOneProp("Contact");
                objHTMLControl.verifyHTMLImageTwoPropContains("WorldOffices");
                //objReport.ExcelReport(testContextInstance.TestName, "Passed", iteration);
                iteration++;
            }
            catch
            {
                iteration++;
                //objReport.ExcelReport(testContextInstance.TestName, "Failed", iteration);
            }

        }

        [TestMethod]
        public void CodedUITestMethod1()
        {
            MessageBox.Show("A");
            //objReport.ExcelReport(testContextInstance.TestName, "Passed", iteration);
            iteration++;
        }
        [DataSource("System.Data.Odbc", "Dsn=Excel Files;dbq=|DataDirectory|\\TestData\\Test.xlsx;defaultdir=.;driverid=1046;maxbuffersize=2048;pagetimeout=5", "Sheet1$", DataAccessMethod.Sequential), TestMethod]//[DataSource("System.Data.Odbc", "Dsn=ExcelFiles;Driver={Microsoft Excel Driver (*.xlsx)};dbq=|DataDirectory|\\Test.xlsx;defaultdir=.;driverid=1046;maxbuffersize=2048;pagetimeout=5;readonly=true", "Sheet1$", DataAccessMethod.Sequential), DeploymentItem("Test.xlsx"), TestMethod]//[DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\TestData\\TestData.csv", "TestData#csv", DataAccessMethod.Sequential), DeploymentItem("TestData.csv"), TestMethod]//[DataSource("System.Data.Odbc", "Dsn=Excel Files;dbq=C:\\Test.xlsx;defaultdir=.;driverid=1046;maxbuffersize=2048;pagetimeout=5", "Sheet1$", DataAccessMethod.Sequential), TestMethod]
        public void NewTest()
        {
            //objHTMLControl.clickHTMLHyperlinkOneProp("Test");
            //objHTMLControl.clickHTMLHyperlinkOneProp("Contact");
            objReport.CreateScreenshotSubFolders(testContextInstance.TestName, TestContext.DataRow.Table.Rows.IndexOf(TestContext.DataRow).ToString());
            string strValue = TestContext.DataRow["Test"].ToString();
            if (strValue == "1")
            {
                
                Console.WriteLine("Pass");
                test.Log(LogStatus.Pass, test.AddScreenCapture(objReport.ScreenshotPass("Test Passed")));
            }
            else
            {
                test.Log(LogStatus.Fail, test.AddScreenCapture(objReport.ScreenshotFail("Test Failed")));
                Assert.Fail("Failed");
            }
        }



        [DataSource("System.Data.Odbc", "Dsn=Excel Files;dbq=|DataDirectory|\\TestData\\Test.xlsx;defaultdir=.;driverid=1046;maxbuffersize=2048;pagetimeout=5", "Sheet1$", DataAccessMethod.Sequential), TestMethod]//[DataSource("System.Data.Odbc", "Dsn=ExcelFiles;Driver={Microsoft Excel Driver (*.xlsx)};dbq=|DataDirectory|\\Test.xlsx;defaultdir=.;driverid=1046;maxbuffersize=2048;pagetimeout=5;readonly=true", "Sheet1$", DataAccessMethod.Sequential), DeploymentItem("Test.xlsx"), TestMethod]//[DataSource("Microsoft.VisualStudio.TestTools.DataSource.CSV", "|DataDirectory|\\TestData\\TestData.csv", "TestData#csv", DataAccessMethod.Sequential), DeploymentItem("TestData.csv"), TestMethod]//[DataSource("System.Data.Odbc", "Dsn=Excel Files;dbq=C:\\Test.xlsx;defaultdir=.;driverid=1046;maxbuffersize=2048;pagetimeout=5", "Sheet1$", DataAccessMethod.Sequential), TestMethod]
        public void NewTestForReRun()
        {
            objReport.CreateScreenshotSubFolders(TestContext.TestName, TestContext.DataRow.Table.Rows.IndexOf(TestContext.DataRow).ToString());
            string strValue = TestContext.DataRow["Test"].ToString();       
            try
            {
                if (_maxTestRuns == 1)
                {
                    strValue = "1";
                }
                if (strValue == "1")
                {
                    Playback.Wait(1000);
                    //MessageBox.Show("Pass");
                    Console.WriteLine("Pass");
                    test.Log(LogStatus.Pass, test.AddScreenCapture(objReport.ScreenshotPass("Test Passed")));
                }
                else
                {
                    Playback.Wait(1000);
                    // MessageBox.Show("Fail"); 
                   // Assert.Fail("Failed");
                }
                objUIMap.UISolutionsTeshouseIntWindow.UISolutionsTeshouseDocument.UIContentsection1Custom.UIDynamicsCRMAXERPCRMsPane.DrawHighlight();
            }
            catch(Exception e)
            {
                if (e.InnerException is UITestControlNotFoundException)
                {
                    test.Log(LogStatus.Info, "Failed at this step" + e.Message, test.AddScreenCapture(objReport.ScreenshotFail("Test Failed")));
                    Type thisType = this.GetType();
                    Object testCall = this;
                    ObjGen.ReRun(TestContext.TestName, ref _maxTestRuns, _maxTestRunsCount, test, extent1, thisType, testCall, null, "MyTestCleanup");
                }
                
            }

        }

        [TestMethod]
        public void extent()
        {
            
               
        }



        #region Additional test attributes

        // You can use the following additional attributes as you write your tests:

        ////Use TestInitialize to run code before running each test 
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        ////Use TestCleanup to run code after each test has run
        [TestCleanup()]
        public void MyTestCleanup()
        {
            //objReport.ExcelReport(testContextInstance.TestName, testContextInstance.CurrentTestOutcome.ToString(), TestContext.DataRow.Table.Rows.IndexOf(TestContext.DataRow).ToString());
            extent1.EndTest(test);
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            extent1.Flush();
            extent1.Close();
        }

        #endregion

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        private TestContext testContextInstance;
    }
}
