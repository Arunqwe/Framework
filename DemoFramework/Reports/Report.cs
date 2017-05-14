using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel1 = Microsoft.Office.Interop.Excel;
using xlObject = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using DemoFramework;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Configuration;
using System.Collections.Specialized;
using RelevantCodes.ExtentReports;

namespace DemoFramework.Reports
{
    class Report
    {
        public static Excel1.Worksheet worksheet1 = null;
        public static Excel1.Worksheet worksheet2 = null;
        public static string resultsDirectoryPath = string.Empty;
        public static TimeSpan previousTestEndTime = new TimeSpan();
        public static string fPath { get; set; }
        public static int stepNum = 1;
        public static object missing = Type.Missing;
        public static int passedCount = 0;
        public static int failedCount = 0;
        public string strScreenshot="";

        public string ScreenpathStringPass;
        public string ScreenathStringFail;
        public string strExtentSSpassPath;
        public string strExtentSSfailPath;
        string DatapathString;
        DemoFramework.GenericFunctions.Excel objExcel = new GenericFunctions.Excel();
        
        // To write the test method result to an excel sheet after execution with Test Name, Iteration, date, time and duration
        public void ExcelReport(string strTestName,string strStatus,string iIteration)
        {
            // Code for Excel report
            xlObject.Application xlApp = new xlObject.Application();
            xlApp.ScreenUpdating = false;
            xlObject.Workbook xlWorkbook;
            string solutionPath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.IndexOf("TestResults"));
            string ExcelPath = solutionPath + ConfigurationManager.AppSettings.Get("ReportsPath");
            xlWorkbook = xlApp.Workbooks.Open(ExcelPath);
            xlObject.Worksheet xlWorksheet = (xlObject.Worksheet)xlWorkbook.Sheets.get_Item(1);
            xlObject.Range last = xlWorksheet.Cells.SpecialCells(xlObject.XlCellType.xlCellTypeLastCell, Type.Missing);
            xlObject.Range range = xlWorksheet.get_Range("A1", last);

            int lastUsedRow = last.Row;
            int lastUsedColumn = last.Column;
            xlWorksheet.Cells[1, 1].EntireRow.Font.Bold = true;
            xlWorksheet.Cells[1, 1] = "Test Name";
            xlWorksheet.Cells[1, 2] = "Data Row";
            xlWorksheet.Cells[1, 3] = "Test Result";
            xlWorksheet.Cells[1, 4] = "Execution Date";
            xlWorksheet.Cells[1, 5] = "Execution Time";
            xlWorksheet.Cells[1, 6] = "Screenshot";
            xlWorksheet.Cells[lastUsedRow + 1, 1] = strTestName;
            xlWorksheet.Cells[lastUsedRow + 1, 2] = iIteration.ToString();
            if (strStatus == "Passed")
            {
                xlWorksheet.Cells[lastUsedRow + 1, 3] = "Passed";
                xlWorksheet.Cells[lastUsedRow + 1, 3].Interior.Color = ColorTranslator.ToOle(Color.Green);
            }
            if(strStatus == "Failed")
            {
                xlWorksheet.Cells[lastUsedRow + 1, 3] = "Failed";
                xlWorksheet.Cells[lastUsedRow + 1, 3].Interior.Color = ColorTranslator.ToOle(Color.Red);
            }
            xlWorksheet.Cells[lastUsedRow + 1, 4] = DateTime.Now.ToString("dd/MM/yyyy");
            xlWorksheet.Cells[lastUsedRow + 1, 5] = DateTime.Now.ToString("HH:mm:ss");
            xlWorksheet.Hyperlinks.Add(xlWorksheet.Cells[lastUsedRow + 1, 6], DatapathString);

            xlWorkbook.Save();
            xlWorkbook.Close();
            xlApp.Quit();
            objExcel.killExcel();
        }

        //Create sub Folders
        public void CreateScreenshotSubFolders(string name,string DataRow)
        {
            DateTime dt = DateTime.Now;
            string solutionPath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.IndexOf("TestResults"));
            string ExcelPath = solutionPath + ConfigurationManager.AppSettings.Get("ScreenshotPath");
            //Create a subfolder - TestMethod Name
            string pathString = System.IO.Path.Combine(ExcelPath, name);
            System.IO.Directory.CreateDirectory(pathString);
            // Create Data Row Folder
            DatapathString= System.IO.Path.Combine(pathString, "DataRow_"+DataRow);
            System.IO.Directory.CreateDirectory(DatapathString);
            //Create SubFolders
            ScreenpathStringPass = System.IO.Path.Combine(DatapathString, "Pass");
            ScreenathStringFail = System.IO.Path.Combine(DatapathString, "Fail");
            System.IO.Directory.CreateDirectory(ScreenpathStringPass);
            System.IO.Directory.CreateDirectory(ScreenathStringFail);
        }

        // To take screenshots
        public string ScreenshotPass(string strComment)
        {
            DateTime dt = DateTime.Now;
            Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics graphics = Graphics.FromImage(printscreen as Image);
            graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);
            strExtentSSpassPath = new Uri(ScreenpathStringPass + "/" + dt.ToString("dd-MM-yyyy-hh-mm-ss-tt") + "_" + strComment + ".jpg").LocalPath;
            printscreen.Save(strExtentSSpassPath, ImageFormat.Jpeg);//@"C:\Users\Vinuchandran\Desktop\Excel\Framework V3.0\DemoFramework_Solution\ResultReport\Screenshot\"
            return strExtentSSpassPath;
        }

        // To take screenshots
        public string ScreenshotFail(string strComment)
        {
            DateTime dt = DateTime.Now;
            Bitmap printscreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics graphics = Graphics.FromImage(printscreen as Image);
            graphics.CopyFromScreen(0, 0, 0, 0, printscreen.Size);
            strExtentSSfailPath = new Uri(ScreenathStringFail + "/" + dt.ToString("dd-MM-yyyy-hh-mm-ss-tt") + "_" + strComment + ".jpg").LocalPath;
            printscreen.Save(strExtentSSfailPath, ImageFormat.Jpeg);//@"C:\Users\Vinuchandran\Desktop\Excel\Framework V3.0\DemoFramework_Solution\ResultReport\Screenshot\"
            return strExtentSSfailPath;
        }

        // Report
        /// <summary>
        /// User this function to enter the parameters to be displayed in the report
        /// </summary>
        public void ReportLog(ExtentTest TestInstance, string VerificationStatus, string ScreenshotComment)
        {

        }

    }
}
