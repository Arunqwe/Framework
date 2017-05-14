using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xlObject = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.CSharp;
using System.Diagnostics;

namespace DemoFramework.GenericFunctions
{
    class Excel
    {
        // To read value from a particular cell in excel sheet
        public string readFromExcelCell(string strExcelPath,string strExcelCell)
        {
            xlObject.Application xlApp = new xlObject.Application();
            xlObject.Workbook xlWorkbook = xlApp.Workbooks.Open(strExcelPath);
            xlObject.Worksheet xlWorksheet = (xlObject.Worksheet)xlWorkbook.Sheets.get_Item(1);
            string strExcel = xlWorksheet.get_Range(strExcelCell).Cells.get_Value().ToString();// Provide the excel cell name [Eg: E5]
            xlWorkbook.Close();
            xlApp.Quit();
            return strExcel;
        }

        // To write value to a particular cell in excel sheet
        public void writeToExcelCell(string strValue, string strRange, string Filename)
        {
            xlObject.Application excel = new xlObject.Application();
            excel.Visible = true;
            xlObject.Workbook wb = excel.Workbooks.Open(Filename);
            xlObject.Worksheet sh = wb.Sheets.Application.ActiveSheet;// Add();
            sh.Cells.Range[strRange].Value = strValue;
            wb.Close(true);
            excel.Quit();
        }

        // To kill all of the _Excel instance in application
        public void killExcel()
        {
            Process[] iexplore = Process.GetProcessesByName("EXCEL");
            foreach (Process item in iexplore)
            {
                item.Kill();
            }
        }


    }
}
