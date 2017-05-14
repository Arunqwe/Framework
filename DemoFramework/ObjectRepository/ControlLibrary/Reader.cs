using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Collections.Specialized;

namespace DemoFramework.ObjectRepository.ControlLibrary
{
    public class Reader
    {
        private const string CONNECTION_STRING = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Extended Properties=Excel 12.0;Mode=Share Deny Write;";
        private const string QUERY = "select * from [{0}$]";
        public List<Keywords> screens;
        public Keywords Found = null;
        static string solutionPath = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.IndexOf("TestResults"));
        string controlPath = solutionPath + ConfigurationManager.AppSettings.Get("ControlsFilePath");
        public DataTable GetSchemaDataFromExcel(string SheetName, string fileName)
        {
            DataTable table = null;

            if (fileName == string.Empty)
            {
                return null;
            }

            string readQuery = string.Format(QUERY, SheetName);

            try
            {
                OleDbConnection con = new OleDbConnection(string.Format(CONNECTION_STRING, fileName));
                OleDbCommand com = new OleDbCommand(readQuery, con);
                con.Open();
                OleDbDataAdapter ad = new OleDbDataAdapter(com);
                table = new DataTable();
                ad.Fill(table);
            }
            catch
            {
                MessageBox.Show("Error while loading the spreadsheet");
                return null;
            }
            return table;
        }



        public List<Keywords> ReadKeywordTable(DataTable table, TestContext testcontext)
        {
            List<Keywords> screens = new List<Keywords>();

            string username = string.Empty;
            string password = string.Empty;

            foreach (DataRow row in table.Rows)
            {
                Keywords screen = new Keywords();
                
                string ControlName = string.Empty;
                string PropertyName1 = string.Empty;
                string PropertyValue1 = string.Empty;
                string PropertyName2 = string.Empty;
                string PropertyValue2 = string.Empty;
                //Edited
                string PropertyName3 = string.Empty;
                string PropertyValue3 = string.Empty;
                string PropertyName4 = string.Empty;
                string PropertyValue4 = string.Empty;
                string PropertyName5 = string.Empty;
                string PropertyValue5 = string.Empty;

                if (row["ControlName"] != null)
                {
                    ControlName = row["ControlName"].ToString();
                }
                if (row["PropertyName1"] != null)
                {
                    PropertyName1 = row["PropertyName1"].ToString();
                }
                if (row["PropertyValue1"] != null)
                {
                    PropertyValue1 = row["PropertyValue1"].ToString();
                }
                if (row["PropertyName2"] != null)
                {
                    PropertyName2 = row["PropertyName2"].ToString();
                }
                if (row["PropertyValue2"] != null)
                {
                    PropertyValue2 = row["PropertyValue2"].ToString();
                }

                //Edited
                if (row["PropertyName3"] != null)
                {
                    PropertyName3 = row["PropertyName3"].ToString();
                }
                if (row["PropertyValue3"] != null)
                {
                    PropertyValue3 = row["PropertyValue3"].ToString();
                }
                if (row["PropertyName4"] != null)
                {
                    PropertyName4 = row["PropertyName4"].ToString();
                }
                if (row["PropertyValue4"] != null)
                {
                    PropertyValue4 = row["PropertyValue4"].ToString();
                }
                if (row["PropertyName5"] != null)
                {
                    PropertyName5 = row["PropertyName5"].ToString();
                }
                if (row["PropertyValue5"] != null)
                {
                    PropertyValue5 = row["PropertyValue5"].ToString();
                }
                screen.ControlName = ControlName;
                screen.PropertyName1 = PropertyName1;
                screen.PropertyValue1 = PropertyValue1;
                screen.PropertyName2 = PropertyName2;
                screen.PropertyValue2 = PropertyValue2;

                //Edited
                screen.PropertyName3 = PropertyName3;
                screen.PropertyValue3 = PropertyValue3;
                screen.PropertyName4 = PropertyName4;
                screen.PropertyValue4 = PropertyValue4;
                screen.PropertyName5 = PropertyName5;
                screen.PropertyValue5 = PropertyValue5;

                //adding keyword to list
                screens.Add(screen);

            }
            return screens;

        }

        public Keywords FindControlinList(string test)
        {
            Reader reader = new Reader();
            DataTable tableKeyword = reader.GetSchemaDataFromExcel("Sheet1", controlPath);
            List<Keywords> screens = reader.ReadKeywordTable(tableKeyword, testContextInstance);
            for (int i=0;i<screens.Count;i++)
            {
                if(screens[i].ControlName == test)
                {
                    Found = screens[i];
                }
                else
                {

                }
            }
            return Found;
            
        }

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
