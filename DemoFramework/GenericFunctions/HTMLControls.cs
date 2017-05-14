using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;
using System.Drawing;
using DemoFramework.Reports;
using System.Diagnostics;
using System.Windows.Input;
using DemoFramework.ObjectRepository.ControlLibrary;

namespace DemoFramework.GenericFunctions
{
    class HTMLControls
    {
        // Variable declaration
        ///////////////----------------------------------------------
        Point pt;
        BrowserWindow _browserWindow;
        bool vFlag = true;
        Reader reader = new Reader();
        //////////////-----------------------------------------------

        /// <summary>
        /// Launching the _browser and keeping the instance alive [Check the instance stay alive after execution]
        /// </summary>
        public void launchBrowser(string strURL)
        {
            BrowserWindow.CurrentBrowser = "ie";
            _browserWindow = BrowserWindow.Launch(strURL);
            _browserWindow.CloseOnPlaybackCleanup = false;
            _browserWindow.Maximized = true;
        }

        // Launching the _browser from Run window
        public void launchBrowserFromRun(string strURL)
        {
            Keyboard.SendKeys("r", ModifierKeys.Windows);
            Keyboard.SendKeys("iexplore " + strURL + "{ENTER}");
            Playback.Wait(5000);
        }

        // To kill all of the _browser instance in application[Internet Explorer]
        public void killBrowser()
        {
            Process[] iexplore = Process.GetProcessesByName("iexplore");
            foreach (Process item in iexplore)
            {
                item.Kill();
            }
        }

       



        ////-------------------------------------HTML CONTROLS----------------------------------
        // This file consist of a list of generic functions created for HTML controls
        //--------------------------------------------------------------------------------------

        // To click on HTML button accepting single property
        public void clickHTMLButtonOneProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlButton _htmlButton = new HtmlButton(_browser);
            _htmlButton.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            Mouse.Click(_htmlButton);
        }

       

        // To click HTML button accepting two property [Using TryGetClickablePoint Method]
        public void clickHTMLButtonCollectionTwoProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlButton _htmlButton = new HtmlButton(_browser);
            _htmlButton.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlButton.SearchProperties[KeyFound.PropertyName2] = KeyFound.PropertyValue2;
            _htmlButton.WaitForControlExist();
            int c = _htmlButton.FindMatchingControls().Count;
            var matchingControls = _htmlButton.FindMatchingControls();
            for (int i = 0; i < c; i++)
            {
                if (matchingControls[i].TryGetClickablePoint(out pt))
                {
                    Mouse.Click(matchingControls[i]);
                    break;
                }
            }
        }

        // To click HTML Button with four properties
        public void clickHTMLButtonFourProp(string PropertyName1, string PropertyValue1, string PropertyName2, string PropertyValue2, string PropertyName3, string PropertyValue3, string PropertyName4, string PropertyValue4)
        {

            BrowserWindow _browser = new BrowserWindow();
            HtmlButton _htmlButton = new HtmlButton(_browser);
            _htmlButton.SearchProperties["ControlType"] = "Button";
            _htmlButton.SearchProperties[PropertyName1] = PropertyValue1;
            _htmlButton.SearchProperties[PropertyName2] = PropertyValue2;
            _htmlButton.SearchProperties[PropertyName3] = PropertyValue3;
            _htmlButton.SearchProperties[PropertyName4] = PropertyValue4;
            if (_htmlButton.Exists)
            {
                //_htmlButton.DrawHighlight();
                Mouse.Click(_htmlButton);
            }
        }

        // To click on HTML list item control from html list
        public void clickHTMLLIstItemTwoPropWithParent(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlList _htmllist = new HtmlList(_browser);
            _htmllist.SearchProperties["ControlType"] = "List";
            _htmllist.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            if (_htmllist.Exists)
            {
                HtmlListItem _htmllistitem = new HtmlListItem(_htmllist);
                _htmllistitem.SearchProperties["ControlType"] = "ListItem";
                _htmllistitem.SearchProperties[KeyFound.PropertyName2] = KeyFound.PropertyValue2;
                Mouse.Click(_htmllistitem);
            }
            else
            {
                Console.WriteLine("List not found");
            }
        }

        // To click on HTML Label control based on two property with property expression operator contains
        public void clickHTMLLabelTwoPropContains(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlLabel _htmlLabel = new HtmlLabel(_browser);
            _htmlLabel.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlLabel.SearchProperties.Add(KeyFound.PropertyName2, KeyFound.PropertyValue2, PropertyExpressionOperator.Contains);
            if (_htmlLabel.Exists)
            {
                Mouse.Click(_htmlLabel);
            }
            else
            {
                
            }
        }

        // To enter data in HTML text area using two properties
        public void enterHTMLTextAreaTwoProp(string strControl, string strInput)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlTextArea _htmlTextArea = new HtmlTextArea(_browser);
            _htmlTextArea.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlTextArea.SearchProperties[KeyFound.PropertyName2] = KeyFound.PropertyValue2;
            _htmlTextArea.Text = strInput;
        }

        // To select HTML checkbox
        public void clickHTMLCheckBoxOneProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlCheckBox _htmlCheckBox = new HtmlCheckBox(_browser);
            _htmlCheckBox.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            Rectangle rec = _htmlCheckBox.BoundingRectangle;
            Mouse.Click(new Point(rec.X + rec.Width / 2, rec.Y + rec.Height / 2));
        }

        // To click HTML Custom accepting two property of which one property with property expression operator contains [Using TryGetClickablePoint Method]
        public void clickHTMLCustomCollectionTwoProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlCustom _htmlCustom = new HtmlCustom(_browser);
            _htmlCustom.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlCustom.SearchProperties.Add(KeyFound.PropertyName2, KeyFound.PropertyValue2, PropertyExpressionOperator.Contains);
            int c = _htmlCustom.FindMatchingControls().Count;
            var matchingControls = _htmlCustom.FindMatchingControls();
            for (int i = 0; i < c; i++)
            {
                if (matchingControls[i].TryGetClickablePoint(out pt))
                {
                    Mouse.Click(matchingControls[i]);
                    break;
                }
            }
        }

        // To click on HTML Span control based on single property
        public void clickHTMLSpanOneProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlSpan _htmlSpan = new HtmlSpan(_browser);
            _htmlSpan.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            if (_htmlSpan.Exists)
            {
                Mouse.Click(_htmlSpan);
            }
            else
            {
                
            }
        }

        // To select a value from HTML ComboBox based on single property
        public void selectHTMLComboBoxOneProp(string strControl, string strInput)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlComboBox _htmlComboBox = new HtmlComboBox(_browser);
            _htmlComboBox.SearchProperties.Add(KeyFound.PropertyName1, KeyFound.PropertyValue1);
            _htmlComboBox.SelectedItem = strInput;
        }

        // To select HTML radio button based on single property
        public void selectHTMLRadioButtonOneProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlRadioButton _htmlRadioButton = new HtmlRadioButton(_browser);
            _htmlRadioButton.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            if (_htmlRadioButton.Exists)
            {
                _htmlRadioButton.Selected = true;
            }
            else
            {
            }
        }

        // To enter value into HTML Edit control using two property
        public void enterHTMLEditTwoProp(string strControl, string strInput)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlEdit _htmlEdit = new HtmlEdit(_browser);
            _htmlEdit.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlEdit.SearchProperties[KeyFound.PropertyName2] = KeyFound.PropertyValue2;
            _htmlEdit.Text = strInput;
        }

        // To enter value into HTML Edit control using two property
        public void clickEnterHTMLEditTwoProp(string strControl, string strInput)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlEdit _htmlEdit = new HtmlEdit(_browser);
            _htmlEdit.SearchProperties.Add(KeyFound.PropertyName1, KeyFound.PropertyValue1);
            _htmlEdit.SearchProperties.Add(KeyFound.PropertyName2, KeyFound.PropertyValue2);
            Mouse.Click(_htmlEdit);
            Keyboard.SendKeys(strInput);
        }

        // To enter value into HTML Edit control using two property [contains]
        public void enterHTMLEditTwoPropContains(string strControl, string strInput)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlEdit _htmlEdit = new HtmlEdit(_browser);
            _htmlEdit.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlEdit.SearchProperties.Add(KeyFound.PropertyName2, KeyFound.PropertyValue2, PropertyExpressionOperator.Contains);
            _htmlEdit.Text = strInput;
        }

        // Validate HTML Edit value using two property and provide a bool output [False if failed]
        public bool verifyHTMLEditTwoProp(string strControl, string strInput)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlEdit _htmlEdit = new HtmlEdit(_browser);
            _htmlEdit.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlEdit.SearchProperties[KeyFound.PropertyName2] = KeyFound.PropertyValue2;
            if (_htmlEdit.Text.Contains(strInput))
            {
                Console.WriteLine("Pass: Verification successfully done");
                Console.WriteLine("Field contains :" + _htmlEdit.Text);
            }
            else
            {
                Console.WriteLine("Failed: Verification Failure");
                vFlag = false;
            }
            return vFlag;
        }

        // To click on HTML Image accepting single property
        public void clickHTMLImageOneProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlImage _htmlImage = new HtmlImage(_browser);
            _htmlImage.SearchProperties["ControlType"] = "Image";
            _htmlImage.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            if (_htmlImage.Exists)
            {
                Mouse.Click(_htmlImage);
            }
        }

        // To click on HTML Image accepting two property
        public void clickHTMLImageTwoProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlImage _htmlImage = new HtmlImage(_browser);
            _htmlImage.SearchProperties["ControlType"] = "Image";
            _htmlImage.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlImage.SearchProperties[KeyFound.PropertyName2] = KeyFound.PropertyValue2;
            if (_htmlImage.Exists)
            {
                Mouse.Click(_htmlImage);
            }
        }

        // Validate HTML Image using two property [contains]
        public bool verifyHTMLImageTwoPropContains(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlImage _htmlImage = new HtmlImage(_browser);
            _htmlImage.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlImage.SearchProperties.Add(KeyFound.PropertyName2, KeyFound.PropertyValue2, PropertyExpressionOperator.Contains);
            if (_htmlImage.Exists)
            {
                Console.WriteLine("Pass: Verification successfully done");
                
            }
            else
            {
                Console.WriteLine("Failed: Verification Failure");
                vFlag = false;
            }
            return vFlag;
        }

        // Click on cell in HTML Table by passing table name and verification values[based on data on a column], row and column index as parameters
        public bool clickHTMLTableCell(HtmlTable UITable,string strVerifyValue, int iCell1, int iCell2)
        {
            
            HtmlCell verifyCell, clickCell;
            for (int i = 0; i <= UITable.RowCount - 1; i++)
            {
                verifyCell = ((HtmlCell)((HtmlRow)UITable.Rows[i]).Cells[iCell1]);
                clickCell = ((HtmlCell)((HtmlRow)UITable.Rows[i]).Cells[iCell2]);

                if (verifyCell.InnerText.Contains(strVerifyValue))
                {
                    Rectangle rec = clickCell.BoundingRectangle;
                    Mouse.Click(new Point(rec.X + rec.Width / 2, rec.Y + rec.Height / 2));
                    vFlag = false;
                    break;
                }
            }
            return vFlag;
        }
        //To verify table cell value
        public bool verifyTableCell(HtmlTable UITable, string strVerifyValue, int Cell)
        {
            bool Flag = false;
            HtmlCell verifyCell;
            for (int i = 0; i <= UITable.RowCount - 1; i++)
            {
                verifyCell = ((HtmlCell)((HtmlRow)UITable.Rows[i]).Cells[Cell]);
                
                if (verifyCell.InnerText.Contains(strVerifyValue))
                {
                    Console.WriteLine("Verified the content and content is " + verifyCell.InnerText);
                    Flag= true;
                    break;
                }
                else
                {
                    Console.WriteLine("Cell contains different value and the content is " + verifyCell.InnerText);
                }
            }
            return Flag;
            
        }

        // To click HTML Table Cell accepting two property of which one property with property expression operator contains [Using TryGetClickablePoint Method]
        public void clickHTMLTableCellTwoPropContains(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlCell _htmlTableCell = new HtmlCell(_browser);
            _htmlTableCell.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlTableCell.SearchProperties.Add(KeyFound.PropertyName2, KeyFound.PropertyValue2, PropertyExpressionOperator.Contains);
            Mouse.Click(_htmlTableCell);
        }

        // To Click on HTML Div control
        public void clickHTMLButtonInsideDivTwoProp(string strControl, string actualresult)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            bool flag = false;
            BrowserWindow _browser = new BrowserWindow();
            HtmlDiv _htmlDiv = new HtmlDiv(_browser);
            _htmlDiv.SearchProperties.Add(KeyFound.PropertyName1, KeyFound.PropertyValue1);
            _htmlDiv.SearchProperties.Add(KeyFound.PropertyName2, KeyFound.PropertyValue2);

            int c = _htmlDiv.FindMatchingControls().Count;
            var matchingControls = _htmlDiv.FindMatchingControls();
            for (int i = 0; i < c; i++)
            {
                if (matchingControls[i].Exists)
                {
                    ICollection<UITestControl> ui = matchingControls[i].GetChildren();
                    foreach (HtmlButton control in ui)
                    {
                        try
                        {
                            Mouse.Click(control);
                            flag = true;
                            break;

                        }
                        catch { }
                    }
                }
                if (flag == true)
                {
                    break;
                }
            }
        }

        //Function applicable to particular Websites
        // To click HTML Div based on single property
        public void clickHTMLDivOneProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlDiv _htmlDiv = new HtmlDiv(_browser);
            _htmlDiv.SearchProperties.Add(KeyFound.PropertyName1, KeyFound.PropertyValue1);
            _htmlDiv.DrawHighlight();
            Rectangle rec = _htmlDiv.BoundingRectangle;
            Mouse.Click(new Point(rec.X + rec.Width / 2, rec.Y + rec.Height / 2));
        }
        //Function applicable to particular Websites
        // To click HTML Div based on single property
        public void clickHTMLDivTwoPropWithParent(string strControl, UITestControl parent)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlDiv _htmlDiv = new HtmlDiv(parent);
            _htmlDiv.SearchProperties.Add(KeyFound.PropertyName1, KeyFound.PropertyValue1);
            _htmlDiv.SearchProperties.Add(KeyFound.PropertyName2, KeyFound.PropertyValue2);
            Rectangle rec = _htmlDiv.BoundingRectangle;
            Mouse.Click(new Point(rec.X + rec.Width / 2, rec.Y + rec.Height / 2));
        }

        //Function applicable to particular Websites
        // To click HTML Div based on single property
        public void clickHTMLDivTwoProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlDiv _htmlDiv = new HtmlDiv();
            _htmlDiv.SearchProperties.Add(KeyFound.PropertyName1, KeyFound.PropertyValue1);
            _htmlDiv.SearchProperties.Add(KeyFound.PropertyName2, KeyFound.PropertyValue2);
            UITestControlCollection test = _htmlDiv.FindMatchingControls();
            foreach (UITestControl test2 in test)
            {
                test2.DrawHighlight();
            }
            Rectangle rec = _htmlDiv.BoundingRectangle;
            Mouse.Hover(new Point(rec.X + rec.Width / 2, rec.Y + rec.Height / 2));
            Mouse.Click();
        }

        //Function applicable to particular Websites
        // To hover HTML Div based on single property
        public void hoverHtmlDivOneProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlDiv _htmlDiv = new HtmlDiv(_browser);
            _htmlDiv.SearchProperties.Add(KeyFound.PropertyName1, KeyFound.PropertyValue1);
            Rectangle rec = _htmlDiv.BoundingRectangle;
            Mouse.Hover(new Point(rec.X + rec.Width / 2, rec.Y + rec.Height / 2));
        }

        //Function applicable to particular Websites
        // To check for all hyperlinks inside div control
        public void searchHTMLHyperlinksInsideDivOneProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            //Search for the Div which has the list of hyperlinks we are searching
            HtmlDiv _htmlDiv = new HtmlDiv();
            _htmlDiv.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlDiv.SearchProperties[KeyFound.PropertyName2] = KeyFound.PropertyValue2;
             //Pass the instance of Div to HtmlControl class
            HtmlControl _htmlControl = new HtmlControl(_htmlDiv);
            _htmlControl.SearchProperties.Add(HtmlControl.PropertyNames.ClassName, "HtmlHyperlink");
            UITestControlCollection _uitestControlCollection = _htmlControl.FindMatchingControls();
            foreach (UITestControl links in _uitestControlCollection)
            {
                //cast the item to HtmlHyperlink type
                HtmlHyperlink _htmlHyperlink = (HtmlHyperlink)links;
                //get the innertext from the link, which inturn returns the link value itself
                Console.WriteLine(_htmlHyperlink.InnerText);
            }
        }
        // To check for all hyperlinks inside custom control
        public void searchHTMLHyperlinksInsideCustomTwoProp(HtmlTable UITable,string strControl)
        {
            
            Keywords KeyFound = reader.FindControlinList(strControl);
            HtmlCustom _htmlCustom = new HtmlCustom(UITable);
            _htmlCustom.SearchProperties.Add("ControlType", "Custom");
            _htmlCustom.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlCustom.SearchProperties[KeyFound.PropertyName2] = KeyFound.PropertyValue2;
            HtmlControl _htmlControl = new HtmlControl(_htmlCustom);
            _htmlControl.SearchProperties.Add(HtmlControl.PropertyNames.ClassName, "HtmlHyperlink");
            UITestControlCollection _uitestControlCollection = _htmlControl.FindMatchingControls();
            foreach (UITestControl links in _uitestControlCollection)
            {
                //cast the item to HtmlHyperlink type
                HtmlHyperlink _htmlHyperlink = (HtmlHyperlink)links;
                //get the innertext from the link, which inturn returns the link value itself
                Console.WriteLine(_htmlHyperlink.InnerText);
            }

        }

        // To click on specific hyperlink based on the property from a list of hidden hyperlink controls
        public void clickHTMLHyperlinkFromHiddenControlsOneProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlHyperlink _htmlHyperlink = new HtmlHyperlink(_browser);
            _htmlHyperlink.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlHyperlink.WaitForControlExist();
            UITestControlCollection _uitestControlCollection = _htmlHyperlink.FindMatchingControls();
            foreach (UITestControl _uitestControl in _uitestControlCollection)
            {
                if (_uitestControl.BoundingRectangle.Width > 0)
                {
                    Mouse.Click(_uitestControl);
                    break;
                }
            }
        }

        // To click on a specific hyperlink
        public void clickHTMLHyperlinkOneProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlHyperlink _htmlHyperlink = new HtmlHyperlink(_browser);
            _htmlHyperlink.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlHyperlink.SearchProperties[KeyFound.PropertyName2] = KeyFound.PropertyValue2;
            _htmlHyperlink.SearchProperties[KeyFound.PropertyName3] = KeyFound.PropertyValue3;
            _htmlHyperlink.WaitForControlExist();
            //Mouse.Click(_htmlHyperlink);
            _htmlHyperlink.DrawHighlight();
        }


        // To click on a specific hyperlink with two property
        public void clickHTMLHyperlinkTwoProp(string strControl)
        {
            Keywords KeyFound = reader.FindControlinList(strControl);
            BrowserWindow _browser = new BrowserWindow();
            HtmlHyperlink _htmlHyperlink = new HtmlHyperlink(_browser);
            _htmlHyperlink.SearchProperties[KeyFound.PropertyName1] = KeyFound.PropertyValue1;
            _htmlHyperlink.SearchProperties[KeyFound.PropertyName2] = KeyFound.PropertyValue2;
            Mouse.Click(_htmlHyperlink);
        }
    }
}
