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


namespace DemoFramework.TestScript
{
    /// <summary>
    /// Summary description for CodedUITest1
    /// </summary>
    [CodedUITest]
    public class CodedUITest1
    {
        DemoFramework.ObjectRepository.UIMaps.UIMapClasses.UIMap obj = new ObjectRepository.UIMaps.UIMapClasses.UIMap();
        public CodedUITest1()
        {
        }

        [TestMethod]
        public void CodedUITestMethod1()
        {
            if (obj.UISolutionsTeshouseIntWindow.UISolutionsTeshouseDocument.UIContentsection1Custom.UIDynamicsCRMAXERPCRMsPane.InnerText.Contains("Dynamics CRM & AX"))
            {
                MessageBox.Show("Pass");
            }
            else
            {
                Assert.Fail();
            }

        }

        #region Additional test attributes

        // You can use the following additional attributes as you write your tests:

        ////Use TestInitialize to run code before running each test 
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        //Use TestCleanup to run code after each test has run
        [TestCleanup()]
        public void CleanUP()
        {
            MessageBox.Show("Clean1");
            if (TestContext.CurrentTestOutcome == UnitTestOutcome.Failed)
            {
                var type = Type.GetType(TestContext.FullyQualifiedTestClassName);
                if (type != null)
                {
                    type = null;
                    MessageBox.Show("Clean2");
                    var instance = Activator.CreateInstance(type.ReflectedType);
                    var method = type.GetMethod(TestContext.TestName);
                    try
                    {
                        MessageBox.Show("Clean3");
                        method.Invoke(instance, null);
                    }
                    catch
                    {

                    }
                }
            }
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
