using Microsoft.VisualStudio.TestTools.UnitTesting;
using RelevantCodes.ExtentReports;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;


namespace DemoFramework.GenericFunctions
{
    class Generic
    {

        // To Rerun the method by providing test name,retry count,Max no. of rerun count,extenttest report instant to be updated, Extenttest report , Classtype, Object and (test initialize and testCleanup method if needed).
        public void ReRun(string testName, ref int _maxTestRuns, int _maxTestRunsCount, ExtentTest test, ExtentReports extent1, Type ClassType, object ObjTestFail, string testInitializeName = null, string testCleanupName = null)
        {
            if (_maxTestRuns > 0)
            {
                _maxTestRuns--;
                if (testCleanupName != null)
                {
                    MethodInfo theCleanupMethod = ClassType.GetMethod(testCleanupName);
                    theCleanupMethod.Invoke(ObjTestFail, null);
                }
                else
                {
                    extent1.EndTest(test);
                }
                test.Log(LogStatus.Info, $"Running Again, Retry count:{_maxTestRunsCount - _maxTestRuns} ");

                if (testInitializeName != null)
                {
                    MethodInfo theInitializeMethod = ClassType.GetMethod(testInitializeName);
                    theInitializeMethod.Invoke(ObjTestFail, null);
                }

                MethodInfo theMethod = ClassType.GetMethod(testName);
                theMethod.Invoke(ObjTestFail, null);
            }
            else
            {
                test.Log(LogStatus.Fail, $"Failed even after retrying {_maxTestRunsCount} times");
                Assert.Fail("Failed even after retrying");
            }


        }


    }
}
