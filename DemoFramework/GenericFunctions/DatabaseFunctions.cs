using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace DemoFramework.GenericFunctions
{
    class DatabaseFunctions
    {
        /// <summary>
        /// Variable declaration section
        /// </summary>
        public string strQueryResult;

        // The below list contains functions to delete data from database based on a query
        public void deleteDataFromDB(string strQuery)
        {
            // Specify the connection string to connect to the database
            string SqlCon = "Data Source=EQAUTOMATION2;Initial Catalog=WorkManagementAzureDev;User ID=WorkManagementAzureDev;Password=WorkManagementAzureDev";
            SqlConnection thisConn = new SqlConnection(SqlCon);
            using (SqlConnection myConnection = new SqlConnection(SqlCon))
            {
                // Sample strQuery = "Delete from [WorkManagementAzureDev].[dbo].[tempMemberData] where FirstName = 'TestMember'";
                SqlCommand oComm = new SqlCommand(strQuery, myConnection);
                myConnection.Open();
                oComm.ExecuteNonQuery();
                myConnection.Close();
            }
        }

        // The function will read data from DB and return the value
        public string readDataFromDB(string strQuery, string strConnectionString, string strColumnName)
        {
            string SqlCon = strConnectionString;// Specify Data Connection string [Eg: Data Source=EQAUTOMATION;Initial Catalog=WorkManagementAzureTest;User ID=WorkManagementAzureTest;Password=WorkManagementAzureTest]
            string SqlQuery = strQuery;// Query passed from the function
            SqlConnection thisConn = new SqlConnection(SqlCon);
            using (SqlConnection myConnection = new SqlConnection(SqlCon))
            {
                SqlCommand oComm = new SqlCommand(SqlQuery, myConnection);
                myConnection.Open();
                using (SqlDataReader oReader = oComm.ExecuteReader())
                {
                    while (oReader.Read())
                    {
                        strQueryResult = oReader[strColumnName].ToString();// Specify the column name from which data needs to be fetched
                    }
                }
            }
            return strQueryResult;
        }

        // This function will write data to DB
        public void writeDataToDB(string strQuery)
        {
            string SqlCon = "Data Source=EQAUTOMATION;Initial Catalog=WorkManagementAzureTest;User ID=WorkManagementAzureTest;Password=WorkManagementAzureTest";
            SqlConnection thisConn = new SqlConnection(SqlCon);
            using (SqlConnection myConnection = new SqlConnection(SqlCon))
            {
                SqlCommand oComm = new SqlCommand(strQuery, myConnection);// Specify the query as per the example [eg: "SET IDENTITY_INSERT [WorkManagementAzureTest].[dbo].[BusinessUnit] ON Insert [dbo].[BusinessUnit] ([Id],[Name],[Description],[CreatedById],[UnitLeadId]) Values (@Id,'CompendiaTestBusinessUnit','Test Business Unit', Null,Null) SET IDENTITY_INSERT [dbo].[BusinessUnit] OFF";]
                myConnection.Open();
                oComm.Parameters.AddWithValue("", "");// Specify the parameter name with value as per the format provided [eg: oComm.Parameters.AddWithValue("@Id", Id);]
                oComm.ExecuteNonQuery();
                myConnection.Close();
            }
        }


    }
}
