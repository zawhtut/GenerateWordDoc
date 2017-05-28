using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Specialized;
using System.Configuration;

namespace GenerateWordDoc
{
    /// <summary>
    /// Represents the AdventureWorks line of business sales data.
    /// </summary>
    public class AdventureWorksSalesData
    {
        private string _source;

        public AdventureWorksSalesData()
        {
            _source = ConfigurationManager.ConnectionStrings["AWConnString"].ConnectionString;
        }

        /// <summary>
        /// Get employee's data: FullName, Phone, Email,
        /// SalesQuota, SalesYTD, TerritoryName.
        /// </summary>
        /// <param name="salesPersonID">Data for employee with SalesPersonID</param>
        /// <returns>StringDictionary</returns>
        public StringDictionary GetSalesPersonData(string salesPersonID)
        {
            StringDictionary SalesPerson = new StringDictionary();

            //Connect to a Microsoft SQL Server database and get data
            const string query = "SELECT [FirstName] + ' ' + [LastName] AS [FullName],[Phone],[EmailAddress], [TerritoryName], [SalesQuota], [SalesYTD] FROM[AdventureWorks].[Sales].[vSalesPerson] WHERE([SalesPersonID] = @SalesPersonID)";


        using (SqlConnection conn = new SqlConnection(_source))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@SalesPersonID", salesPersonID);
                SqlDataReader dr = cmd.ExecuteReader();

                if (dr.Read())
                {

                    SalesPerson.Add("FullName", (string)dr["FullName"]);
                    SalesPerson.Add("Phone", (string)dr["Phone"]);
                    SalesPerson.Add("Email", (string)dr["EmailAddress"]);
                    if (dr["SalesQuota"] is System.DBNull)
                    {
                        SalesPerson.Add("SalesQuota", "0");
                    }
                    else
                    {
                        SalesPerson.Add("SalesQuota", Convert.ToString((decimal)dr["SalesQuota"]));
                    }
                    if (dr["SalesYTD"] is System.DBNull)
                    {
                        SalesPerson.Add("SalesYTD", "0");
                    }
                    else
                    {
                        SalesPerson.Add("SalesYTD", Convert.ToString((decimal)dr["SalesYTD"]));
                    }
                    if (dr["TerritoryName"] is System.DBNull)
                    {
                        SalesPerson.Add("TerritoryName", "NA");
                    }
                    else
                    {
                        SalesPerson.Add("TerritoryName", (string)dr["TerritoryName"]);
                    }
                }

                dr.Close();
                conn.Close();
            }

            return SalesPerson;
        }

        /// <summary>
        /// Get a table that shows all sales for 2003 and 2004 
        /// for all employees that belong to employee's
        /// sales territory so you can compare sales data
        /// and performance.
        /// </summary>
        /// <param name="Territory">Get all sales/employee for territory</param>
        /// <returns>Table with 2003 and 2004 sales/employee</returns>
        public DataTable GetSalesByTerritory(string Territory)
        {
            //Connect to a Microsoft SQL Server database and get data
            String source = ConfigurationManager.ConnectionStrings["AWConnString"].
            ConnectionString;
            string query = "SELECT [FullName], [2003], [2004] FROM [AdventureWorks].[Sales].[vSalesPersonSalesByFiscalYears] WHERE([SalesTerritory] ='" + Territory + "')";


        using (SqlConnection conn = new SqlConnection(_source))
            {
                conn.Open();
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "SalesByTerritory");
                conn.Close();

                return ds.Tables["SalesByTerritory"];
            }
        }

    }
}