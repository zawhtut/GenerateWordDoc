<%@ Page Language="C#" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="GenerateWordDoc" %>

<script runat="server">
    protected void btnExportToWord_Click(object sender, EventArgs e) {
        SalesReportBuilder report = new SalesReportBuilder(ddlEmployee.SelectedItem.Value);
        lblResult.Text = "Report saved at: " + report.CreateDocument();
    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Word 2007 Sales Report Generator</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <img src="Resources/headerImage.gif" style="width: 264px; height: 107px" alt="Adventure Works"/>
            <h1>
                <span style="font-family: Verdana">Sales Report Generator</span>
            </h1>
            <hr />
            <h3>
               <span style="font-family: Verdana; color:Maroon;">Create sales report for employee:</span>
            </h3>
            <table border="0" style="font-family: Verdana">
                <tr>
                    <td>
                        <asp:DropDownList
                        ID="ddlEmployee"
                        runat="server"
                        DataSourceID="EmployeeDataSource"
                        AutoPostBack="True"
                        DataTextField="FullName"
                        DataValueField="SalesPersonID"
                        Width="200px">
                        </asp:DropDownList>
                        <asp:SqlDataSource
                        ID="EmployeeDataSource"
                        runat="server"
                        ConnectionString="<%$ ConnectionStrings:AWConnString %>"
                        SelectCommand="SELECT [SalesPersonID],[FirstName]
                        + ' ' + [LastName] AS [FullName] FROM [Sales].[vSalesPerson]">
                        </asp:SqlDataSource>
                    </td>
                    <td valign="top">
                        <asp:Button
                        ID="btnExportToWord"
                        runat="server"
                        Text="Export to Word 2007"
                        OnClick="btnExportToWord_Click" />
                    </td>
                </tr>
            </table>
            <br />
            <h3>
               <span style="font-family: Verdana; color:Blue;">
                  <asp:Label
                  ID="lblResult"
                  runat="server" />
               </span>
            </h3>
        </div>
    </form>
</body>
</html>