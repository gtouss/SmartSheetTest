using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;
using System.Collections;

namespace SmartSheetTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Set the Access Token
            Token token = new Token();
            token.AccessToken = System.Configuration.ConfigurationManager.AppSettings["smartSheetAccessToken"].ToString();

            // Use the Smartsheet Builder to create a Smartsheet
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();

            // Get home
            Home home = smartsheet.Home().GetHome(new ObjectInclusion[] { ObjectInclusion.TEMPLATES });

            // List Sheets
            IList<Sheet> homeSheets = smartsheet.Sheets().ListSheets();
            foreach (Sheet tmpSheet in homeSheets)
            {
                
            	Console.WriteLine("sheet:" + tmpSheet.ID + " " + tmpSheet.Name);
                Output.Text = Output.Text + tmpSheet.ID + " " + tmpSheet.Name + "\r\n";
            }
        }

        private void UpdateSheet_Click(object sender, EventArgs e)
        {
            // Set the Access Token
            Token token = new Token();
            token.AccessToken = System.Configuration.ConfigurationManager.AppSettings["smartSheetAccessToken"].ToString();

            // Use the Smartsheet Builder to create a Smartsheet
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();

            var sheetList = getSheetList();

            long sheetID = sheetList.First(kvp => kvp.Key == "Old Incidents").Value;

            //Get the list of columns
            var columnList = getColumnList(sheetID);

            connString = "Server=srv-it3.celerant.com;Database=SupportSSRS;Trusted_Connection=Yes;";
            DataTable table1 = new DataTable();
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                cmd = "select customer as Client,Incident_ID as IncidentID,AssignedTo,cast(StartDate as date) as StartDate,cast(LastActionDate as date) as LastActionDate,Problem,TeamName from supportssrs.dbo.TB_INCIDENT where StartDate<dateadd(dd,-7,cast(getdate() as date)) and TeamName like '%team%' and enddate is null order by TeamName";
                using (SqlDataAdapter reader = new SqlDataAdapter(cmd, conn))
                {
                    reader.Fill(table1);
                    dataGridView1.DataSource = table1;
                }
            }
            foreach (DataRow dtrow in table1.Rows)
            {
                Cell cellClient = new Cell();
                cellClient.Value = dtrow["Client"];
                cellClient.Strict = false;
                cellClient.ColumnId = columnList.First(kvp => kvp.Key == "Client").Value;

                Cell cellIncidentID = new Cell();
                cellIncidentID.Value = dtrow["IncidentID"];
                cellIncidentID.ColumnId = columnList.First(kvp => kvp.Key == "IncidentID").Value;

                Cell cellAssignedTo = new Cell();
                cellAssignedTo.Value = dtrow["AssignedTo"];
                cellAssignedTo.ColumnId = columnList.First(kvp => kvp.Key == "AssignedTo").Value;
                
                Cell cellStartDate = new Cell();
                cellStartDate.Value = dtrow["StartDate"];
                cellStartDate.ColumnId = columnList.First(kvp => kvp.Key == "StartDate").Value;
                
                Cell cellLastActionDate = new Cell();
                cellLastActionDate.Value = dtrow["LastActionDate"];
                cellLastActionDate.ColumnId = columnList.First(kvp => kvp.Key == "LastActionDate").Value;
               
                Cell cellProblem = new Cell();
                cellProblem.Value = dtrow["Problem"];
                cellProblem.ColumnId = columnList.First(kvp => kvp.Key == "Problem").Value;
                
                Cell cellTeamName = new Cell();
                cellTeamName.Value = dtrow["TeamName"];
                cellTeamName.ColumnId = columnList.First(kvp => kvp.Key == "TeamName").Value;

                //// Store the cells in a list
                List<Cell> cellList = new List<Cell>();
                cellList.Add(cellClient);
                cellList.Add(cellIncidentID);
                cellList.Add(cellAssignedTo);
                cellList.Add(cellStartDate);
                cellList.Add(cellLastActionDate);
                cellList.Add(cellProblem);
                cellList.Add(cellTeamName);
                //// Create a row and add the list of cells to the row
                Row row = new Row();

                row.Cells = cellList;
                //// Add two rows to a list of rows.
                List<Row> rows = new List<Row>();
                rows.Add(row);
                RowWrapper rowWrapper = new RowWrapper.InsertRowsBuilder().SetRows(rows).SetToBottom(true).Build();
                smartsheet.Sheets().Rows().InsertRows(1449506131732356, rowWrapper);
            }
        }

        private void ListColumns_Click(object sender, EventArgs e)
        {
            // Set the Access Token
            Token token = new Token();
            token.AccessToken = System.Configuration.ConfigurationManager.AppSettings["smartSheetAccessToken"].ToString();

            // Use the Smartsheet Builder to create a Smartsheet
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();
            IList<Column> sheetCols = smartsheet.Sheets().Columns().ListColumns(1449506131732356);

            foreach (Column tmpCols in sheetCols)
            {

                Console.WriteLine("sheet:" + tmpCols.ID + " " + tmpCols.Title);
                Output.Text = Output.Text + tmpCols.ID + " " + tmpCols.Title + "\r\n";
            }
        }

        private List<KeyValuePair<string, long>> getSheetList()
        {
            Token token = new Token();
            token.AccessToken = System.Configuration.ConfigurationManager.AppSettings["smartSheetAccessToken"].ToString();

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();
            Home home = smartsheet.Home().GetHome(new ObjectInclusion[] { ObjectInclusion.TEMPLATES });
            IList<Sheet> homeSheets = smartsheet.Sheets().ListSheets();
            var sheetNameIDList = new List<KeyValuePair<string, long>>();

            foreach (Sheet tmpSheet in homeSheets)
            {
                sheetNameIDList.Add(new KeyValuePair<string, long>((string)tmpSheet.Name, (long)tmpSheet.ID));
            }
            return sheetNameIDList;
        }

        private List<KeyValuePair<string, long>> getColumnList(long sheetID)
        {
            Token token = new Token();
            token.AccessToken = System.Configuration.ConfigurationManager.AppSettings["smartSheetAccessToken"].ToString();

            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();
            IList<Column> sheetCols = smartsheet.Sheets().Columns().ListColumns(sheetID);
            var columnTitleIDList = new List<KeyValuePair<string, long>>();

            foreach (Column tmpCols in sheetCols)
            {
                columnTitleIDList.Add(new KeyValuePair<string, long>((string)tmpCols.Title, (long)tmpCols.ID));
            }
            return columnTitleIDList;
        }

        public string connString { get; set; }

        public string cmd { get; set; }

    }
}
