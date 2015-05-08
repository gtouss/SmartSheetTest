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
            // Use the Smartsheet Builder to create a Smartsheet
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();

            connString = "Server=srv-it3.celerant.com;Database=SupportSSRS;Trusted_Connection=Yes;";
            DataTable table1 = new DataTable();
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                cmd = "select customer,Incident_ID,AssignedTo,cast(StartDate as date),cast(LastActionDate as date),Problem,TeamName from supportssrs.dbo.TB_INCIDENT where StartDate<dateadd(dd,-7,cast(getdate() as date)) and TeamName like '%team%' and enddate is null order by TeamName";
                using (SqlDataAdapter reader = new SqlDataAdapter(cmd, conn))
                {
                    reader.Fill(table1);
                    dataGridView1.DataSource = table1;
                }
            }
            foreach (DataRow dtrow in table1.Rows)
            {
                Cell cellClient = new Cell();
                cellClient.Value = dtrow.ItemArray[0];
                cellClient.Strict = false;
                cellClient.ColumnId = 7656332250638212;
                Cell cellIncidentID = new Cell();
                cellIncidentID.Value = dtrow.ItemArray[1];
                cellIncidentID.ColumnId = 2026832716425092;
                Cell cellAssignedTo = new Cell();
                cellAssignedTo.Value = dtrow.ItemArray[2];
                cellAssignedTo.ColumnId = 1078710095898500;
                Cell cellStartDate = new Cell();
                cellStartDate.Value = dtrow.ItemArray[3];
                cellStartDate.ColumnId = 5582309723268996;
                Cell cellLastActionDate = new Cell();
                cellLastActionDate.Value = dtrow.ItemArray[4];
                cellLastActionDate.ColumnId = 3330509909583748;
                Cell cellProblem = new Cell();
                cellProblem.Value = dtrow.ItemArray[5];
                cellProblem.ColumnId = 7834109536954244;
                Cell cellTeamLead = new Cell();
                cellTeamLead.Value = dtrow.ItemArray[6];
                cellTeamLead.ColumnId = 7706340568131460;

                //// Store the cells in a list
                List<Cell> cells1 = new List<Cell>();
                cells1.Add(cellClient);
                cells1.Add(cellIncidentID);
                cells1.Add(cellAssignedTo);
                cells1.Add(cellStartDate);
                cells1.Add(cellLastActionDate);
                cells1.Add(cellProblem);
                cells1.Add(cellTeamLead);
                //// Create a row and add the list of cells to the row
                Row row = new Row();

                row.Cells = cells1;
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

            // Use the Smartsheet Builder to create a Smartsheet
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();
            IList<Column> sheetCols = smartsheet.Sheets().Columns().ListColumns(1449506131732356);
            foreach (Column tmpCols in sheetCols)
            {

                Console.WriteLine("sheet:" + tmpCols.ID + " " + tmpCols.Title);
                Output.Text = Output.Text + tmpCols.ID + " " + tmpCols.Title + "\r\n";
            }
        }

        public string connString { get; set; }

        public string cmd { get; set; }

    }
}
