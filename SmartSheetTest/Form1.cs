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
using System.Diagnostics;

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
            DataTable table1 = getData("select Customer as Client,Incident_ID as IncidentID,AssignedTo,cast(StartDate as date) as StartDate,cast(LastActionDate as date) as LastAction,Problem,TeamName from supportssrs.dbo.TB_INCIDENT where StartDate<dateadd(dd,-7,cast(getdate() as date)) and TeamName like '%team%' and enddate is null order by TeamName asc, Customer asc");
            populateRowsFromTable(table1);
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

        private void syncButton_Click(object sender, EventArgs e)
        {
            string commaSeparatedIncidents = getIncidentNumbersFromSheet();
            if (commaSeparatedIncidents.Length > 0)
            {
                var newIncidentsTable = getData("select Customer as Client,Incident_ID as IncidentID,AssignedTo,cast(StartDate as date) as StartDate,cast(LastActionDate as date) as LastAction,Problem,TeamName from supportssrs.dbo.TB_INCIDENT where StartDate<dateadd(dd,-7,cast(getdate() as date)) and TeamName like '%team%' and enddate is null and Incident_ID not in (" + commaSeparatedIncidents + ") order by TeamName asc, Customer asc");
                populateRowsFromTable(newIncidentsTable);

                var closedIncidentsTable = getData("select Customer as Client,Incident_ID as IncidentID,AssignedTo,cast(StartDate as date) as StartDate,cast(LastActionDate as date) as LastAction,Problem,TeamName from supportssrs.dbo.TB_INCIDENT where TeamName like '%team%' and enddate is not null and Incident_ID in (" + commaSeparatedIncidents + ") order by TeamName asc, Customer asc");
                deleteRowsFromTable(closedIncidentsTable);                
            }
        }

        private int getColumnIndex(string columnName)
        {
            int columnIndex = -1;
            Token token = new Token();
            token.AccessToken = System.Configuration.ConfigurationManager.AppSettings["smartSheetAccessToken"].ToString();
            var sheetList = getSheetList();
            long sheetID = sheetList.First(kvp => kvp.Key == "Old Incidents").Value;
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();

            IList<Column> sheetCols = smartsheet.Sheets().Columns().ListColumns(sheetID);
            foreach (Column tmpCols in sheetCols)
            {
                if (tmpCols.Title == "IncidentID") columnIndex = (int)tmpCols.Index;
            }
            return columnIndex;
        }

        private void deleteRowsFromTable(DataTable tableData)
        {
            Token token = new Token();
            token.AccessToken = System.Configuration.ConfigurationManager.AppSettings["smartSheetAccessToken"].ToString();
            var sheetList = getSheetList();
            long sheetID = sheetList.First(kvp => kvp.Key == "Old Incidents").Value;
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();
            Sheet sheet = smartsheet.Sheets().GetSheet((long)sheetID, new ObjectInclusion[] { ObjectInclusion.DATA, ObjectInclusion.COLUMNS });

            int columnIndex = getColumnIndex("IncidentID");

            if (columnIndex != -1)
            {
                foreach (Row tmpRow in sheet.Rows)
                {
                    int value;
                    bool isNumeric = int.TryParse(tmpRow.Cells[columnIndex].Value.ToString(), out value);
                    if (isNumeric == true)
                    {
                        int searchIncident = value;
                        bool contains = tableData.AsEnumerable()
                                        .Any(row => searchIncident == row.Field<int>("IncidentID"));
                        if (contains == true) smartsheet.Rows().DeleteRow((long)tmpRow.ID);
                    };
                }
            }
        }

        private string getIncidentNumbersFromSheet()
        {
            string commaSeparatedIncidents = "";
            Token token = new Token();
            token.AccessToken = System.Configuration.ConfigurationManager.AppSettings["smartSheetAccessToken"].ToString();
            var sheetList = getSheetList();
            long sheetID = sheetList.First(kvp => kvp.Key == "Old Incidents").Value;
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();
            Sheet sheet = smartsheet.Sheets().GetSheet((long)sheetID, new ObjectInclusion[] { ObjectInclusion.DATA, ObjectInclusion.COLUMNS });

            int columnIndex = getColumnIndex("IncidentID");

            if (columnIndex != -1)
            {
                foreach (Row tmpRow in sheet.Rows)
                {
                    int value;
                    bool isNumeric = int.TryParse(tmpRow.Cells[columnIndex].Value.ToString(), out value);
                    if (isNumeric == true)
                        commaSeparatedIncidents = commaSeparatedIncidents + value + ",";
                }
                if (commaSeparatedIncidents.Contains(',') == true) commaSeparatedIncidents = commaSeparatedIncidents.TrimEnd(',');
            }
            return commaSeparatedIncidents;
        }

        private DataTable getData(string query)
        {
            connString = "Server=srv-it3.celerant.com;Database=SupportSSRS;Trusted_Connection=Yes;";
            DataTable table1 = new DataTable();
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                using (SqlDataAdapter reader = new SqlDataAdapter(query, conn))
                {
                    reader.Fill(table1);
                    dataGridView1.DataSource = table1;
                }
                conn.Close();
            }
            return table1;
        }

        private void populateRowsFromTable(DataTable tableData)
        {
            Token token = new Token();
            token.AccessToken = System.Configuration.ConfigurationManager.AppSettings["smartSheetAccessToken"].ToString();
            SmartsheetClient smartsheet = new SmartsheetBuilder().SetAccessToken(token.AccessToken).Build();
            var sheetList = getSheetList();
            long sheetID = sheetList.First(kvp => kvp.Key == "Old Incidents").Value;
            var columnList = getColumnList(sheetID);

            foreach (DataRow dtrow in tableData.Rows)
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
                cellLastActionDate.Value = dtrow["LastAction"];
                cellLastActionDate.ColumnId = columnList.First(kvp => kvp.Key == "LastAction").Value;

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
                smartsheet.Sheets().Rows().InsertRows(sheetID, rowWrapper);
            }
        }
    }
}