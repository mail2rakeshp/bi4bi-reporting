using Microsoft.Identity.Client;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using Microsoft.AnalysisServices.AdomdClient;
using System.Data;
using System.Windows.Controls;
using System.Web.UI;
using System.IO;
using System.Text;
using System.Windows.Input;
using System.Diagnostics;
using System.Web;
using System.Web.Security;
using System.Windows.Navigation;
using System.Windows.Threading;
using System.ComponentModel;
using System.Windows.Media;
using System.Threading;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reflection;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using GetMetaData;
using System.Windows.Documents;
using Microsoft.SqlServer;
using Microsoft.SqlServer.Management.Smo;
using static IronPython.Modules.PythonSocket;
using System.Windows.Shapes;
using Microsoft.Scripting.Hosting;
using static Community.CsharpSqlite.Sqlite3;
using IronPython.Modules;
using static IronPython.Modules._ast;
using static IronPython.SQLite.PythonSQLite;


namespace GetMetaData
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>


    public partial class PowerBi : Window
    {

        string[] scopes = new string[] { "user.read" };

        private System.Windows.Forms.NotifyIcon MyNotifyIcon;
        private static string PythonPath1;
        Microsoft.Office.Interop.Excel.Application excel;
        Microsoft.Office.Interop.Excel.Workbook workBook;
        Microsoft.Office.Interop.Excel.Worksheet workSheet;
        Microsoft.Office.Interop.Excel.Range cellRange;
        BackgroundWorker worker;
        string[] items;
        string[] items1;
        string[] items2;
        string[] itemCombo;
        string workspacename = "";
        string workspacename1 = "";
        string serverlabel = "";
        BackgroundWorker backgroundWorker1 = new BackgroundWorker();
        BackgroundWorker backgroundWorker2 = new BackgroundWorker();
        BackgroundWorker backgroundWorker2_1 = new BackgroundWorker();
        BackgroundWorker backgroundWorker3 = new BackgroundWorker();
        DataSet ds;
        DataSet ds1;
        int ReportCnt = 0;
        int ColumnsCnt = 0;
        int CalcCnt = 0;
        string[] selectedItems;

        public CheckBox ReportsCheckbox; //Like here
        List<DDL_Report> objCountryList;
        List<string> listreports = new List<string>();
        List<string> listreports1 = new List<string>();
        List<string> listreports2 = new List<string>();
        DataView view;
        public PowerBi()
        {
            //System.Threading.Thread.Sleep(3000);
            InitializeComponent();
            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            backgroundWorker1.ProgressChanged += backgroundWorker1_ProgressChanged;
            backgroundWorker1.RunWorkerCompleted += backgroundWorker1_RunWorkerCompleted;  //Tell the user how the process went
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;

            backgroundWorker2.DoWork += backgroundWorker2_DoWork;
            backgroundWorker2.ProgressChanged += backgroundWorker2_ProgressChanged;
            backgroundWorker2.RunWorkerCompleted += backgroundWorker2_RunWorkerCompleted;  //Tell the user how the process went
            backgroundWorker2.WorkerReportsProgress = true;
            backgroundWorker2.WorkerSupportsCancellation = true;

            backgroundWorker2_1.DoWork += backgroundWorker2_DoWork;
            backgroundWorker2_1.ProgressChanged += backgroundWorker2_ProgressChanged;
            backgroundWorker2_1.RunWorkerCompleted += backgroundWorker2_RunWorkerCompleted;  //Tell the user how the process went
            backgroundWorker2_1.WorkerReportsProgress = true;
            backgroundWorker2_1.WorkerSupportsCancellation = true;

            backgroundWorker3.DoWork += backgroundWorker3_DoWork;
            backgroundWorker3.ProgressChanged += backgroundWorker3_ProgressChanged;
            backgroundWorker3.RunWorkerCompleted += backgroundWorker3_RunWorkerCompleted;  //Tell the user how the process went
            backgroundWorker3.WorkerReportsProgress = true;
            backgroundWorker3.WorkerSupportsCancellation = true;



            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            MyNotifyIcon.Icon = new System.Drawing.Icon(
                            @"Final.ico");
            MyNotifyIcon.MouseDoubleClick +=
                new System.Windows.Forms.MouseEventHandler(MyNotifyIcon_MouseDoubleClick);
            WrapCheck.Visibility = Visibility.Collapsed;
            BorderBox.Visibility = Visibility.Collapsed;

            Workspace.IsChecked = true;
            DatasetCheck.IsChecked = true;
            ReportsCheck.IsChecked = true;
            ColumnsCheck.IsChecked = true;
            Source.IsChecked = true;
            CalcTablesCheck.IsChecked = true;
            CalcColumnsCheck.IsChecked = true;
            MeasuresCheck.IsChecked = true;
            Relationships.IsChecked = true;
            button1.Visibility = Visibility.Visible;
            ReqButton.Visibility = Visibility.Visible;
            Show_by_Report.Visibility = Visibility.Visible;
            CallGraphButton.Visibility = Visibility.Visible;
            Output.IsEnabled = false;
            ProcessImage.Visibility = Visibility.Hidden;
            OutputImage.Visibility = Visibility.Hidden;
            ImageToolTip.Text = "Fill the mandatory fields in the below sequence : ";
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("1. Workspace Connection -  Premium Workspace Connection from where the Reports can be accessed");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("2. Get Reports - Clicking on this will fetch the list of all reports from the Workspace connection");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("3. Target SQL Server - Server where the Metadata information will be inserted ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("4. Generate Metadata -Start the process for Metadata generation ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("5. Generate Output/Requirement Doc - generate output or requirement document based on the metadata inserted in Step 4 ");

            objCountryList = new List<DDL_Report>();


        }
        private void BindCountryDropDown()
        {
            ComboBoxZone.ItemsSource = items;
        }
        private void ddlCountry_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void ddlCountry_TextChanged(object sender, TextChangedEventArgs e)
        {


            view.RowFilter = string.Format("CATALOG_NAME Like '%{0}%'", ComboBoxZone.Text.ToString());
            if (ComboBoxZone.Text.ToString() != "")
            {
                // view.RowFilter = string.Format("CATALOG_NAME Like '%{0}%'", ComboBoxZone.Text.ToString());
                ComboBoxZone.ItemsSource = view;

                ComboBoxZone.SelectedItem = selectedItems;



            }
            else
            {

                ComboBoxZone.ItemsSource = ds.Tables[0].DefaultView;
            }
        }
        private void WindowMainName_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // popup1.Visibility = Visibility.Collapsed;
            // PopText.Visibility = Visibility.Collapsed;

        }
        private void AllCheckbocx_Checked(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show(((CheckBox)sender).Content.ToString());

            listreports.Add(((CheckBox)sender).Content.ToString());
            listreports1.Add(((CheckBox)sender).Content.ToString());

            // ((CheckBox)sender).IsChecked = true;
            items = listreports.ToArray();
            items1 = listreports1.ToArray();

            selectedItems = items;
            //popup1.Visibility = Visibility.Visible;
            //PopText.Visibility = Visibility.Visible;
            string toDisplay = string.Join(Environment.NewLine, items);
            BorderSelected.Visibility = Visibility.Visible;
            PopText.Visibility = Visibility.Visible;
            LabelSelectedReports.Visibility = Visibility.Visible;
            PopText.Text = toDisplay;

        }

        private void AllCheckbocx_Checked_1(object sender, RoutedEventArgs e)
        {
            listreports.Add(((CheckBox)sender).Content.ToString());
            listreports2.Add(((CheckBox)sender).Content.ToString());
            // ((CheckBox)sender).IsChecked = true;
            items = listreports.ToArray();
            items2 = listreports2.ToArray();
            selectedItems = items;
            //popup1.Visibility = Visibility.Visible;
            //PopText.Visibility = Visibility.Visible;
            string toDisplay = string.Join(Environment.NewLine, items);
            BorderSelected.Visibility = Visibility.Visible;
            PopText.Visibility = Visibility.Visible;
            LabelSelectedReports.Visibility = Visibility.Visible;
            PopText.Text = toDisplay;

        }

        private void AllCheckbocx_Unchecked(object sender, RoutedEventArgs e)
        {

            listreports.Remove(((CheckBox)sender).Content.ToString());
            listreports1.Remove(((CheckBox)sender).Content.ToString());
            items = listreports.ToArray();
            items1 = listreports.ToArray();
            //popup1.Visibility = Visibility.Visible;
            //PopText.Visibility = Visibility.Visible;

            string toDisplay = string.Join(Environment.NewLine, items);
            PopText.Text = toDisplay;
            if (items.Count() > 0)
            {
                BorderSelected.Visibility = Visibility.Visible;
                PopText.Visibility = Visibility.Visible;
                LabelSelectedReports.Visibility = Visibility.Visible;
            }
            else
            {

                BorderSelected.Visibility = Visibility.Collapsed;
                PopText.Visibility = Visibility.Collapsed;
                LabelSelectedReports.Visibility = Visibility.Collapsed;
            }
        }

        private void AllCheckbocx_Unchecked_1(object sender, RoutedEventArgs e)
        {

            listreports.Remove(((CheckBox)sender).Content.ToString());
            listreports2.Remove(((CheckBox)sender).Content.ToString());
            items = listreports.ToArray();
            items2 = listreports.ToArray();
            //popup1.Visibility = Visibility.Visible;
            //PopText.Visibility = Visibility.Visible;

            string toDisplay = string.Join(Environment.NewLine, items);
            PopText.Text = toDisplay;
            if (items.Count() > 0)
            {
                BorderSelected.Visibility = Visibility.Visible;
                PopText.Visibility = Visibility.Visible;
                LabelSelectedReports.Visibility = Visibility.Visible;
            }
            else
            {
                BorderSelected.Visibility = Visibility.Collapsed;
                PopText.Visibility = Visibility.Collapsed;
                LabelSelectedReports.Visibility = Visibility.Collapsed;
            }
        }



        private void Window_Mouse_Double(object sender, RoutedEventArgs e)
        {
            if (this.WindowState == WindowState.Maximized)
            {
                this.WindowState = WindowState.Normal;

            }
            else
            {
                this.WindowState = WindowState.Maximized;


            }
        }
        public string check(string queryString1)
        {

            List<string> conditions = new List<string>();
            if (Workspace.IsChecked == true)
            {
                int pos = ResultText.Text.LastIndexOf("/") + 1;
                conditions.Add(" '" + ResultText.Text.Substring(pos, ResultText.Text.Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]");


            }
            if (ReportsCheck.IsChecked == true)
            {
                conditions.Add(" [CATALOG_NAME] AS [Report Name]");
            }
            if (DatasetCheck.IsChecked == true)
            {
                conditions.Add(" [DIMENSION_UNIQUE_NAME] AS [Dataset Name]");

            }
            if (ColumnsCheck.IsChecked == true)
            {
                conditions.Add(" LEVEL_CAPTION AS [Column Name]");

            }
            if (conditions.Any() && (ColumnsCheck.IsChecked == true || CalcColumnsCheck.IsChecked == true || CalcTablesCheck.IsChecked == true || MeasuresCheck.IsChecked == true))
            {

                queryString1 = "SELECT DISTINCT" + string.Join(",", conditions.ToArray()) + " FROM $System.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 AND LEVEL_NAME <> '(All)' order by [DIMENSION_UNIQUE_NAME]  ";
            }
            else if (Workspace.IsChecked == true || ReportsCheck.IsChecked == true || DatasetCheck.IsChecked == true && (ColumnsCheck.IsChecked == false && DatasetCheck.IsChecked == false && CalcColumnsCheck.IsChecked == false && CalcTablesCheck.IsChecked == false && MeasuresCheck.IsChecked == false))
            {
                queryString1 = "SELECT DISTINCT" + string.Join(",", conditions.ToArray()) + " FROM $System.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 order by [DIMENSION_UNIQUE_NAME]  ";
            }
            else
            {
                MessageBox.Show("Please Choose any items from the list to view");
                //int pos = ResultText.Text.LastIndexOf("/") + 1;
                //WorkspaceLabel.Content = "'" + ResultText.Text.Substring(pos, ResultText.Text.Length - pos).Replace("%20", " ") + "' AS [Workspace]";
                // queryString1 = "SELECT "+ WorkspaceLabel.Content.ToString()+ ",[CATALOG_NAME] as [Report Name],[DIMENSION_UNIQUE_NAME] AS [Dataset Name],LEVEL_CAPTION AS [Column Name] FROM $system.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 AND LEVEL_NAME <> '(All)' order by [DIMENSION_UNIQUE_NAME] ";

            }

            return queryString1;



        }

        public string check2(string queryString1)
        {

            List<string> conditions = new List<string>();
            if (Workspace.IsChecked == true)
            {
                int pos = ResultText2.Text.LastIndexOf("/") + 1;
                conditions.Add(" '" + ResultText2.Text.Substring(pos, ResultText2.Text.Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]");


            }
            if (ReportsCheck.IsChecked == true)
            {
                conditions.Add(" [CATALOG_NAME] AS [Report Name]");
            }
            if (DatasetCheck.IsChecked == true)
            {
                conditions.Add(" [DIMENSION_UNIQUE_NAME] AS [Dataset Name]");

            }
            if (ColumnsCheck.IsChecked == true)
            {
                conditions.Add(" LEVEL_CAPTION AS [Column Name]");

            }
            if (conditions.Any() && (ColumnsCheck.IsChecked == true || CalcColumnsCheck.IsChecked == true || CalcTablesCheck.IsChecked == true || MeasuresCheck.IsChecked == true))
            {

                queryString1 = "SELECT DISTINCT" + string.Join(",", conditions.ToArray()) + " FROM $System.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 AND LEVEL_NAME <> '(All)' order by [DIMENSION_UNIQUE_NAME]  ";
            }
            else if (Workspace.IsChecked == true || ReportsCheck.IsChecked == true || DatasetCheck.IsChecked == true && (ColumnsCheck.IsChecked == false && DatasetCheck.IsChecked == false && CalcColumnsCheck.IsChecked == false && CalcTablesCheck.IsChecked == false && MeasuresCheck.IsChecked == false))
            {
                queryString1 = "SELECT DISTINCT" + string.Join(",", conditions.ToArray()) + " FROM $System.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 order by [DIMENSION_UNIQUE_NAME]  ";
            }
            else
            {
                MessageBox.Show("Please Choose any items from the list to view");
                //int pos = ResultText.Text.LastIndexOf("/") + 1;
                //WorkspaceLabel.Content = "'" + ResultText.Text.Substring(pos, ResultText.Text.Length - pos).Replace("%20", " ") + "' AS [Workspace]";
                // queryString1 = "SELECT "+ WorkspaceLabel.Content.ToString()+ ",[CATALOG_NAME] as [Report Name],[DIMENSION_UNIQUE_NAME] AS [Dataset Name],LEVEL_CAPTION AS [Column Name] FROM $system.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 AND LEVEL_NAME <> '(All)' order by [DIMENSION_UNIQUE_NAME] ";

            }

            return queryString1;



        }

        public string check3(string queryString1)
        {

            List<string> conditions = new List<string>();
            if (Workspace.IsChecked == true)
            {
                int pos = ResultText3.Text.LastIndexOf("/") + 1;
                conditions.Add(" '" + ResultText3.Text.Substring(pos, ResultText3.Text.Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]");


            }
            if (ReportsCheck.IsChecked == true)
            {
                conditions.Add(" [CATALOG_NAME] AS [Report Name]");
            }
            if (DatasetCheck.IsChecked == true)
            {
                conditions.Add(" [DIMENSION_UNIQUE_NAME] AS [Dataset Name]");

            }
            if (ColumnsCheck.IsChecked == true)
            {
                conditions.Add(" LEVEL_CAPTION AS [Column Name]");

            }
            if (conditions.Any() && (ColumnsCheck.IsChecked == true || CalcColumnsCheck.IsChecked == true || CalcTablesCheck.IsChecked == true || MeasuresCheck.IsChecked == true))
            {

                queryString1 = "SELECT DISTINCT" + string.Join(",", conditions.ToArray()) + " FROM $System.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 AND LEVEL_NAME <> '(All)' order by [DIMENSION_UNIQUE_NAME]  ";
            }
            else if (Workspace.IsChecked == true || ReportsCheck.IsChecked == true || DatasetCheck.IsChecked == true && (ColumnsCheck.IsChecked == false && DatasetCheck.IsChecked == false && CalcColumnsCheck.IsChecked == false && CalcTablesCheck.IsChecked == false && MeasuresCheck.IsChecked == false))
            {
                queryString1 = "SELECT DISTINCT" + string.Join(",", conditions.ToArray()) + " FROM $System.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 order by [DIMENSION_UNIQUE_NAME]  ";
            }
            else
            {
                MessageBox.Show("Please Choose any items from the list to view");
                //int pos = ResultText.Text.LastIndexOf("/") + 1;
                //WorkspaceLabel.Content = "'" + ResultText.Text.Substring(pos, ResultText.Text.Length - pos).Replace("%20", " ") + "' AS [Workspace]";
                // queryString1 = "SELECT "+ WorkspaceLabel.Content.ToString()+ ",[CATALOG_NAME] as [Report Name],[DIMENSION_UNIQUE_NAME] AS [Dataset Name],LEVEL_CAPTION AS [Column Name] FROM $system.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 AND LEVEL_NAME <> '(All)' order by [DIMENSION_UNIQUE_NAME] ";

            }

            return queryString1;



        }



        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            // This is where the processor intensive code should go
            ExecuteMethodAsync();

            // if we need any output to be used, put it in the DoWorkEventArgs object
            e.Result = "all done";
            //If the process exits the loop, ensure that progress is set to 100%
            //Remember in the loop we set i < 100 so in theory the process will complete at 99%         
            backgroundWorker1.ReportProgress(100);

        }
        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {

            }
            else if (e.Error != null)
            {

            }
            else
            {
                if (ReportCnt > 0 || ColumnsCnt > 0 || CalcCnt > 0)
                {

                    Animation.Visibility = Visibility.Collapsed;
                    ServerStack.Visibility = Visibility.Visible;
                    button1.Visibility = Visibility.Visible;
                    ReqButton.Visibility = Visibility.Visible;
                    Show_by_Report.Visibility = Visibility.Visible;
                    CallGraphButton.Visibility = Visibility.Visible;
                    SignOutButton.Visibility = Visibility.Visible;
                    GenerateMetadata.Visibility = Visibility.Collapsed;
                    Output.Visibility = Visibility.Collapsed;
                    Output.IsEnabled = true;
                    GenerateMetadata.IsEnabled = false;
                    Output.IsChecked = true;
                    ProcessImage.Visibility = Visibility.Hidden;
                    OutputImage.Visibility = Visibility.Hidden;
                    //StackGrid.Visibility = Visibility.Hidden;
                    //dataGrid1.Visibility = Visibility.Collapsed;

                    MetadataToolTip.Text = "Please Find the summary of items inserted into the server " + serverlabel.ToString();
                    MetadataToolTip.AppendText(Environment.NewLine);
                    MetadataToolTip.AppendText("Number of Reports = " + ReportCnt + "\r\n");
                    MetadataToolTip.AppendText("Number of Columns = " + ColumnsCnt + "\r\n");
                    MetadataToolTip.AppendText("Number of Calculations = " + CalcCnt + "\r\n");

                    OutputToolTip.Text = "Generate Power BI Report - The generated metadata is presented in a read-able format in a Power BI Report\r\n";
                    OutputToolTip.AppendText("Requirement Document Generator - Generate Requirement Document for easier hand-over which will help in migration");


                }
                else
                {

                    Animation.Visibility = Visibility.Collapsed;
                    ServerStack.Visibility = Visibility.Visible;
                    button1.Visibility = Visibility.Visible;
                    ReqButton.Visibility = Visibility.Visible;
                    Show_by_Report.Visibility = Visibility.Visible;
                    GenerateMetadata.Visibility = Visibility.Visible;
                    Output.Visibility = Visibility.Collapsed;
                    Output.IsEnabled = false;
                    //   GenerateMetadata.IsEnabled = true;
                    //  GenerateMetadata.IsChecked = true;

                    MessageBox.Show("Issues Found in the Metadata Process. Please contact the administrator for further clarification");

                }
            }
        }
        private async void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {


            items = GetArrayfromCombo();
            if (LabelSelectedReports.IsVisible == false)
            {
                items1 = GetArrayfromCombo();
                items2 = GetArrayfromCombo1();
            }
            workspacename = ResultText.Text.ToString();
            serverlabel = Server.Text.ToString();
            if (String.IsNullOrEmpty(serverlabel.ToString()))
            {


                MessageBox.Show("Enter the mandatory fields");
                Animation.Visibility = Visibility.Collapsed;
                ServerStack.Visibility = Visibility.Visible;
                button1.Visibility = Visibility.Visible;
                ReqButton.Visibility = Visibility.Visible;
                Show_by_Report.Visibility = Visibility.Visible;
                CallGraphButton.Visibility = Visibility.Visible;
                SignOutButton.Visibility = Visibility.Visible;
                GenerateMetadata.Visibility = Visibility.Collapsed;
                Output.Visibility = Visibility.Collapsed;
                Output.IsEnabled = false;
                // GenerateMetadata.IsEnabled = true;
                //  GenerateMetadata.IsChecked = false;

            }
            else
            {
                Animation.Visibility = Visibility.Visible;
                ServerStack.Visibility = Visibility.Hidden;
                button1.Visibility = Visibility.Collapsed;
                ReqButton.Visibility = Visibility.Collapsed;
                Show_by_Report.Visibility = Visibility.Collapsed;
                CallGraphButton.Visibility = Visibility.Collapsed;
                SignOutButton.Visibility = Visibility.Collapsed;
                GenerateMetadata.Visibility = Visibility.Collapsed;
                Output.Visibility = Visibility.Collapsed;
                ComboBoxZone.Text = "";
                WindowMainName.Height = 766;
                backgroundWorker1.RunWorkerAsync();

            }

        }




        public void ExecuteMethodAsync()
        {



            string authResult = null;
            DisplayBasicTokenInfo(authResult);

            SqlConnection SQLConnection = new SqlConnection();
            SQLConnection.ConnectionString = "Data Source =" + serverlabel.ToString() + "; Initial Catalog =Power BI Metadata; " + "Integrated Security=true;";

            string QueryReport = "select count(DISTINCT [Report Name]) from dbo.Metadata";
            //Execute Queries and save results into variables
            SqlCommand CmdCntReport = SQLConnection.CreateCommand();
            CmdCntReport.CommandText = QueryReport;
            SQLConnection.Open();
            ReportCnt = (Int32)CmdCntReport.ExecuteScalar();
            SQLConnection.Close();


            string QueryColumns = "select count(DISTINCT [Column Name]) from dbo.Metadata";
            //Execute Queries and save results into variables
            SqlCommand CmdCntColumns = SQLConnection.CreateCommand();
            CmdCntColumns.CommandText = QueryColumns;
            SQLConnection.Open();

            ColumnsCnt = (Int32)CmdCntColumns.ExecuteScalar();
            SQLConnection.Close();

            string QueryCalc = "SELECT SUM([Calc 1]) FROM ";
            QueryCalc += "\n (";
            QueryCalc += "\n select COUNT(DISTINCT [Calculated Column Expression]) [Calc 1] from dbo.Metadata";
            QueryCalc += "\n where [Calculated Column Expression] is not null";
            QueryCalc += "\n UNION ALL ";
            QueryCalc += "\n select COUNT(DISTINCT [Calculated Measure Expression]) [Calc 2] from dbo.Metadata";
            QueryCalc += "\n where [Calculated Measure Expression] is not null";
            QueryCalc += "\n UNION ALL";
            QueryCalc += "\n select COUNT(DISTINCT [Calculated Table Expression]) [Calc 3] from dbo.Metadata";
            QueryCalc += "\n where [Calculated Table Expression] is not null";
            QueryCalc += "\n ) A";
            //Execute Queries and save results into variables
            SqlCommand CmdCntCalc = SQLConnection.CreateCommand();
            CmdCntCalc.CommandText = QueryCalc;
            SQLConnection.Open();
            CalcCnt = (Int32)CmdCntCalc.ExecuteScalar();
            SQLConnection.Close();



            try
            {

                SQLConnection.Open();

                string script = File.ReadAllText(System.IO.Path.Combine(Environment.CurrentDirectory, @"Scripts\", "vw_Metadata.sql"));

                // split script on GO command
                IEnumerable<string> commandStrings = Regex.Split(script, @"^\s*GO\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);
                foreach (string commandString in commandStrings)
                {
                    if (commandString.Trim() != "")
                    {
                        new SqlCommand(commandString, SQLConnection).ExecuteNonQuery();
                    }
                }

                script = File.ReadAllText(System.IO.Path.Combine(Environment.CurrentDirectory, @"Scripts\", "vw_Metadata_Calculations.sql"));

                // split script on GO command
                commandStrings = Regex.Split(script, @"^\s*GO\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);
                foreach (string commandString in commandStrings)
                {
                    if (commandString.Trim() != "")
                    {
                        new SqlCommand(commandString, SQLConnection).ExecuteNonQuery();
                    }
                }

                script = File.ReadAllText(System.IO.Path.Combine(Environment.CurrentDirectory, @"Scripts\", "vw_Metadata_Columns.sql"));

                // split script on GO command
                commandStrings = Regex.Split(script, @"^\s*GO\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);
                foreach (string commandString in commandStrings)
                {
                    if (commandString.Trim() != "")
                    {
                        new SqlCommand(commandString, SQLConnection).ExecuteNonQuery();
                    }
                }

                script = File.ReadAllText(System.IO.Path.Combine(Environment.CurrentDirectory, @"Scripts\", "vw_Metadata_STM.sql"));

                // split script on GO command
                commandStrings = Regex.Split(script, @"^\s*GO\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);
                foreach (string commandString in commandStrings)
                {
                    if (commandString.Trim() != "")
                    {
                        new SqlCommand(commandString, SQLConnection).ExecuteNonQuery();
                    }
                }



                string scriptp = "\nimport urllib";
                scriptp += "\nimport pandas as pd";
                scriptp += "\nimport numpy as np";
                scriptp += "\nfrom sqlalchemy import create_engine";
                scriptp += "\nimport pyodbc";
                scriptp += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + serverlabel.ToString() + ";DATABASE=Power BI Metadata;Trusted_Connection=yes;\")";
                scriptp += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted), fast_executemany=True)";
                scriptp += "\nconn_str = (\"DRIVER={SQL Server Native Client 11.0};SERVER=" + serverlabel.ToString() + ";DATABASE=Power BI Metadata;Trusted_Connection=yes;\")   ";
                scriptp += "\ncnxn = pyodbc.connect(conn_str)";
                scriptp += "\ncursor = cnxn.cursor()";
                scriptp += "\ntable_names = []";
                scriptp += "\nfor row in cursor.tables():";
                scriptp += "\n  if row.table_type == 'TABLE':";
                scriptp += "\n      table_names.append(row.table_name)";
                scriptp += "\nif 'trace_xe_action_map' in table_names: table_names.remove('trace_xe_action_map')";
                scriptp += "\nif 'trace_xe_event_map' in table_names: table_names.remove('trace_xe_event_map')";
                scriptp += "\nfor each_table in table_names:";
                scriptp += "\n    df = pd.read_sql('select * from ' + each_table, cnxn)";
                scriptp += "\n    df = df.drop_duplicates()";
                scriptp += "\n    df.to_sql(each_table, schema = 'dbo', if_exists = 'replace', con = engine, index = False)";
                scriptp += "\ndata = pd.read_sql('''";
                scriptp += "\nselect metadata.*,[Column Name] AS [Actual Source Column],";
                scriptp += "\n CASE WHEN metadata.[Dataset Name]='Internal Date Table' AND metadata.[Calculated Tables Expression]='Yes' THEN 'Yes'";
                scriptp += "\n  WHEN metadata.[Dataset Name]='Internal Date Table' AND metadata.[Calculated Tables Expression]='No' THEN 'No'";
                scriptp += "\n  ELSE 'Yes' END AS [Flag_Inclusion]";
                scriptp += "\n  from ";
                scriptp += "\n(SELECT [Workspace]";
                scriptp += "\n      ,[Report Name]";
                scriptp += "\n      ,CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' ";
                scriptp += "\n	  WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ";
                scriptp += "\n	  ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END AS [Dataset Name]";
                scriptp += "\n      ,[Column Name]";
                scriptp += "\n      ,[Calculated Column Expression]";
                scriptp += "\n      ,[Calculated Measure Expression]";
                scriptp += "\n      ,[Calculated Table Expression]";
                scriptp += "\n      ,[Source]";
                scriptp += "\n      ,[Database Or Path]";
                scriptp += "\n      ,[Query]";
                scriptp += "\n      ,[DatabaseItem]";
                scriptp += "\n	  ,CASE WHEN [From Table Name] like '%LocalDateTable%' THEN 'Internal Date Table' ";
                scriptp += "\n	  WHEN [From Table Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template'";
                scriptp += "\n	  ELSE [From Table Name] ";
                scriptp += "\n	  END AS [From Table Name]";
                scriptp += "\n	  ,CASE WHEN [To Table Name] like '%LocalDateTable%' THEN 'Internal Date Table' ";
                scriptp += "\n	  WHEN [To Table Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ";
                scriptp += "\n	  ELSE [To Table Name] ";
                scriptp += "\n	  END AS [To Table Name]";
                scriptp += "\n      ,[From Column Name]";
                scriptp += "\n      ,[To Column Name]";
                scriptp += "\n      ,[Refreshed Time]";
                scriptp += "\n	  , CASE WHEN [From Table Name] IS NULL THEN 'No' ELSE 'Yes' END AS [Relationships Set]";
                scriptp += "\n	  , CASE WHEN [Calculated Column Expression] IS NULL THEN 'No' ELSE 'Yes' END AS [Calculated Columns Expressions]";
                scriptp += "\n	  , CASE WHEN [Calculated Measure Expression] IS NULL THEN 'No' ELSE 'Yes' END AS [Calculated Measures Expressions]";
                scriptp += "\n	  , CASE WHEN [Calculated Table Expression] IS NULL THEN 'No' ELSE 'Yes' END AS [Calculated Tables Expression]";
                scriptp += "\n	  ,CONCAT([Workspace]";
                scriptp += "\n      ,[Report Name]";
                scriptp += "\n      ,CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' ";
                scriptp += "\n	  WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ";
                scriptp += "\n	  ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END";
                scriptp += "\n      ,[Column Name]) JOIN_WITH_Dictionary";
                //  scriptp += "\n	  --,CHARINDEX('Table.RenameColumns',[Steps])";
                //  scriptp += "\n--	  ,CASE WHEN CHARINDEX('Table.RenameColumns',[Steps])>0  and ";
                //  scriptp += "\n--	  CHARINDEX('"'+[Column Name]+'"',Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                //  scriptp += "\n--, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                //  scriptp += "\n--,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                //  scriptp += "\n--, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1) )>0  THEN 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                //  scriptp += "\n--	  ELSE NULL END";
                scriptp += "\n	  ,CASE WHEN CHARINDEX('Table.RenameColumns',[Steps])>0  and CHARINDEX('\"'+[Column Name]+'\"', CASE WHEN CHARINDEX('Table.RenameColumns',query)=0 AND CHARINDEX('Table.RenameColumns',[Steps])>0 THEN";
                scriptp += "\n	  Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1) END )>0  ";
                scriptp += "\nand [Calculated Column Expression] is null and [Calculated Measure Expression] is  null and [Calculated Table Expression] is  null and CHARINDEX('Table.RenameColumns',query)=0";
                scriptp += "\nTHEN 	  ";
                scriptp += "\nREVERSE(Substring(REVERSE(SUBSTRING(Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1)";
                scriptp += "\n, 0, charindex('\"'+[Column Name]+ '\"}', Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1), 0)))";
                scriptp += "\n, charindex('\"', REVERSE(SUBSTRING(Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1)";
                scriptp += "\n, 0, charindex('\"'+[Column Name]+ '\"}', Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1), 0)))";
                scriptp += "\n,charindex('\"', REVERSE(SUBSTRING(Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1)";
                scriptp += "\n, 0, charindex('\"'+[Column Name]+ '\"}', Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1), 0)))";
                scriptp += "\n,charindex('\"', REVERSE(SUBSTRING(Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1)";
                scriptp += "\n, 0, charindex('\"'+[Column Name]+ '\"}', Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1), 0)))";
                scriptp += "\n)))+1";
                scriptp += "\n, charindex('\"', REVERSE(SUBSTRING(Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1)";
                scriptp += "\n, 0, charindex('\"'+[Column Name]+ '\"}', Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1), 0)))";
                scriptp += "\n,charindex('\"', REVERSE(SUBSTRING(Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1)";
                scriptp += "\n, 0, charindex('\"'+[Column Name]+ '\"}', Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1), 0)))";
                scriptp += "\n,charindex('\"', REVERSE(SUBSTRING(Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1)";
                scriptp += "\n, 0, charindex('\"'+[Column Name]+ '\"}', Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1), 0)))";
                scriptp += "\n,charindex('\"', REVERSE(SUBSTRING(Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1)";
                scriptp += "\n, 0, charindex('\"'+[Column Name]+ '\"}', Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1), 0)))";
                scriptp += "\n)))+1)-charindex('\"', REVERSE(SUBSTRING(Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1)";
                scriptp += "\n, 0, charindex('\"'+[Column Name]+ '\"}', Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1), 0)))";
                scriptp += "\n,charindex('\"', REVERSE(SUBSTRING(Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1)";
                scriptp += "\n, 0, charindex('\"'+[Column Name]+ '\"}', Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1), 0)))";
                scriptp += "\n,charindex('\"', REVERSE(SUBSTRING(Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1)";
                scriptp += "\n, 0, charindex('\"'+[Column Name]+ '\"}', Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1), 0))))))-1))";
                scriptp += "\nELSE [Column Name]";
                scriptp += "\nEND [Source Column]";
                scriptp += "\n,CASE WHEN CHARINDEX('Table.RenameColumns',[Steps])>0  and CHARINDEX('\"'+[Column Name]+'\"', CASE WHEN CHARINDEX('Table.RenameColumns',query)=0 AND CHARINDEX('Table.RenameColumns',[Steps])>0 THEN";
                scriptp += "\n	  Substring('Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ),charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('Table.RenameColumns', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))";
                scriptp += "\n,charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) )";
                scriptp += "\n, (charindex('})', 'Table.RenameColumns'+SUBSTRING(Steps,charindex('Table.Renamecolumns',Steps) + LEN('Table.Renamecolumns'), LEN(Steps) ), 1)))-1) END )>0  ";
                scriptp += "\nand [Calculated Column Expression] is null and [Calculated Measure Expression] is  null and [Calculated Table Expression] is  null and CHARINDEX('Table.RenameColumns',query)=0";
                scriptp += "\nTHEN 'Yes' ELSE 'No' END AS [Change in Columns]";
                scriptp += "\n  ,[Steps]";
                scriptp += "\n--into #Metadata";
                scriptp += "\n  FROM [Power BI Metadata].[dbo].[Metadata] (NOLOCK)";
                scriptp += "\nWHERE ";
                scriptp += "\n CHARINDEX('{\"'+[Column Name]+'\"}',[Steps])=0";
                scriptp += "\n UNION ALL";
                scriptp += "\n SELECT [Workspace]";
                scriptp += "\n      ,[Report Name]";
                scriptp += "\n      ,CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' ";
                scriptp += "\n	  WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ";
                scriptp += "\n	  ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END AS [Dataset Name]";
                scriptp += "\n      ,[Column Name]";
                scriptp += "\n      ,[Calculated Column Expression]";
                scriptp += "\n      ,[Calculated Measure Expression]";
                scriptp += "\n      ,[Calculated Table Expression]";
                scriptp += "\n      ,[Source]";
                scriptp += "\n      ,[Database Or Path]";
                scriptp += "\n      ,[Query]";
                scriptp += "\n      ,[DatabaseItem]";
                scriptp += "\n	  ,CASE WHEN [From Table Name] like '%LocalDateTable%' THEN 'Internal Date Table' ";
                scriptp += "\n	  WHEN [From Table Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template'";
                scriptp += "\n	  ELSE [From Table Name] ";
                scriptp += "\n	  END AS [From Table Name]";
                scriptp += "\n	  ,CASE WHEN [To Table Name] like '%LocalDateTable%' THEN 'Internal Date Table' ";
                scriptp += "\n	  WHEN [To Table Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ";
                scriptp += "\n	  ELSE [To Table Name] ";
                scriptp += "\n	  END AS [To Table Name]";
                scriptp += "\n      ,[From Column Name]";
                scriptp += "\n      ,[To Column Name]";
                scriptp += "\n      ,[Refreshed Time]";
                scriptp += "\n	  , CASE WHEN [From Table Name] IS NULL THEN 'No' ELSE 'Yes' END AS [Relationships Set]";
                scriptp += "\n	  , CASE WHEN [Calculated Column Expression] IS NULL THEN 'No' ELSE 'Yes' END AS [Calculated Columns Expressions]";
                scriptp += "\n	  , CASE WHEN [Calculated Measure Expression] IS NULL THEN 'No' ELSE 'Yes' END AS [Calculated Measures Expressions]";
                scriptp += "\n	  , CASE WHEN [Calculated Table Expression] IS NULL THEN 'No' ELSE 'Yes' END AS [Calculated Tables Expression]";
                scriptp += "\n	  ,CONCAT([Workspace]";
                scriptp += "\n      ,[Report Name]";
                scriptp += "\n      ,CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' ";
                scriptp += "\n	  WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ";
                scriptp += "\n	  ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END";
                scriptp += "\n      ,[Column Name]) JOIN_WITH_Dictionary";
                scriptp += "\n	, [Column Name]  AS [Source Column]";
                scriptp += "\n , 'No' [Change in Columns] ";
                scriptp += "\n  ,[Steps]";
                scriptp += "\n  FROM [Power BI Metadata].[dbo].[Metadata] (NOLOCK)";
                scriptp += "\nWHERE ";
                scriptp += "\n CHARINDEX('{\"'+[Column Name]+'\"}',[Steps])<>0";
                scriptp += "\n )  as metadata";
                scriptp += "\n''', cnxn)";
                scriptp += "\nworkspaces = data['Workspace'].unique().tolist()";
                scriptp += "\nreport_name = data['Report Name'].unique().tolist()";
                scriptp += "\nresult = []";
                scriptp += "\nfor each_workspace in workspaces:";
                scriptp += "\n    for each_report_name in report_name:";
                scriptp += "\n        inner_column_name = []";
                scriptp += "\n        inner_cce = []";
                scriptp += "\n        inner_cme = []";
                scriptp += "\n        inner_cte = []";
                scriptp += "\n        inner_source = []";
                scriptp += "\n        inner_dop = []";
                scriptp += "\n        inner_query = []";
                scriptp += "\n        temp_data = data[(data['Workspace'] == each_workspace) & (data['Report Name'] == each_report_name)]";
                scriptp += "\n        for _, each_row in temp_data.iterrows():";
                scriptp += "\n            inner_column_name.append(each_row['Column Name'])";
                scriptp += "\n            inner_cce.append(each_row['Calculated Column Expression'])";
                scriptp += "\n            inner_cme.append(each_row['Calculated Measure Expression'])";
                scriptp += "\n            inner_cte.append(each_row['Calculated Table Expression'])";
                scriptp += "\n            inner_source.append(each_row['Source'])";
                scriptp += "\n            inner_dop.append(each_row['Database Or Path'])";
                scriptp += "\n            inner_query.append(each_row['Query'])";
                scriptp += "\n        inner_column_name = [i for i in inner_column_name if i is not None]";
                scriptp += "\n        inner_cce = [i for i in inner_cce if i is not None]";
                scriptp += "\n        inner_cme = [i for i in inner_cme if i is not None]";
                scriptp += "\n        inner_cte = [i for i in inner_cte if i is not None]";
                scriptp += "\n        inner_source = [i for i in inner_source if i is not None]";
                scriptp += "\n        inner_dop = [i for i in inner_dop if i is not None]";
                scriptp += "\n        inner_query = [i for i in inner_query if i is not None]";
                scriptp += "\n        result.append([each_workspace, each_report_name, inner_column_name,";
                scriptp += "\n                       inner_cce, inner_cme, inner_cte, inner_source, inner_dop, inner_query])";
                scriptp += "\npercentage_table = []";
                scriptp += "\nfor i in range(0, len(result)):";
                scriptp += "\n    for j in range(i+1, len(result)):";
                scriptp += "\n        column_per = 0";
                scriptp += "\n        cce_per = 0";
                scriptp += "\n        cme_per = 0";
                scriptp += "\n        cte_per = 0";
                scriptp += "\n        source_per = 0";
                scriptp += "\n        dop_per = 0";
                scriptp += "\n        query_per = 0";
                scriptp += "\n        if result[i][0] == result[j][0]:";
                scriptp += "\n            if result[i][2] == [] and result[j][2] == []:";
                scriptp += "\n                column_per = np.nan";
                scriptp += "\n            elif result[i][2] != [] or result[j][2] != []:";
                scriptp += "\n                column_per = len(set(result[i][2]).intersection(set(result[j][2]))) / float(len(set(result[i][2] + result[j][2]))) * 100";
                scriptp += "\n            if result[i][3] == [] and result[j][3] == []:";
                scriptp += "\n                cce_per = np.nan";
                scriptp += "\n            elif result[i][3] != [] or result[j][3] != []:";
                scriptp += "\n                cce_per = len(set(result[i][3]).intersection(set(result[j][3]))) / float(len(set(result[i][3] + result[j][3]))) * 100";
                scriptp += "\n            if result[i][4] == [] and result[j][4] == []:";
                scriptp += "\n                cme_per = np.nan";
                scriptp += "\n            elif result[i][4] != [] or result[j][4] != []:";
                scriptp += "\n                cme_per = len(set(result[i][4]).intersection(set(result[j][4]))) / float(len(set(result[i][4] + result[j][4]))) * 100";
                scriptp += "\n            if result[i][5] == [] and result[j][5] == []:";
                scriptp += "\n                cte_per = np.nan";
                scriptp += "\n            elif result[i][5] != [] or result[j][5] != []:";
                scriptp += "\n                cte_per = len(set(result[i][5]).intersection(set(result[j][5]))) / float(len(set(result[i][5] + result[j][5]))) * 100";
                scriptp += "\n            if result[i][6] == [] and result[j][6] == []:";
                scriptp += "\n                source_per = np.nan";
                scriptp += "\n            elif result[i][6] != [] or result[j][6] != []:";
                scriptp += "\n                source_per = len(set(result[i][6]).intersection(set(result[j][6]))) / float(len(set(result[i][6] + result[j][6]))) * 100";
                scriptp += "\n            if result[i][7] == [] and result[j][7] == []:";
                scriptp += "\n                dop_per = np.nan";
                scriptp += "\n            elif result[i][7] != [] or result[j][7] != []:";
                scriptp += "\n                dop_per = len(set(result[i][7]).intersection(set(result[j][7]))) / float(len(set(result[i][7] + result[j][7]))) * 100";
                scriptp += "\n            if result[i][8] == [] and result[j][8] == []:";
                scriptp += "\n                query_per = np.nan";
                scriptp += "\n            elif set(result[i][8]) == {'No Query Available'} and set(result[j][8]) == {'No Query Available'}:";
                scriptp += "\n                query_per = np.nan";
                scriptp += "\n            elif result[i][8] != [] or result[j][8] != []:";
                scriptp += "\n                query_per = len(set(result[i][8]).intersection(set(result[j][8]))) / float(len(set(result[i][8] + result[j][8]))) * 100";
                scriptp += "\n            percentage_table.append([result[i][0], result[i][1], result[j][1], column_per, cce_per, cme_per, cte_per, source_per, dop_per, query_per])";
                //     scriptp += "\n percentage_table = []";

                scriptp += "\npercentage_table_df = pd.DataFrame(percentage_table, columns = ['Workspace', 'Report A', 'Report B', 'Column Name', ";
                scriptp += "\n                                                                'Calculated Column Expression',";
                scriptp += "\n                                                                'Calculated Measure Expression', 'Calculated Table Expression',";
                scriptp += "\n                                                                'Source', 'Database Or Path', 'Query'])";

                scriptp += "\npercentage_table_df.to_sql('power_bi_report_match_percentage', schema='dbo',if_exists = 'replace', con = engine, index=False, index_label='myField')";

                string pathp = Directory.GetCurrentDirectory() + @"\PythonFile\PowerBI_Process_Python.py";

                File.SetAttributes(pathp, FileAttributes.Normal);
                if (File.Exists(pathp))
                {
                    File.Delete(pathp);
                }
                using (StreamWriter writer = File.CreateText(pathp))
                {
                    writer.WriteLine(scriptp);

                }
                run_cmd();

                /* try
                 {
                     string workingDirectory = Directory.GetCurrentDirectory() + @"\PythonFile";
                     var process = new Process
                     {
                         StartInfo = new ProcessStartInfo
                         {
                             FileName = "cmd.exe",
                             RedirectStandardInput = true,
                             UseShellExecute = false,
                             RedirectStandardError = true,
                             CreateNoWindow = true,
                             WorkingDirectory = workingDirectory
                         }


                     };
                     process.Start();


                     using (var sw = process.StandardInput)
                     {
                         if (sw.BaseStream.CanWrite)
                         {
                            // string TextPython = "C:\\Users\\Rakesh.P\\Anaconda3";
                             // Batch script to activate Anaconda
                             //  sw.WriteLine(TextPython.Text + @"\Scripts\activate.bat");
                            // sw.WriteLine("C:\\Users\\Rakesh.P\\Anaconda3\\Scripts\activate.bat");
                             sw.WriteLine(@"C:\Users\Rakesh.P\Anaconda3\Scripts\activate.bat");
                             // Activate your environment
                             // sw.WriteLine("conda activate py3.9.7");
                             // run your script. You can also pass in arguments
                             sw.WriteLine("PowerBI_Process_Python.py");
                         }
                     }
                     //string output = process.StandardOutput.ReadToEnd();

                     process.WaitForExit();
                 }
                 catch (Exception ex)
                 {
                     MessageBox.Show(ex.Message.ToString());
                 } */

            }
            catch (SqlException er)
            {

            }
            finally
            {
                SQLConnection.Close();

            }



        }

        public async Task<string> GetHttpContentWithToken(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        /// <summary>
        /// Sign out the current user
        /// </summary>
        private async void SignOutButton_Click(object sender, RoutedEventArgs e)
        {

            ResultText.Text = "";
            ResultText2.Text = "";
            Server.Text = "";
            TextPython.Text = "";

            //  GenerateMetadata.IsChecked = false;
            Output.IsChecked = false;
            Output.IsEnabled = false;
            GenerateMetadata.IsEnabled = true;

            ProcessImage.Visibility = Visibility.Hidden;
            OutputImage.Visibility = Visibility.Hidden;
            ComboBoxZone.ItemsSource = null;
            ComboBoxZone.Items.Clear();
            button1.Visibility = Visibility.Visible;
            ReqButton.Visibility = Visibility.Visible;
            Show_by_Report.Visibility = Visibility.Visible;
            CallGraphButton.Visibility = Visibility.Visible;
            SignOutButton.Visibility = Visibility.Visible;


            BorderSelected.Visibility = Visibility.Collapsed;
            PopText.Visibility = Visibility.Collapsed;
            LabelSelectedReports.Visibility = Visibility.Collapsed;



        }
        public string[] GetArrayfromCombo()
        {
            List<string> list = new List<string>();
            foreach (DataRowView item in ComboBoxZone.Items)
            {
                string arr = item.Row[0].ToString();

                list.Add(arr);

            }
            string[] str = list.ToArray();
            return str;
        }
        public string[] GetArrayfromCombo1()
        {
            List<string> list = new List<string>();
            foreach (DataRowView item in ComboBoxZone1.Items)
            {
                string arr = item.Row[0].ToString();

                list.Add(arr);

            }
            string[] str = list.ToArray();
            return str;
        }

        private async void DisplayBasicTokenInfo(String authResult)
        {
            string query = "";

            // StackGrid.Visibility = Visibility.Hidden;
            // WindowMainName.Height = 766;
            //TokenInfoText.Text = "";
            DataTable dt = new DataTable();
            //DataTable dtUsage = new DataTable();
            DataTable dtUsage = new DataTable();
            DataTable dtUsage1 = new DataTable();
            DataTable dtUsage2 = new DataTable();
            DataTable dtUsage3 = new DataTable();
            DataTable dtUsage4 = new DataTable();
            DataTable dtUsage5 = new DataTable();
            DataTable dtUsage6 = new DataTable();
            DataTable dtUsage7 = new DataTable();
            DataTable dtUsage8 = new DataTable();
            DataTable dtCombo2 = new DataTable();
            DataTable dtUsageCombo2 = new DataTable();
            DataTable dtUsage1Combo2 = new DataTable();
            DataTable dtUsage2Combo2 = new DataTable();
            DataTable dtUsage3Combo2 = new DataTable();
            DataTable dtUsage4Combo2 = new DataTable();
            DataTable dtUsage5Combo2 = new DataTable();
            DataTable dtUsage6Combo2 = new DataTable();
            DataTable dtUsage7Combo2 = new DataTable();
            DataTable dtUsage8Combo2 = new DataTable();

            if (workspacename != "")
            {

                try
                {
                     
                                     
                        foreach (string item in items1)
                        {
                            //AdomdConnection connection = new AdomdConnection();
                            // connection.ConnectionString = GetConnectionString(ResultText.Text, item.Row[0].ToString());
                            // connection.Open();
                            //MessageBox.Show(item.ToString());  
                            //DataTable dt = new DataTable();
                            AdomdConnection connection = new AdomdConnection();
                            connection.ConnectionString = GetConnectionString(workspacename.ToString(), item.ToString());
                            connection.Open();
                            string queryString = "";



                            int pos = workspacename.ToString().LastIndexOf("/") + 1;
                            //MessageBox.Show(workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + " - " + item.ToString());
                            // WorkspaceLabel.Content = "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]";
                            queryString = "SELECT DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace], [CATALOG_NAME] AS [Report Name], [DIMENSION_UNIQUE_NAME] AS [Dataset Name], LEVEL_CAPTION AS [Column Name] FROM $System.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 AND LEVEL_NAME <> '(All)' order by [DIMENSION_UNIQUE_NAME]   ";
                            //queryString = check(query);
                            AdomdCommand cmd = connection.CreateCommand();
                            cmd.CommandText = queryString;
                            AdomdDataAdapter ad = new AdomdDataAdapter(queryString, connection);
                            ad.Fill(dt);


                            DataTable dt2 = new DataTable();
                            string queryString1 = "select DISTINCT" + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace], DATABASE_NAME as [Report Name],'['+[TABLE]+']' AS [Dataset Name],OBJECT AS [Column Name],EXPRESSION AS [Calculated Column Expression] from $SYSTEM.DISCOVER_CALC_DEPENDENCY WHERE OBJECT_TYPE = 'CALC_COLUMN' ";
                            AdomdCommand cmd1 = connection.CreateCommand();
                            cmd1.CommandText = queryString1;
                            AdomdDataAdapter ad1 = new AdomdDataAdapter(queryString1, connection);
                            ad1.Fill(dt2);

                            dt2.PrimaryKey = new DataColumn[] {
                        dt2.Columns["Report Name"],dt2.Columns["Dataset Name"],dt2.Columns["Column Name"] };


                            dt.Merge(dt2);
                            //  dt.DefaultView.Sort = "Dataset Name ASC";

                            DataTable dt4 = new DataTable();
                            string queryString3 = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace],  DATABASE_NAME as [Report Name],'['+[TABLE]+']' AS [Dataset Name],OBJECT AS [Column Name],EXPRESSION AS [Calculated Measure Expression] from $SYSTEM.DISCOVER_CALC_DEPENDENCY WHERE OBJECT_TYPE = 'MEASURE' ";
                            AdomdCommand cmd3 = connection.CreateCommand();
                            cmd3.CommandText = queryString3;
                            AdomdDataAdapter ad3 = new AdomdDataAdapter(queryString3, connection);
                            ad3.Fill(dt4);



                            dt.Merge(dt4);
                            //   dt.DefaultView.Sort = "Dataset Name ASC";




                            DataTable dt3 = new DataTable();
                            string queryString2 = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace],DATABASE_NAME as [Report Name],'['+[TABLE]+']' AS [Dataset Name],OBJECT AS [Column Name],EXPRESSION AS [Calculated Table Expression] from $SYSTEM.DISCOVER_CALC_DEPENDENCY WHERE OBJECT_TYPE = 'CALC_TABLE' ";
                            AdomdCommand cmd2 = connection.CreateCommand();
                            cmd2.CommandText = queryString2;
                            AdomdDataAdapter ad2 = new AdomdDataAdapter(queryString2, connection);
                            ad2.Fill(dt3);

                            dt3.PrimaryKey = new DataColumn[] {
                        dt3.Columns["Report Name"],dt3.Columns["Dataset Name"],dt3.Columns["Column Name"] };


                            dt.Merge(dt3);

                            // dt.DefaultView.Sort = "DatasetName ASC";

                        }
                   
                    foreach (string item in items1)
                    {
                        AdomdConnection connection = new AdomdConnection();
                        connection.ConnectionString = GetConnectionString(workspacename.ToString(), item.ToString());
                        connection.Open();
                        string queryString = "";
                        //Combo 1

                        int pos = workspacename.ToString().LastIndexOf("/") + 1;


                        DataTable dt3 = new DataTable();
                        string queryString2 = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],TableID,QueryDefinition FROM $SYSTEM.TMSCHEMA_PARTITIONS ";
                        AdomdCommand cmd2 = connection.CreateCommand();
                        cmd2.CommandText = queryString2;
                        AdomdDataAdapter ad2 = new AdomdDataAdapter(queryString2, connection);
                        ad2.Fill(dt3);


                        DataTable dt4 = new DataTable();
                        string queryString3 = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[ID] as [TableID],'['+[Name]+']' as [Table Name] FROM $SYSTEM.TMSCHEMA_TABLES ";
                        AdomdCommand cmd4 = connection.CreateCommand();
                        cmd4.CommandText = queryString3;
                        AdomdDataAdapter ad4 = new AdomdDataAdapter(queryString3, connection);
                        ad4.Fill(dt4);


                        var JoinResult = (from p in dt3.AsEnumerable()
                                          join t in dt4.AsEnumerable()
                                          on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("TableID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("TableID") } into ps
                                          from tnew in ps.DefaultIfEmpty()
                                          select new
                                          {
                                              WorkspaceName = p.Field<string>("Workspace"),
                                              ReportName = p.Field<string>("Report Name"),
                                              TableName = tnew.Field<string>("Table Name"),
                                              //Query1 = p.Field<string>("QueryDefinition").Trim().Replace(" ","").Replace(@"\r\n?|\n",""),
                                              // Query2 = findNthOccur(p.Field<string>("QueryDefinition"),'"',2),
                                              QueryDef = p.Field<string>("QueryDefinition"),
                                              //Check1= p.Field<string>("QueryDefinition").IndexOf("Item") > 0  && p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).IndexOf(".")>0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : p.Field<string>("QueryDefinition").IndexOf("Item") > 0  ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 6) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) - 1) + "." + p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : "",
                                              Source = p.Field<string>("QueryDefinition").IndexOf("Database") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 2) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) - 1) : "File Source/Derived Table",
                                              Path = p.Field<string>("QueryDefinition").IndexOf("Contents") > 0 || p.Field<string>("QueryDefinition").IndexOf("Files") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 2) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) - 1) : p.Field<string>("QueryDefinition").IndexOf("Database") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 3) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 4) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 3) - 1) : p.Field<string>("QueryDefinition").IndexOf("Table.NestedJoin") > 0 || p.Field<string>("QueryDefinition").IndexOf("Table.FromRows") > 0 ? "Derived Table inside PBI" : "No Database or Path available",
                                              Query = p.Field<string>("QueryDefinition").IndexOf("Query=") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 6) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) - 1).Replace("#(lf)", "") : p.Field<string>("QueryDefinition").IndexOf("NativeQuery") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).Replace("#(lf)", "") : "No Query Available",
                                              DatabaseItem = p.Field<string>("QueryDefinition").IndexOf("Item") > 0 && p.Field<string>("QueryDefinition").IndexOf("Contents") <= 0 && p.Field<string>("QueryDefinition").IndexOf("Query") <= 0 && p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).IndexOf(".") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : p.Field<string>("QueryDefinition").IndexOf("Item") > 0 && p.Field<string>("QueryDefinition").IndexOf("Contents") <= 0 && p.Field<string>("QueryDefinition").IndexOf("Query") <= 0 && p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).IndexOf(".") <= 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 6) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) - 1) + "." + p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : "No Database Item available",
                                          }).ToList();

                        dt4 = LINQResultToDataTable(JoinResult);
                        dt4.Columns["WorkspaceName"].ColumnName = "Workspace";
                        dt4.Columns["ReportName"].ColumnName = "Report Name";
                        dt4.Columns["TableName"].ColumnName = "Dataset Name";
                        dt4.Columns["Source"].ColumnName = "Source";
                        dt4.Columns["Path"].ColumnName = "Database Or Path";
                        // dt4.Columns["Query"].ColumnName = "Advance Editor Steps";


                        var JoinResult1 = (from p in dt.AsEnumerable()
                                           join t in dt4.AsEnumerable()
                                           on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<string>("Dataset Name") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<string>("Dataset Name") } into ps
                                           from tnew in ps
                                           select new
                                           {

                                               WorkspaceName = p.Field<string>("Workspace"),
                                               DatasetName = p.Field<string>("Dataset Name"),
                                               ReportName = p.Field<string>("Report Name"),
                                               ColumnName = p.Field<string>("Column Name"),
                                               Source = tnew == null ? "" : tnew.Field<string>("Source"),
                                               Path = tnew == null ? "" : tnew.Field<string>("Database Or Path"),
                                               Query = tnew == null ? "" : tnew.Field<string>("Query"),
                                               DatabaseItem = tnew == null ? "" : tnew.Field<string>("DatabaseItem"),
                                               // Check1= tnew == null ? "" : tnew.Field<string>("Check1"),
                                               Steps = tnew == null ? "" : tnew.Field<string>("QueryDef")
                                               //Check= tnew == null ? "" : tnew.Field<string>("Check")

                                           }).ToList();

                        dt4 = LINQResultToDataTable(JoinResult1);
                        dt4.Columns["WorkspaceName"].ColumnName = "Workspace";
                        dt4.Columns["ReportName"].ColumnName = "Report Name";
                        dt4.Columns["ColumnName"].ColumnName = "Column Name";
                        dt4.Columns["DatasetName"].ColumnName = "Dataset Name";
                        dt4.Columns["Source"].ColumnName = "Source";
                        dt4.Columns["Path"].ColumnName = "Database Or Path";
                        //dt4.Columns["Query"].ColumnName = "Advance Editor Steps";
                        dt4.PrimaryKey = new DataColumn[] {
                    dt4.Columns["Report Name"],dt4.Columns["Dataset Name"],dt4.Columns["Column Name"] };

                        dt.PrimaryKey = new DataColumn[] {
                    dt.Columns["Report Name"],dt.Columns["Dataset Name"],dt.Columns["Column Name"] };

                        dt.Merge(dt4);
                        /*dt.Columns["WorkspaceName"].ColumnName = "Workspace";
                        dt.Columns["ReportName"].ColumnName = "Report Name";
                        dt.Columns["ColumnName"].ColumnName = "Column Name";
                        dt.Columns["DatasetName"].ColumnName = "Dataset Name";*/
                        //dt.Columns["Source"].ColumnName = "Source";
                        //dt.Columns["Path"].ColumnName = "Database Or Path";
                        //dt.Columns["Query"].ColumnName = "Advance Editor Steps";

                        //dt.DefaultView.Sort = "DatasetName ASC";



                        pos = workspacename.ToString().LastIndexOf("/") + 1;


                        dt3 = new DataTable();
                        queryString2 = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],FromTableID,FromColumnID,ToTableID,ToColumnID,RefreshedTime FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS ";
                        cmd2 = connection.CreateCommand();
                        cmd2.CommandText = queryString2;
                        ad2 = new AdomdDataAdapter(queryString2, connection);
                        ad2.Fill(dt3);

                        if (dt3.Rows.Count > 0)
                        {



                            DataTable dt4Master = new DataTable();
                            string queryStringMaster = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[ID] AS [Dataset ID] ,'['+[Name]+']'  AS [Dataset Name] FROM $SYSTEM.TMSCHEMA_TABLES";
                            AdomdCommand cmd4Master = connection.CreateCommand();
                            cmd4Master.CommandText = queryStringMaster;
                            AdomdDataAdapter ad4Master = new AdomdDataAdapter(queryStringMaster, connection);
                            ad4Master.Fill(dt4Master);

                            DataTable dt4ColumnMaster = new DataTable();
                            string queryStringColumnMaster = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[ID] AS [Column ID],ExplicitName AS [Column Name],InferredName FROM $SYSTEM.TMSCHEMA_COLUMNS";
                            AdomdCommand cmd4ColumnMaster = connection.CreateCommand();
                            cmd4ColumnMaster.CommandText = queryStringColumnMaster;
                            AdomdDataAdapter ad4ColumnMaster = new AdomdDataAdapter(queryStringColumnMaster, connection);
                            ad4ColumnMaster.Fill(dt4ColumnMaster);

                            //MessageBox.Show(dt3.Columns["RefreshedTime"].DataType.ToString());


                            var JoinResult4 = (from p in dt3.AsEnumerable()
                                               join t in dt4Master.AsEnumerable()
                                               on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("FromTableID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Dataset ID") } into ps
                                               from tnew in ps.DefaultIfEmpty()
                                               select new
                                               {
                                                   WorkspaceName = p.Field<string>("Workspace"),
                                                   ReportName = p.Field<string>("Report Name"),
                                                   FromTableID = p.Field<System.UInt64>("FromTableID"),
                                                   ToTableID = p.Field<System.UInt64>("ToTableID"),
                                                   FromColumnID = p.Field<System.UInt64>("FromColumnID"),
                                                   ToColumnID = p.Field<System.UInt64>("ToColumnID"),
                                                   RefreshedTime = p.Field<System.DateTime>("RefreshedTime"),
                                                   FromTableName = tnew.Field<string>("Dataset Name")

                                               }).ToList();

                            dt3 = LINQResultToDataTable(JoinResult4);
                            dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                            dt3.Columns["ReportName"].ColumnName = "Report Name";
                            dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                            dt3.Columns["ToTableID"].ColumnName = "To Table ID";
                            dt3.Columns["FromTableID"].ColumnName = "From Table ID";

                            var JoinResult2 = (from p in dt3.AsEnumerable()
                                               join t in dt4Master.AsEnumerable()
                                               on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("To Table ID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Dataset ID") } into ps
                                               from tnew in ps.DefaultIfEmpty()
                                               select new
                                               {
                                                   WorkspaceName = p.Field<string>("Workspace"),
                                                   ReportName = p.Field<string>("Report Name"),
                                                   FromTableName = p.Field<string>("From Table Name"),
                                                   FromColumnID = p.Field<System.UInt64>("FromColumnID"),
                                                   ToTableName = tnew.Field<string>("Dataset Name"),
                                                   ToColumnID = p.Field<System.UInt64>("ToColumnID"),
                                                   RefreshedTime = p.Field<System.DateTime>("RefreshedTime")

                                               }).ToList();
                            dt3 = LINQResultToDataTable(JoinResult2);
                            dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                            dt3.Columns["ReportName"].ColumnName = "Report Name";
                            dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                            dt3.Columns["FromColumnID"].ColumnName = "From Column ID";
                            dt3.Columns["ToTableName"].ColumnName = "To Table Name";
                            dt3.Columns["ToColumnID"].ColumnName = "To Column ID";

                            var JoinResult3 = (from p in dt3.AsEnumerable()
                                               join t in dt4ColumnMaster.AsEnumerable()
                                               on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("From Column ID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Column ID") } into ps
                                               from tnew in ps.DefaultIfEmpty()
                                               select new
                                               {
                                                   WorkspaceName = p.Field<string>("Workspace"),
                                                   ReportName = p.Field<string>("Report Name"),
                                                   FromTableName = p.Field<string>("From Table Name"),
                                                   FromColumnID = p.Field<System.UInt64>("From Column ID"),
                                                   FromColumnName = tnew.Field<string>("Column Name") == null ? tnew.Field<string>("InferredName") : tnew.Field<string>("Column Name"),
                                                   ToTableName = p.Field<string>("To Table Name"),
                                                   ToColumnID = p.Field<System.UInt64>("To Column ID"),
                                                   RefreshedTime = p.Field<System.DateTime>("RefreshedTime")

                                               }).ToList();
                            dt3 = LINQResultToDataTable(JoinResult3);
                            dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                            dt3.Columns["ReportName"].ColumnName = "Report Name";
                            dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                            dt3.Columns["FromColumnID"].ColumnName = "From Column ID";
                            dt3.Columns["FromColumnName"].ColumnName = "From Column Name";
                            dt3.Columns["ToTableName"].ColumnName = "To Table Name";
                            dt3.Columns["ToColumnID"].ColumnName = "To Column ID";

                            var JoinResultTemp = (from p in dt3.AsEnumerable()
                                                  join t in dt4ColumnMaster.AsEnumerable()
                                                  on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("To Column ID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Column ID") } into ps
                                                  from tnew in ps.DefaultIfEmpty()
                                                  select new
                                                  {
                                                      WorkspaceName = p.Field<string>("Workspace"),
                                                      ReportName = p.Field<string>("Report Name"),
                                                      DatasetName = p.Field<string>("From Table Name"),
                                                      ColumnName = p.Field<string>("From Column Name"),
                                                      FromTableName = p.Field<string>("From Table Name"),
                                                      FromColumnName = p.Field<string>("From Column Name"),
                                                      ToTableName = p.Field<string>("To Table Name"),
                                                      ToColumnName = tnew.Field<string>("Column Name") == null ? tnew.Field<string>("InferredName") : tnew.Field<string>("Column Name"),
                                                      RefreshedTime = p.Field<System.DateTime>("RefreshedTime")



                                                  }).ToList();
                            dt3 = LINQResultToDataTable(JoinResultTemp);

                            dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                            dt3.Columns["ReportName"].ColumnName = "Report Name";
                            dt3.Columns["ColumnName"].ColumnName = "Column Name";
                            dt3.Columns["DatasetName"].ColumnName = "Dataset Name";
                            dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                            dt3.Columns["FromColumnName"].ColumnName = "From Column Name";
                            dt3.Columns["ToTableName"].ColumnName = "To Table Name";
                            dt3.Columns["ToColumnName"].ColumnName = "To Column Name";
                            dt3.Columns["RefreshedTime"].ColumnName = "Refreshed Time";

                            dt.Merge(dt3);

                        }




                        int posUsage = workspacename.ToString().LastIndexOf("/") + 1;


                        string queryUsage = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME,COLUMN_ID,ATTRIBUTE_NAME AS COLUMN_NAME,DATATYPE AS [Data Type],DICTIONARY_SIZE AS DICTIONARY_SIZE_BYTES,COLUMN_ENCODING AS COLUMN_ENCODING_INT from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMNS WHERE COLUMN_TYPE='BASIC_DATA' ";
                        AdomdCommand cmdUsage = connection.CreateCommand();
                        cmdUsage.CommandText = queryUsage;
                        AdomdDataAdapter ad4Usage = new AdomdDataAdapter(queryUsage, connection);
                        ad4Usage.Fill(dtUsage);




                        string queryUsage1 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME,COLUMN_ID AS STRUCTURE_NAME,USED_SIZE,TABLE_ID AS HIERARCHY_ID from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE LEFT( TABLE_ID,2 )='U$' ";
                        AdomdCommand cmdUsage1 = connection.CreateCommand();
                        cmdUsage1.CommandText = queryUsage1;
                        AdomdDataAdapter ad4Usage1 = new AdomdDataAdapter(queryUsage1, connection);
                        ad4Usage1.Fill(dtUsage1);





                        string queryUsage2 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME,COLUMN_ID AS STRUCTURE_NAME,SEGMENT_NUMBER,TABLE_PARTITION_NUMBER,USED_SIZE,TABLE_ID AS COLUMN_HIERARCHY_ID from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE LEFT( TABLE_ID,2 )='H$' ";
                        AdomdCommand cmdUsage2 = connection.CreateCommand();
                        cmdUsage2.CommandText = queryUsage2;
                        AdomdDataAdapter ad4Usage2 = new AdomdDataAdapter(queryUsage2, connection);
                        ad4Usage2.Fill(dtUsage2);




                        string queryUsage3 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME, PARTITION_NAME,COLUMN_ID AS COLUMN_NAME , SEGMENT_NUMBER,TABLE_PARTITION_NUMBER,RECORDS_COUNT AS SEGMENT_ROWS,USED_SIZE,COMPRESSION_TYPE,BITS_COUNT,BOOKMARK_BITS_COUNT,VERTIPAQ_STATE from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE RIGHT(LEFT( TABLE_ID,2 ),1)<>'$' ";
                        AdomdCommand cmdUsage3 = connection.CreateCommand();
                        cmdUsage3.CommandText = queryUsage3;
                        AdomdDataAdapter ad4Usage3 = new AdomdDataAdapter(queryUsage3, connection);
                        ad4Usage3.Fill(dtUsage3);




                        string queryUsage4 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME, TABLE_ID AS RELATIONSHIP_ID,USED_SIZE from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE  LEFT( TABLE_ID,2 )='R$' ";
                        AdomdCommand cmdUsage4 = connection.CreateCommand();
                        cmdUsage4.CommandText = queryUsage4;
                        AdomdDataAdapter ad4Usage4 = new AdomdDataAdapter(queryUsage4, connection);
                        ad4Usage4.Fill(dtUsage4);




                        string queryUsage5 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[NAME] AS TABLE_NAME,[RefreshedTime] FROM  $SYSTEM.TMSCHEMA_PARTITIONS  ";
                        AdomdCommand cmdUsage5 = connection.CreateCommand();
                        cmdUsage5.CommandText = queryUsage5;
                        AdomdDataAdapter ad4Usage5 = new AdomdDataAdapter(queryUsage5, connection);
                        ad4Usage5.Fill(dtUsage5);




                        string queryUsage6 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[ID] AS [Table ID],[Name] AS [Table Name] FROM  $SYSTEM.TMSCHEMA_TABLES ";
                        AdomdCommand cmdUsage6 = connection.CreateCommand();
                        cmdUsage6.CommandText = queryUsage6;
                        AdomdDataAdapter ad4Usage6 = new AdomdDataAdapter(queryUsage6, connection);
                        ad4Usage6.Fill(dtUsage6);




                        string queryUsage7 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],TABLEID AS [Table ID], [ID] AS [Column ID],ExplicitName AS [Column Name] FROM $SYSTEM.TMSCHEMA_COLUMNS ";
                        AdomdCommand cmdUsage7 = connection.CreateCommand();
                        cmdUsage7.CommandText = queryUsage7;
                        AdomdDataAdapter ad4Usage7 = new AdomdDataAdapter(queryUsage7, connection);
                        ad4Usage7.Fill(dtUsage7);




                        string queryUsage8 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[ID] AS [Relationship ID],[FromTableID],[FromColumnID],[FromCardinality],[ToTableID],[ToColumnID],[ToCardinality],[IsActive],CrossFilteringBehavior FROM $System.TMSCHEMA_RELATIONSHIPS";
                        AdomdCommand cmdUsage8 = connection.CreateCommand();
                        cmdUsage8.CommandText = queryUsage8;
                        AdomdDataAdapter ad4Usage8 = new AdomdDataAdapter(queryUsage8, connection);
                        ad4Usage8.Fill(dtUsage8);







                    }
                }
            
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString());
                }
            }
            if (workspacename1 != "")
            {


                try
                {


                    //  ComboBoxZone.DataContext = null;
                    // ComboBoxZone.ItemsSource = null;
                    //ComboBoxZone.Text = "";
                    //   Animation.Visibility = Visibility.Visible;
                    // ServerStack.Visibility = Visibility.Hidden;
                    //StackGrid.Visibility = Visibility.Hidden;
                    //items = new string[ComboBoxZone.Items.Count];


                    foreach (string item in items2)
                    {
                        //AdomdConnection connection = new AdomdConnection();
                        // connection.ConnectionString = GetConnectionString(ResultText.Text, item.Row[0].ToString());
                        // connection.Open();
                        //MessageBox.Show(item.ToString());  
                        //DataTable dt = new DataTable();
                        AdomdConnection connection = new AdomdConnection();
                        connection.ConnectionString = GetConnectionString(workspacename1.ToString(), item.ToString());
                        connection.Open();
                        string queryString = "";



                        int pos = workspacename1.ToString().LastIndexOf("/") + 1;
                        //MessageBox.Show(workspacename1.ToString().Substring(pos, workspacename1.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + " - " + item.ToString());
                        // WorkspaceLabel.Content = "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]";
                        queryString = "SELECT DISTINCT " + "'" + workspacename1.ToString().Substring(pos, workspacename1.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace], [CATALOG_NAME] AS [Report Name], [DIMENSION_UNIQUE_NAME] AS [Dataset Name], LEVEL_CAPTION AS [Column Name] FROM $System.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 AND LEVEL_NAME <> '(All)' order by [DIMENSION_UNIQUE_NAME]   ";
                        //queryString = check(query);
                        AdomdCommand cmd = connection.CreateCommand();
                        cmd.CommandText = queryString;
                        AdomdDataAdapter ad = new AdomdDataAdapter(queryString, connection);
                        ad.Fill(dtCombo2);


                        DataTable dt2 = new DataTable();
                        string queryString1 = "select DISTINCT" + "'" + workspacename1.ToString().Substring(pos, workspacename1.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace], DATABASE_NAME as [Report Name],'['+[TABLE]+']' AS [Dataset Name],OBJECT AS [Column Name],EXPRESSION AS [Calculated Column Expression] from $SYSTEM.DISCOVER_CALC_DEPENDENCY WHERE OBJECT_TYPE = 'CALC_COLUMN' ";
                        AdomdCommand cmd1 = connection.CreateCommand();
                        cmd1.CommandText = queryString1;
                        AdomdDataAdapter ad1 = new AdomdDataAdapter(queryString1, connection);
                        ad1.Fill(dt2);

                        dt2.PrimaryKey = new DataColumn[] {
                    dt2.Columns["Report Name"],dt2.Columns["Dataset Name"],dt2.Columns["Column Name"] };


                        dt.Merge(dt2);
                        //  dt.DefaultView.Sort = "Dataset Name ASC";

                        DataTable dt4 = new DataTable();
                        string queryString3 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(pos, workspacename1.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace],  DATABASE_NAME as [Report Name],'['+[TABLE]+']' AS [Dataset Name],OBJECT AS [Column Name],EXPRESSION AS [Calculated Measure Expression] from $SYSTEM.DISCOVER_CALC_DEPENDENCY WHERE OBJECT_TYPE = 'MEASURE' ";
                        AdomdCommand cmd3 = connection.CreateCommand();
                        cmd3.CommandText = queryString3;
                        AdomdDataAdapter ad3 = new AdomdDataAdapter(queryString3, connection);
                        ad3.Fill(dt4);



                        dt.Merge(dt4);
                        //   dt.DefaultView.Sort = "Dataset Name ASC";




                        DataTable dt3 = new DataTable();
                        string queryString2 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(pos, workspacename1.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace],DATABASE_NAME as [Report Name],'['+[TABLE]+']' AS [Dataset Name],OBJECT AS [Column Name],EXPRESSION AS [Calculated Table Expression] from $SYSTEM.DISCOVER_CALC_DEPENDENCY WHERE OBJECT_TYPE = 'CALC_TABLE' ";
                        AdomdCommand cmd2 = connection.CreateCommand();
                        cmd2.CommandText = queryString2;
                        AdomdDataAdapter ad2 = new AdomdDataAdapter(queryString2, connection);
                        ad2.Fill(dt3);

                        dt3.PrimaryKey = new DataColumn[] {
                    dt3.Columns["Report Name"],dt3.Columns["Dataset Name"],dt3.Columns["Column Name"] };


                        dt.Merge(dt3);

                        // dt.DefaultView.Sort = "DatasetName ASC";

                    }
                    foreach (string item in items2)
                    {
                        AdomdConnection connection = new AdomdConnection();
                        connection.ConnectionString = GetConnectionString(workspacename1.ToString(), item.ToString());
                        connection.Open();
                        string queryString = "";
                        //Combo 1

                        int pos = workspacename1.ToString().LastIndexOf("/") + 1;
                        // MessageBox.Show(item.ToString());

                        DataTable dt3 = new DataTable();
                        string queryString2 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(pos, workspacename1.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],TableID,QueryDefinition FROM $SYSTEM.TMSCHEMA_PARTITIONS ";
                        AdomdCommand cmd2 = connection.CreateCommand();
                        cmd2.CommandText = queryString2;
                        AdomdDataAdapter ad2 = new AdomdDataAdapter(queryString2, connection);
                        ad2.Fill(dt3);


                        DataTable dt4 = new DataTable();
                        string queryString3 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(pos, workspacename1.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[ID] as [TableID],'['+[Name]+']' as [Table Name] FROM $SYSTEM.TMSCHEMA_TABLES ";
                        AdomdCommand cmd4 = connection.CreateCommand();
                        cmd4.CommandText = queryString3;
                        AdomdDataAdapter ad4 = new AdomdDataAdapter(queryString3, connection);
                        ad4.Fill(dt4);


                        var JoinResult = (from p in dt3.AsEnumerable()
                                          join t in dt4.AsEnumerable()
                                          on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("TableID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("TableID") } into ps
                                          from tnew in ps.DefaultIfEmpty()
                                          select new
                                          {
                                              WorkspaceName = p.Field<string>("Workspace"),
                                              ReportName = p.Field<string>("Report Name"),
                                              TableName = tnew.Field<string>("Table Name"),
                                              //Query1 = p.Field<string>("QueryDefinition").Trim().Replace(" ","").Replace(@"\r\n?|\n",""),
                                              // Query2 = findNthOccur(p.Field<string>("QueryDefinition"),'"',2),
                                              QueryDef = p.Field<string>("QueryDefinition"),
                                              //Check1= p.Field<string>("QueryDefinition").IndexOf("Item") > 0  && p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).IndexOf(".")>0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : p.Field<string>("QueryDefinition").IndexOf("Item") > 0  ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 6) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) - 1) + "." + p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : "",
                                              Source = p.Field<string>("QueryDefinition").IndexOf("Database") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 2) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) - 1) : "File Source/Derived Table",
                                              Path = p.Field<string>("QueryDefinition").IndexOf("Contents") > 0 || p.Field<string>("QueryDefinition").IndexOf("Files") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 2) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) - 1) : p.Field<string>("QueryDefinition").IndexOf("Database") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 3) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 4) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 3) - 1) : p.Field<string>("QueryDefinition").IndexOf("Table.NestedJoin") > 0 || p.Field<string>("QueryDefinition").IndexOf("Table.FromRows") > 0 ? "Derived Table inside PBI" : "No Database or Path available",
                                              Query = p.Field<string>("QueryDefinition").IndexOf("Query=") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 6) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) - 1).Replace("#(lf)", "") : p.Field<string>("QueryDefinition").IndexOf("NativeQuery") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).Replace("#(lf)", "") : "No Query Available",
                                              DatabaseItem = p.Field<string>("QueryDefinition").IndexOf("Item") > 0 && p.Field<string>("QueryDefinition").IndexOf("Contents") <= 0 && p.Field<string>("QueryDefinition").IndexOf("Query") <= 0 && p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).IndexOf(".") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : p.Field<string>("QueryDefinition").IndexOf("Item") > 0 && p.Field<string>("QueryDefinition").IndexOf("Contents") <= 0 && p.Field<string>("QueryDefinition").IndexOf("Query") <= 0 && p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).IndexOf(".") <= 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 6) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) - 1) + "." + p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : "No Database Item available",
                                          }).ToList();

                        dt4 = LINQResultToDataTable(JoinResult);
                        dt4.Columns["WorkspaceName"].ColumnName = "Workspace";
                        dt4.Columns["ReportName"].ColumnName = "Report Name";
                        dt4.Columns["TableName"].ColumnName = "Dataset Name";
                        dt4.Columns["Source"].ColumnName = "Source";
                        dt4.Columns["Path"].ColumnName = "Database Or Path";
                        // dt4.Columns["Query"].ColumnName = "Advance Editor Steps";


                        var JoinResult1 = (from p in dt.AsEnumerable()
                                           join t in dt4.AsEnumerable()
                                           on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<string>("Dataset Name") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<string>("Dataset Name") } into ps
                                           from tnew in ps
                                           select new
                                           {

                                               WorkspaceName = p.Field<string>("Workspace"),
                                               DatasetName = p.Field<string>("Dataset Name"),
                                               ReportName = p.Field<string>("Report Name"),
                                               ColumnName = p.Field<string>("Column Name"),
                                               Source = tnew == null ? "" : tnew.Field<string>("Source"),
                                               Path = tnew == null ? "" : tnew.Field<string>("Database Or Path"),
                                               Query = tnew == null ? "" : tnew.Field<string>("Query"),
                                               DatabaseItem = tnew == null ? "" : tnew.Field<string>("DatabaseItem"),
                                               // Check1= tnew == null ? "" : tnew.Field<string>("Check1"),
                                               Steps = tnew == null ? "" : tnew.Field<string>("QueryDef")
                                               //Check= tnew == null ? "" : tnew.Field<string>("Check")

                                           }).ToList();

                        dt4 = LINQResultToDataTable(JoinResult1);
                        dt4.Columns["WorkspaceName"].ColumnName = "Workspace";
                        dt4.Columns["ReportName"].ColumnName = "Report Name";
                        dt4.Columns["ColumnName"].ColumnName = "Column Name";
                        dt4.Columns["DatasetName"].ColumnName = "Dataset Name";
                        dt4.Columns["Source"].ColumnName = "Source";
                        dt4.Columns["Path"].ColumnName = "Database Or Path";
                        //dt4.Columns["Query"].ColumnName = "Advance Editor Steps";
                        dt4.PrimaryKey = new DataColumn[] {
                    dt4.Columns["Report Name"],dt4.Columns["Dataset Name"],dt4.Columns["Column Name"] };

                        dt.PrimaryKey = new DataColumn[] {
                    dt.Columns["Report Name"],dt.Columns["Dataset Name"],dt.Columns["Column Name"] };

                        dt.Merge(dt4);
                        /*dt.Columns["WorkspaceName"].ColumnName = "Workspace";
                        dt.Columns["ReportName"].ColumnName = "Report Name";
                        dt.Columns["ColumnName"].ColumnName = "Column Name";
                        dt.Columns["DatasetName"].ColumnName = "Dataset Name";*/
                        //dt.Columns["Source"].ColumnName = "Source";
                        //dt.Columns["Path"].ColumnName = "Database Or Path";
                        //dt.Columns["Query"].ColumnName = "Advance Editor Steps";

                        //dt.DefaultView.Sort = "DatasetName ASC";



                        pos = workspacename1.ToString().LastIndexOf("/") + 1;


                        dt3 = new DataTable();
                        queryString2 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(pos, workspacename1.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],FromTableID,FromColumnID,ToTableID,ToColumnID,RefreshedTime FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS ";
                        cmd2 = connection.CreateCommand();
                        cmd2.CommandText = queryString2;
                        ad2 = new AdomdDataAdapter(queryString2, connection);
                        ad2.Fill(dt3);

                        if (dt3.Rows.Count > 0)
                        {



                            DataTable dt4Master = new DataTable();
                            string queryStringMaster = "select DISTINCT " + "'" + workspacename1.ToString().Substring(pos, workspacename1.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[ID] AS [Dataset ID] ,'['+[Name]+']'  AS [Dataset Name] FROM $SYSTEM.TMSCHEMA_TABLES";
                            AdomdCommand cmd4Master = connection.CreateCommand();
                            cmd4Master.CommandText = queryStringMaster;
                            AdomdDataAdapter ad4Master = new AdomdDataAdapter(queryStringMaster, connection);
                            ad4Master.Fill(dt4Master);

                            DataTable dt4ColumnMaster = new DataTable();
                            string queryStringColumnMaster = "select DISTINCT " + "'" + workspacename1.ToString().Substring(pos, workspacename1.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[ID] AS [Column ID],ExplicitName AS [Column Name],InferredName FROM $SYSTEM.TMSCHEMA_COLUMNS";
                            AdomdCommand cmd4ColumnMaster = connection.CreateCommand();
                            cmd4ColumnMaster.CommandText = queryStringColumnMaster;
                            AdomdDataAdapter ad4ColumnMaster = new AdomdDataAdapter(queryStringColumnMaster, connection);
                            ad4ColumnMaster.Fill(dt4ColumnMaster);

                            //MessageBox.Show(dt3.Columns["RefreshedTime"].DataType.ToString());


                            var JoinResult4 = (from p in dt3.AsEnumerable()
                                               join t in dt4Master.AsEnumerable()
                                               on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("FromTableID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Dataset ID") } into ps
                                               from tnew in ps.DefaultIfEmpty()
                                               select new
                                               {
                                                   WorkspaceName = p.Field<string>("Workspace"),
                                                   ReportName = p.Field<string>("Report Name"),
                                                   FromTableID = p.Field<System.UInt64>("FromTableID"),
                                                   ToTableID = p.Field<System.UInt64>("ToTableID"),
                                                   FromColumnID = p.Field<System.UInt64>("FromColumnID"),
                                                   ToColumnID = p.Field<System.UInt64>("ToColumnID"),
                                                   RefreshedTime = p.Field<System.DateTime>("RefreshedTime"),
                                                   FromTableName = tnew.Field<string>("Dataset Name")

                                               }).ToList();

                            dt3 = LINQResultToDataTable(JoinResult4);
                            dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                            dt3.Columns["ReportName"].ColumnName = "Report Name";
                            dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                            dt3.Columns["ToTableID"].ColumnName = "To Table ID";
                            dt3.Columns["FromTableID"].ColumnName = "From Table ID";

                            var JoinResult2 = (from p in dt3.AsEnumerable()
                                               join t in dt4Master.AsEnumerable()
                                               on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("To Table ID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Dataset ID") } into ps
                                               from tnew in ps.DefaultIfEmpty()
                                               select new
                                               {
                                                   WorkspaceName = p.Field<string>("Workspace"),
                                                   ReportName = p.Field<string>("Report Name"),
                                                   FromTableName = p.Field<string>("From Table Name"),
                                                   FromColumnID = p.Field<System.UInt64>("FromColumnID"),
                                                   ToTableName = tnew.Field<string>("Dataset Name"),
                                                   ToColumnID = p.Field<System.UInt64>("ToColumnID"),
                                                   RefreshedTime = p.Field<System.DateTime>("RefreshedTime")

                                               }).ToList();
                            dt3 = LINQResultToDataTable(JoinResult2);
                            dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                            dt3.Columns["ReportName"].ColumnName = "Report Name";
                            dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                            dt3.Columns["FromColumnID"].ColumnName = "From Column ID";
                            dt3.Columns["ToTableName"].ColumnName = "To Table Name";
                            dt3.Columns["ToColumnID"].ColumnName = "To Column ID";

                            var JoinResult3 = (from p in dt3.AsEnumerable()
                                               join t in dt4ColumnMaster.AsEnumerable()
                                               on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("From Column ID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Column ID") } into ps
                                               from tnew in ps.DefaultIfEmpty()
                                               select new
                                               {
                                                   WorkspaceName = p.Field<string>("Workspace"),
                                                   ReportName = p.Field<string>("Report Name"),
                                                   FromTableName = p.Field<string>("From Table Name"),
                                                   FromColumnID = p.Field<System.UInt64>("From Column ID"),
                                                   FromColumnName = tnew.Field<string>("Column Name") == null ? tnew.Field<string>("InferredName") : tnew.Field<string>("Column Name"),
                                                   ToTableName = p.Field<string>("To Table Name"),
                                                   ToColumnID = p.Field<System.UInt64>("To Column ID"),
                                                   RefreshedTime = p.Field<System.DateTime>("RefreshedTime")

                                               }).ToList();
                            dt3 = LINQResultToDataTable(JoinResult3);
                            dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                            dt3.Columns["ReportName"].ColumnName = "Report Name";
                            dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                            dt3.Columns["FromColumnID"].ColumnName = "From Column ID";
                            dt3.Columns["FromColumnName"].ColumnName = "From Column Name";
                            dt3.Columns["ToTableName"].ColumnName = "To Table Name";
                            dt3.Columns["ToColumnID"].ColumnName = "To Column ID";

                            var JoinResultTemp = (from p in dt3.AsEnumerable()
                                                  join t in dt4ColumnMaster.AsEnumerable()
                                                  on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("To Column ID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Column ID") } into ps
                                                  from tnew in ps.DefaultIfEmpty()
                                                  select new
                                                  {
                                                      WorkspaceName = p.Field<string>("Workspace"),
                                                      ReportName = p.Field<string>("Report Name"),
                                                      DatasetName = p.Field<string>("From Table Name"),
                                                      ColumnName = p.Field<string>("From Column Name"),
                                                      FromTableName = p.Field<string>("From Table Name"),
                                                      FromColumnName = p.Field<string>("From Column Name"),
                                                      ToTableName = p.Field<string>("To Table Name"),
                                                      ToColumnName = tnew.Field<string>("Column Name") == null ? tnew.Field<string>("InferredName") : tnew.Field<string>("Column Name"),
                                                      RefreshedTime = p.Field<System.DateTime>("RefreshedTime")



                                                  }).ToList();
                            dt3 = LINQResultToDataTable(JoinResultTemp);

                            dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                            dt3.Columns["ReportName"].ColumnName = "Report Name";
                            dt3.Columns["ColumnName"].ColumnName = "Column Name";
                            dt3.Columns["DatasetName"].ColumnName = "Dataset Name";
                            dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                            dt3.Columns["FromColumnName"].ColumnName = "From Column Name";
                            dt3.Columns["ToTableName"].ColumnName = "To Table Name";
                            dt3.Columns["ToColumnName"].ColumnName = "To Column Name";
                            dt3.Columns["RefreshedTime"].ColumnName = "Refreshed Time";

                            dt.Merge(dt3);

                        }




                        int posUsage = workspacename1.ToString().LastIndexOf("/") + 1;


                        string queryUsage = "select DISTINCT " + "'" + workspacename1.ToString().Substring(posUsage, workspacename1.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME,COLUMN_ID,ATTRIBUTE_NAME AS COLUMN_NAME,DATATYPE AS [Data Type],DICTIONARY_SIZE AS DICTIONARY_SIZE_BYTES,COLUMN_ENCODING AS COLUMN_ENCODING_INT from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMNS WHERE COLUMN_TYPE='BASIC_DATA' ";
                        AdomdCommand cmdUsage = connection.CreateCommand();
                        cmdUsage.CommandText = queryUsage;
                        AdomdDataAdapter ad4Usage = new AdomdDataAdapter(queryUsage, connection);
                        ad4Usage.Fill(dtUsageCombo2);




                        string queryUsage1 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(posUsage, workspacename1.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME,COLUMN_ID AS STRUCTURE_NAME,USED_SIZE,TABLE_ID AS HIERARCHY_ID from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE LEFT( TABLE_ID,2 )='U$' ";
                        AdomdCommand cmdUsage1 = connection.CreateCommand();
                        cmdUsage1.CommandText = queryUsage1;
                        AdomdDataAdapter ad4Usage1 = new AdomdDataAdapter(queryUsage1, connection);
                        ad4Usage1.Fill(dtUsage1Combo2);





                        string queryUsage2 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(posUsage, workspacename1.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME,COLUMN_ID AS STRUCTURE_NAME,SEGMENT_NUMBER,TABLE_PARTITION_NUMBER,USED_SIZE,TABLE_ID AS COLUMN_HIERARCHY_ID from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE LEFT( TABLE_ID,2 )='H$' ";
                        AdomdCommand cmdUsage2 = connection.CreateCommand();
                        cmdUsage2.CommandText = queryUsage2;
                        AdomdDataAdapter ad4Usage2 = new AdomdDataAdapter(queryUsage2, connection);
                        ad4Usage2.Fill(dtUsage2Combo2);




                        string queryUsage3 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(posUsage, workspacename1.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME, PARTITION_NAME,COLUMN_ID AS COLUMN_NAME , SEGMENT_NUMBER,TABLE_PARTITION_NUMBER,RECORDS_COUNT AS SEGMENT_ROWS,USED_SIZE,COMPRESSION_TYPE,BITS_COUNT,BOOKMARK_BITS_COUNT,VERTIPAQ_STATE from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE RIGHT(LEFT( TABLE_ID,2 ),1)<>'$' ";
                        AdomdCommand cmdUsage3 = connection.CreateCommand();
                        cmdUsage3.CommandText = queryUsage3;
                        AdomdDataAdapter ad4Usage3 = new AdomdDataAdapter(queryUsage3, connection);
                        ad4Usage3.Fill(dtUsage3Combo2);




                        string queryUsage4 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(posUsage, workspacename1.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME, TABLE_ID AS RELATIONSHIP_ID,USED_SIZE from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE  LEFT( TABLE_ID,2 )='R$' ";
                        AdomdCommand cmdUsage4 = connection.CreateCommand();
                        cmdUsage4.CommandText = queryUsage4;
                        AdomdDataAdapter ad4Usage4 = new AdomdDataAdapter(queryUsage4, connection);
                        ad4Usage4.Fill(dtUsage4Combo2);




                        string queryUsage5 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(posUsage, workspacename1.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[NAME] AS TABLE_NAME,[RefreshedTime] FROM  $SYSTEM.TMSCHEMA_PARTITIONS  ";
                        AdomdCommand cmdUsage5 = connection.CreateCommand();
                        cmdUsage5.CommandText = queryUsage5;
                        AdomdDataAdapter ad4Usage5 = new AdomdDataAdapter(queryUsage5, connection);
                        ad4Usage5.Fill(dtUsage5Combo2);




                        string queryUsage6 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(posUsage, workspacename1.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[ID] AS [Table ID],[Name] AS [Table Name] FROM  $SYSTEM.TMSCHEMA_TABLES ";
                        AdomdCommand cmdUsage6 = connection.CreateCommand();
                        cmdUsage6.CommandText = queryUsage6;
                        AdomdDataAdapter ad4Usage6 = new AdomdDataAdapter(queryUsage6, connection);
                        ad4Usage6.Fill(dtUsage6Combo2);




                        string queryUsage7 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(posUsage, workspacename1.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],TABLEID AS [Table ID], [ID] AS [Column ID],ExplicitName AS [Column Name] FROM $SYSTEM.TMSCHEMA_COLUMNS ";
                        AdomdCommand cmdUsage7 = connection.CreateCommand();
                        cmdUsage7.CommandText = queryUsage7;
                        AdomdDataAdapter ad4Usage7 = new AdomdDataAdapter(queryUsage7, connection);
                        ad4Usage7.Fill(dtUsage7Combo2);




                        string queryUsage8 = "select DISTINCT " + "'" + workspacename1.ToString().Substring(posUsage, workspacename1.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + item.ToString() + "' as [Report Name],[ID] AS [Relationship ID],[FromTableID],[FromColumnID],[FromCardinality],[ToTableID],[ToColumnID],[ToCardinality],[IsActive],CrossFilteringBehavior FROM $System.TMSCHEMA_RELATIONSHIPS";
                        AdomdCommand cmdUsage8 = connection.CreateCommand();
                        cmdUsage8.CommandText = queryUsage8;
                        AdomdDataAdapter ad4Usage8 = new AdomdDataAdapter(queryUsage8, connection);
                        ad4Usage8.Fill(dtUsage8Combo2);



                        dt.Merge(dtCombo2);
                        dtUsage.Merge(dtUsageCombo2);
                        dtUsage1.Merge(dtUsage1Combo2);
                        dtUsage2.Merge(dtUsage2Combo2);
                        dtUsage3.Merge(dtUsage3Combo2);
                        dtUsage4.Merge(dtUsage4Combo2);
                        dtUsage5.Merge(dtUsage5Combo2);
                        dtUsage6.Merge(dtUsage6Combo2);
                        dtUsage7.Merge(dtUsage7Combo2);
                        dtUsage8.Merge(dtUsage8Combo2);



                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message.ToString());
                }
            }


            createsqltable(dt, "Metadata");

            createsqltableUsage(dtUsage, "Dictionary_Usage");
            createsqltableUsage(dtUsage1, "User_Hierarchy");
            createsqltableUsage(dtUsage2, "Hierarchy");
            createsqltableUsage(dtUsage3, "Data_Size");
            createsqltableUsage(dtUsage4, "Relationships_Size");
            createsqltableUsage(dtUsage5, "Last_Update");
            createsqltableUsage(dtUsage6, "TMSchema_Table");
            createsqltableUsage(dtUsage7, "TMSchema_Columns");
            createsqltableUsage(dtUsage8, "TMSchema_Relationships");





        }




        private void ResultText_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }
        private string GetConnectionString(string server, string db)
        {



            return $"Provider=MSOLAP;" +
                $"Data Source={server};" +
                $"Initial Catalog={db};" +
                $"Persist Security Info=True;" +
                $"Impersonation Level=Impersonate";
            //Animation.Visibility = Visibility.Collapsed;

        }
        private string GetConnectionStringForCombo(string server)
        {

            return $"Provider=MSOLAP;" +
                $"Data Source={server};" +
                $"Persist Security Info=True;" +
                $"Impersonation Level=Impersonate";
            // Animation.Visibility = Visibility.Collapsed;
        }

        public async void BindComboBox(ComboBox comboBoxName)
        {

            try
            {
                if (ResultText.Text != "" && ResultText2.Visibility == Visibility.Collapsed)
                {
                    AdomdConnection connection = new AdomdConnection();
                    connection.ConnectionString = GetConnectionStringForCombo(ResultText.Text);
                    connection.Open();
                    string queryString = "SELECT DISTINCT [CATALOG_NAME] FROM $System.DBSCHEMA_CATALOGS;";
                    AdomdCommand cmd = connection.CreateCommand();
                    cmd.CommandText = queryString;
                    AdomdDataAdapter ad = new AdomdDataAdapter(queryString, connection);
                    DataSet ds = new DataSet();
                    ad.Fill(ds, "DBSCHEMA_CATALOGS");
                    comboBoxName.ItemsSource = ds.Tables[0].DefaultView;
                    comboBoxName.SelectedIndex = 1;
                    comboBoxName.SelectedValue= ds.Tables[0].Columns["CATALOG_NAME"].ToString();
                    comboBoxName.DisplayMemberPath = ds.Tables[0].Columns["CATALOG_NAME"].ToString();
                    comboBoxName.SelectedValuePath = ds.Tables[0].Columns["CATALOG_NAME"].ToString();

                    SignOutButton.Visibility = Visibility.Collapsed;



                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Connection String should be valid");
            }
        }
        public async void BindComboBoxServer2(ComboBox comboBoxName)
        {
            Animation.Visibility = Visibility.Visible;
            ServerStack.Visibility = Visibility.Hidden;
            //StackGrid.Visibility = Visibility.Hidden;0
            try
            {
                if (ResultText2.Text != "" && ResultText3.Visibility == Visibility.Collapsed)
                {

                    AdomdConnection connection = new AdomdConnection();
                    connection.ConnectionString = GetConnectionStringForCombo(ResultText2.Text);
                    connection.Open();
                    string queryString = "SELECT DISTINCT [CATALOG_NAME] FROM $System.DBSCHEMA_CATALOGS;";
                    AdomdCommand cmd = connection.CreateCommand();
                    cmd.CommandText = queryString;
                    AdomdDataAdapter ad = new AdomdDataAdapter(queryString, connection);
                    DataSet ds = new DataSet();
                    ad.Fill(ds, "DBSCHEMA_CATALOGS");
                    comboBoxName.ItemsSource = ds.Tables[0].DefaultView;
                    comboBoxName.DisplayMemberPath = ds.Tables[0].Columns["CATALOG_NAME"].ToString();
                    comboBoxName.SelectedValuePath = ds.Tables[0].Columns["CATALOG_NAME"].ToString();

                    CallGraphButton.IsEnabled = true;
                    CallGraphButton.Visibility = Visibility.Visible;
                    Show_by_Report.Visibility = Visibility.Visible;
                    WrapCheck.Visibility = Visibility.Collapsed;
                    BorderBox.Visibility = Visibility.Collapsed;
                    SignOutButton.Visibility = Visibility.Collapsed;

                    Get_Database.Visibility = Visibility.Visible;
                    Animation.Visibility = Visibility.Collapsed;
                    StackGrid.Visibility = Visibility.Visible;
                    ServerStack.Visibility = Visibility.Visible;
                }
                else if (ResultText2.Text != "" && ResultText3.Visibility == Visibility.Visible && ResultText3.Text != "")
                {

                    AdomdConnection connection = new AdomdConnection();
                    connection.ConnectionString = GetConnectionStringForCombo(ResultText2.Text);
                    connection.Open();
                    string queryString = "SELECT DISTINCT [CATALOG_NAME] FROM $System.DBSCHEMA_CATALOGS;";
                    AdomdCommand cmd = connection.CreateCommand();
                    cmd.CommandText = queryString;
                    AdomdDataAdapter ad = new AdomdDataAdapter(queryString, connection);
                    DataSet ds = new DataSet();
                    ad.Fill(ds, "DBSCHEMA_CATALOGS");
                    comboBoxName.ItemsSource = ds.Tables[0].DefaultView;
                    comboBoxName.DisplayMemberPath = ds.Tables[0].Columns["CATALOG_NAME"].ToString();
                    comboBoxName.SelectedValuePath = ds.Tables[0].Columns["CATALOG_NAME"].ToString();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Connection String should be valid");
            }
        }

        public async void BindComboBoxServer3(ComboBox comboBoxName)
        {
            Animation.Visibility = Visibility.Visible;
            ServerStack.Visibility = Visibility.Hidden;
            //StackGrid.Visibility = Visibility.Hidden;
            try
            {


                AdomdConnection connection = new AdomdConnection();
                connection.ConnectionString = GetConnectionStringForCombo(ResultText3.Text);
                connection.Open();
                string queryString = "SELECT DISTINCT [CATALOG_NAME] FROM $System.DBSCHEMA_CATALOGS;";
                AdomdCommand cmd = connection.CreateCommand();
                cmd.CommandText = queryString;
                AdomdDataAdapter ad = new AdomdDataAdapter(queryString, connection);
                DataSet ds = new DataSet();
                ad.Fill(ds, "DBSCHEMA_CATALOGS");
                comboBoxName.ItemsSource = ds.Tables[0].DefaultView;
                comboBoxName.DisplayMemberPath = ds.Tables[0].Columns["CATALOG_NAME"].ToString();
                comboBoxName.SelectedValuePath = ds.Tables[0].Columns["CATALOG_NAME"].ToString();

                CallGraphButton.IsEnabled = true;
                CallGraphButton.Visibility = Visibility.Visible;
                Show_by_Report.Visibility = Visibility.Visible;
                WrapCheck.Visibility = Visibility.Collapsed;
                BorderBox.Visibility = Visibility.Collapsed;
                SignOutButton.Visibility = Visibility.Collapsed;

                Get_Database.Visibility = Visibility.Visible;
                Animation.Visibility = Visibility.Collapsed;
                StackGrid.Visibility = Visibility.Visible;
                ServerStack.Visibility = Visibility.Visible;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Connection String should be valid");
            }
        }


        private void TokenInfoText_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }
        private void backgroundWorker2_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {

            // This is where the processor intensive code should go
            BindBox();
            // AddElementsInList();
            // BindCountryDropDown();


            // if we need any output to be used, put it in the DoWorkEventArgs object
            e.Result = "all done";
            //If the process exits the loop, ensure that progress is set to 100%
            //Remember in the loop we set i < 100 so in theory the process will complete at 99%
            backgroundWorker1.ReportProgress(100);
        }
        private void backgroundWorker2_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {

            }
            else if (e.Error != null)
            {

            }
            else
            {
                if (ResultText.Text != "" && ResultText2.Visibility == Visibility.Collapsed && ResultText3.Visibility == Visibility.Collapsed)
                {
                    ComboBoxZone.ItemsSource = null;
                    ComboBoxZone.Items.Clear();
                    ComboBoxZone.ItemsSource = ds.Tables[0].DefaultView;
                    view = ds.Tables[0].DefaultView;

                    //ComboBoxZone.SelectedValuePath = ds.Tables[0].Columns["CATALOG_NAME"].ToString();
                    //dataGrid1.Visibility = Visibility.Visible;
                    Animation.Visibility = Visibility.Collapsed;
                    StackGrid.Visibility = Visibility.Visible;
                    button1.Visibility = Visibility.Visible;
                    ReqButton.Visibility = Visibility.Visible;
                    Show_by_Report.Visibility = Visibility.Visible;
                    CallGraphButton.Visibility = Visibility.Visible;
                    ServerStack.Visibility = Visibility.Visible;
                    SignOutButton.Visibility = Visibility.Visible;
                    GenerateMetadata.Visibility = Visibility.Collapsed;
                    Output.Visibility = Visibility.Collapsed;

                    if (GenerateMetadata.IsChecked == true)
                    {
                        //   button1.Visibility = Visibility.Visible;
                        //  ReqButton.Visibility = Visibility.Visible;
                        //  Show_by_Report.Visibility = Visibility.Visible;
                        // CallGraphButton.Visibility = Visibility.Visible;
                    }
                }
                else if (ResultText.Text != "" && ResultText2.Text != "" && ResultText3.Visibility == Visibility.Collapsed)
                {
                    ComboBoxZone.ItemsSource = null;
                    ComboBoxZone.Items.Clear();
                    ComboBoxZone.ItemsSource = ds.Tables[0].DefaultView;
                    view = ds.Tables[0].DefaultView;

                    ComboBoxZone1.ItemsSource = null;
                    ComboBoxZone1.Items.Clear();
                    ComboBoxZone1.ItemsSource = ds1.Tables[0].DefaultView;
                    Animation.Visibility = Visibility.Collapsed;
                    ServerStack.Visibility = Visibility.Visible;
                    GenerateMetadata.Visibility = Visibility.Collapsed;
                    Output.Visibility = Visibility.Collapsed;

                    if (GenerateMetadata.IsChecked == true)
                    {
                        // button1.Visibility = Visibility.Collapsed;
                        // ReqButton.Visibility = Visibility.Collapsed;
                        //  Show_by_Report.Visibility = Visibility.Visible;
                        // CallGraphButton.Visibility = Visibility.Visible;
                    }
                }
            }
        }

        private async void CallDatabaseList(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(ResultText.Text) && ResultText2.Visibility == Visibility.Collapsed && ResultText3.Visibility == Visibility.Collapsed)
            {
                MessageBox.Show("Enter Valid Workspace connection");
            }
            else if ((String.IsNullOrEmpty(ResultText.Text) || String.IsNullOrEmpty(ResultText2.Text)) && ResultText2.Visibility == Visibility.Visible && ResultText3.Visibility == Visibility.Collapsed)
            {
                MessageBox.Show("Enter Valid Workspace connection");
            }
            else if (ResultText.Visibility == Visibility.Visible && ResultText2.Visibility == Visibility.Visible && ResultText3.Visibility == Visibility.Collapsed && ResultText.Text == ResultText2.Text)
            {
                MessageBox.Show("Workspace 1 and Workspace 2 looks similar. Try using different workspaces for better results.");
            }
            else if (ResultText.Text != "" && ResultText2.Visibility == Visibility.Collapsed && ResultText3.Visibility == Visibility.Collapsed)
            {
                Animation.Visibility = Visibility.Visible;
                ServerStack.Visibility = Visibility.Hidden;
                button1.Visibility = Visibility.Collapsed;
                ReqButton.Visibility = Visibility.Collapsed;
                Show_by_Report.Visibility = Visibility.Collapsed;
                CallGraphButton.Visibility = Visibility.Collapsed;
                GenerateMetadata.Visibility = Visibility.Collapsed;
                SignOutButton.Visibility = Visibility.Collapsed;
                Output.Visibility = Visibility.Collapsed;
                workspacename = ResultText.Text.ToString();
                backgroundWorker2.RunWorkerAsync();
            }
            else if (ResultText.Text != "" && ResultText2.Text != "" && ResultText3.Visibility == Visibility.Collapsed)
            {
                Animation.Visibility = Visibility.Visible;
                ServerStack.Visibility = Visibility.Hidden;
                button1.Visibility = Visibility.Collapsed;
                ReqButton.Visibility = Visibility.Collapsed;
                Show_by_Report.Visibility = Visibility.Collapsed;
                CallGraphButton.Visibility = Visibility.Collapsed;
                GenerateMetadata.Visibility = Visibility.Collapsed;
                SignOutButton.Visibility = Visibility.Collapsed;
                Output.Visibility = Visibility.Collapsed;
                workspacename = ResultText.Text.ToString();
                workspacename1 = ResultText2.Text.ToString();
                backgroundWorker2.RunWorkerAsync();

            }

        }

        private async void BindBox()
        {
            if (workspacename != "" && workspacename1 == "")
            {

                try
                {
                    AdomdConnection connection = new AdomdConnection();
                    connection.ConnectionString = GetConnectionStringForCombo(workspacename.ToString());
                    connection.Open();
                    string queryString = "SELECT DISTINCT [CATALOG_NAME] FROM $System.DBSCHEMA_CATALOGS;";
                    AdomdCommand cmd = connection.CreateCommand();
                    cmd.CommandText = queryString;
                    AdomdDataAdapter ad = new AdomdDataAdapter(queryString, connection);
                    ds = new DataSet();
                    ad.Fill(ds, "DBSCHEMA_CATALOGS");



                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    //MessageBox.Show("Connection String should be valid");
                }
            }
            else if (workspacename != "" && workspacename1 != "")
            {

                try
                {
                    AdomdConnection connection = new AdomdConnection();
                    connection.ConnectionString = GetConnectionStringForCombo(workspacename.ToString());
                    connection.Open();
                    string queryString = "SELECT DISTINCT [CATALOG_NAME] FROM $System.DBSCHEMA_CATALOGS;";
                    AdomdCommand cmd = connection.CreateCommand();
                    cmd.CommandText = queryString;
                    AdomdDataAdapter ad = new AdomdDataAdapter(queryString, connection);
                    ds = new DataSet();
                    ad.Fill(ds, "DBSCHEMA_CATALOGS");



                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    //MessageBox.Show("Connection String should be valid");
                }

                try
                {
                    AdomdConnection connection = new AdomdConnection();
                    connection.ConnectionString = GetConnectionStringForCombo(workspacename1.ToString());
                    connection.Open();
                    string queryString = "SELECT DISTINCT [CATALOG_NAME] FROM $System.DBSCHEMA_CATALOGS;";
                    AdomdCommand cmd = connection.CreateCommand();
                    cmd.CommandText = queryString;
                    AdomdDataAdapter ad = new AdomdDataAdapter(queryString, connection);
                    ds1 = new DataSet();
                    ad.Fill(ds1, "DBSCHEMA_CATALOGS");



                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    //MessageBox.Show("Connection String should be valid");
                }
            }



        }

        private async void run_cmd()
        {
            try
            {
                //System.Threading.Thread.Sleep(10000);
                string workingDirectory = Directory.GetCurrentDirectory() + @"\PythonFile";
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        RedirectStandardInput = true,
                        UseShellExecute = false,
                        RedirectStandardError = true,
                        CreateNoWindow = true,
                        WorkingDirectory = workingDirectory
                    }


                };
                process.Start();
                // System.Threading.Thread.Sleep(10000);
                // string TextPython = "C:\\Users\\Rakesh.P\\Anaconda3";
                using (var sw = process.StandardInput)
                {
                    if (sw.BaseStream.CanWrite)
                    {

                        sw.WriteLine(PythonPath1 + @"\Scripts\activate.bat");
                        sw.WriteLine("python " + '"' + workingDirectory + @"\PowerBI_Process_Python.py" + '"');
                    }
                }
                //string output = process.StandardOutput.ReadToEnd();

                process.WaitForExit();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }


        }

        public async Task<string> GetHttpContentWithTokenCombo(string url, string token)
        {
            Animation.Visibility = Visibility.Visible;
            ServerStack.Visibility = Visibility.Hidden;
            //StackGrid.Visibility = Visibility.Hidden;
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
            Animation.Visibility = Visibility.Collapsed;
            StackGrid.Visibility = Visibility.Visible;
            ServerStack.Visibility = Visibility.Visible;
        }

        private async void button1_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = new DataTable();


            WindowMainName.Height = 766;
            //dataGrid1.Visibility = Visibility.Collapsed;
            StackGrid.Visibility = Visibility.Collapsed;
            //ScrollViewer.Visibility = Visibility.Collapsed;
            Animation.Visibility = Visibility.Visible;
            ServerStack.Visibility = Visibility.Hidden;
            MessageBox.Show("Generating Report. Please Wait ....");
            // string file = @"Metadata Output.pbix";
            string fileName = "BI4BI - Power BI.pbix";
            string path = System.IO.Path.Combine(Environment.CurrentDirectory, @"Report\", fileName);
            Process.Start(path);
            //MessageBox.Show(path);

            /* dataGrid1.SelectAll();
             dataGrid1.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;

             ApplicationCommands.Copy.Execute(null, dataGrid1);
             dataGrid1.UnselectAll();
             try
             {

                 Microsoft.Office.Interop.Excel.Application excelApp;
                 Microsoft.Office.Interop.Excel.Workbook excelWkbk;
                 Microsoft.Office.Interop.Excel.Worksheet excelWksht;
                 object misValue = System.Reflection.Missing.Value;
                 excelApp = new Microsoft.Office.Interop.Excel.Application();
                 excelApp.Visible = true;
                 excelWkbk = excelApp.Workbooks.Add(misValue);
                 excelWksht = (Microsoft.Office.Interop.Excel.Worksheet)excelWkbk.Worksheets.get_Item(1);
                 Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)excelWksht.Cells[1, 1];
                 CR.Select();
                 excelWksht.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);


                 /*StreamWriter swObj = new StreamWriter("Metadata.csv");
                 swObj.WriteLine(Clipboardresult);
                 swObj.Close();
                 Process.Start("Metadata.csv")
             }
             catch
             {
                 MessageBox.Show("Please close the Existing file and try again");
             };*/
            Animation.Visibility = Visibility.Collapsed;
            StackGrid.Visibility = Visibility.Visible;
            ServerStack.Visibility = Visibility.Visible;
            // StackGrid.Visibility = Visibility.Visible;
            //ScrollViewer.Visibility = Visibility.Visible;
            //dataGrid1.Visibility = Visibility.Visible;
        }

        private async void copyDataGridContentToClipboard()
        {
            WindowMainName.Height = 766;
            // dataGrid1.Visibility = Visibility.Collapsed;
            StackGrid.Visibility = Visibility.Collapsed;
            // ScrollViewer.Visibility = Visibility.Collapsed;
            Animation.Visibility = Visibility.Visible;
            ServerStack.Visibility = Visibility.Hidden;

            //  dataGrid1.SelectAll();
            //  dataGrid1.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;

            //ApplicationCommands.Copy.Execute(null, dataGrid1);
            // dataGrid1.UnselectAll();
        }
        public string RenderControl(System.Web.UI.Control ctrl)
        {
            Animation.Visibility = Visibility.Visible;
            ServerStack.Visibility = Visibility.Hidden;
            //StackGrid.Visibility = Visibility.Hidden;
            StringBuilder sb = new StringBuilder();
            StringWriter tw = new StringWriter(sb);
            HtmlTextWriter hw = new HtmlTextWriter(tw);

            ctrl.RenderControl(hw);
            return sb.ToString();
            Animation.Visibility = Visibility.Collapsed;
            StackGrid.Visibility = Visibility.Visible;
            ServerStack.Visibility = Visibility.Visible;
        }

        public void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.ShowDialog();
            TextPython.Text = dialog.SelectedPath;
            PythonPath1 = TextPython.Text;

        }

        private void backgroundWorker3_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {

            // This is where the processor intensive code should go
            //ShowSelectedReport();
            ExecuteMethodAsync();
            //Python code to be added here


            //Python code to be ended here
            // if we need any output to be used, put it in the DoWorkEventArgs object
            e.Result = "all done";
            //If the process exits the loop, ensure that progress is set to 100%
            //Remember in the loop we set i < 100 so in theory the process will complete at 99%
            backgroundWorker1.ReportProgress(100);
        }
        private void backgroundWorker3_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker3_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {

            }
            else if (e.Error != null)
            {

            }
            else
            {
                if (ReportCnt > 0 || ColumnsCnt > 0 || CalcCnt > 0)
                {

                    Animation.Visibility = Visibility.Collapsed;
                    ServerStack.Visibility = Visibility.Visible;
                    button1.Visibility = Visibility.Visible;
                    ReqButton.Visibility = Visibility.Visible;
                    Show_by_Report.Visibility = Visibility.Visible;
                    CallGraphButton.Visibility = Visibility.Visible;
                    SignOutButton.Visibility = Visibility.Visible;
                    GenerateMetadata.Visibility = Visibility.Collapsed;
                    Output.Visibility = Visibility.Collapsed;
                    Output.IsEnabled = true;
                    GenerateMetadata.IsEnabled = false;
                    Output.IsChecked = true;
                    ProcessImage.Visibility = Visibility.Hidden;
                    OutputImage.Visibility = Visibility.Hidden;
                    //StackGrid.Visibility = Visibility.Hidden;
                    //dataGrid1.Visibility = Visibility.Collapsed;

                    MetadataToolTip.Text = "Please Find the summary of items inserted into the server " + serverlabel.ToString();
                    MetadataToolTip.AppendText(Environment.NewLine);
                    MetadataToolTip.AppendText("Number of Reports = " + ReportCnt + "\r\n");
                    MetadataToolTip.AppendText("Number of Columns = " + ColumnsCnt + "\r\n");
                    MetadataToolTip.AppendText("Number of Calculations = " + CalcCnt + "\r\n");

                    OutputToolTip.Text = "Generate Power BI Report - The generated metadata is presented in a read-able format in a Power BI Report\r\n";
                    OutputToolTip.AppendText("Requirement Document Generator - Generate Requirement Document for easier hand-over which will help in migration");


                }
                else
                {

                    Animation.Visibility = Visibility.Collapsed;
                    ServerStack.Visibility = Visibility.Visible;
                    button1.Visibility = Visibility.Visible;
                    ReqButton.Visibility = Visibility.Visible;
                    Show_by_Report.Visibility = Visibility.Visible;
                    CallGraphButton.Visibility = Visibility.Visible;
                    SignOutButton.Visibility = Visibility.Visible;
                    GenerateMetadata.Visibility = Visibility.Collapsed;
                    Output.Visibility = Visibility.Collapsed;
                    Output.IsEnabled = false;
                    //   GenerateMetadata.IsEnabled = false;
                    //  GenerateMetadata.IsChecked = false;

                    MessageBox.Show("Issues Found in the Metadata Process. Please contact the administrator for further clarification");

                }
            }
        }
        private async void Show_By_Report(object sender, RoutedEventArgs e)
        {
            button1.Visibility = Visibility.Visible;
            ReqButton.Visibility = Visibility.Visible;
            Show_by_Report.Visibility = Visibility.Visible;
            CallGraphButton.Visibility = Visibility.Visible;
            SignOutButton.Visibility = Visibility.Visible;
            //MessageBox.Show("");
            itemCombo = items;
            workspacename = ResultText.Text.ToString();
            serverlabel = Server.Text.ToString();
            if (String.IsNullOrEmpty(serverlabel.ToString()))
            {
                                        
                MessageBox.Show("Enter the mandatory fields");
                Animation.Visibility = Visibility.Collapsed;
                ServerStack.Visibility = Visibility.Visible;
                button1.Visibility = Visibility.Visible;
                ReqButton.Visibility = Visibility.Visible;
                Show_by_Report.Visibility = Visibility.Visible;
                CallGraphButton.Visibility = Visibility.Visible;
                SignOutButton.Visibility = Visibility.Visible;
                GenerateMetadata.Visibility = Visibility.Collapsed;
                Output.Visibility = Visibility.Collapsed;
                Output.IsEnabled = false;
                //GenerateMetadata.IsEnabled = true;
                // GenerateMetadata.IsChecked = false;

            }
            else if (items == null )
                {
                    MessageBox.Show("Please select atleast one report from the reports dropdown");
                    Animation.Visibility = Visibility.Collapsed;
                    ServerStack.Visibility = Visibility.Visible;
                    button1.Visibility = Visibility.Visible;
                    ReqButton.Visibility = Visibility.Visible;
                    Show_by_Report.Visibility = Visibility.Visible;
                    CallGraphButton.Visibility = Visibility.Visible;
                    SignOutButton.Visibility = Visibility.Visible;
                    GenerateMetadata.Visibility = Visibility.Collapsed;
                    Output.Visibility = Visibility.Collapsed;
                    Output.IsEnabled = false;
            }
                else
                {
                Animation.Visibility = Visibility.Visible;
                ServerStack.Visibility = Visibility.Hidden;
                button1.Visibility = Visibility.Collapsed;
                ReqButton.Visibility = Visibility.Collapsed;
                Show_by_Report.Visibility = Visibility.Collapsed;
                CallGraphButton.Visibility = Visibility.Collapsed;
                SignOutButton.Visibility = Visibility.Collapsed;
                GenerateMetadata.Visibility = Visibility.Collapsed;
                Output.Visibility = Visibility.Collapsed;
                ComboBoxZone.Text = "";
                WindowMainName.Height = 766;
                backgroundWorker3.RunWorkerAsync();
            }
            }

        

        private async void ShowSelectedReport()
        {

            try
            {
                string query = "";

                // StackGrid.Visibility = Visibility.Hidden;
                // WindowMainName.Height = 766;
                //TokenInfoText.Text = "";
                DataTable dt = new DataTable();
                AdomdConnection connection = new AdomdConnection();
                connection.ConnectionString = GetConnectionString(workspacename.ToString(), itemCombo.ToString());
                connection.Open();
                string queryString = "";
                //DataTable dtUsage = new DataTable();
                DataTable dtUsage = new DataTable();
                DataTable dtUsage1 = new DataTable();
                DataTable dtUsage2 = new DataTable();
                DataTable dtUsage3 = new DataTable();
                DataTable dtUsage4 = new DataTable();
                DataTable dtUsage5 = new DataTable();
                DataTable dtUsage6 = new DataTable();
                DataTable dtUsage7 = new DataTable();
                DataTable dtUsage8 = new DataTable();

                //  ComboBoxZone.DataContext = null;
                // ComboBoxZone.ItemsSource = null;
                //ComboBoxZone.Text = "";
                //   Animation.Visibility = Visibility.Visible;
                // ServerStack.Visibility = Visibility.Hidden;
                //StackGrid.Visibility = Visibility.Hidden;
                //items = new string[ComboBoxZone.Items.Count];
                //AdomdConnection connection = new AdomdConnection();
                // connection.ConnectionString = GetConnectionString(ResultText.Text, item.Row[0].ToString());
                // connection.Open();
                //MessageBox.Show(item.ToString());  
                //DataTable dt = new DataTable();




                int pos = workspacename.ToString().LastIndexOf("/") + 1;
                // WorkspaceLabel.Content = "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]";
                queryString = "SELECT DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace], [CATALOG_NAME] AS [Report Name], [DIMENSION_UNIQUE_NAME] AS [Dataset Name], LEVEL_CAPTION AS [Column Name] FROM $System.MDSchema_levels WHERE CUBE_NAME  ='Model' AND level_origin=2 AND LEVEL_NAME <> '(All)' order by [DIMENSION_UNIQUE_NAME]   ";
                //queryString = check(query);
                AdomdCommand cmd = connection.CreateCommand();
                cmd.CommandText = queryString;
                AdomdDataAdapter ad = new AdomdDataAdapter(queryString, connection);
                ad.Fill(dt);


                DataTable dt2 = new DataTable();
                string queryString1 = "select DISTINCT" + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace], DATABASE_NAME as [Report Name],'['+[TABLE]+']' AS [Dataset Name],OBJECT AS [Column Name],EXPRESSION AS [Calculated Column Expression] from $SYSTEM.DISCOVER_CALC_DEPENDENCY WHERE OBJECT_TYPE = 'CALC_COLUMN' ";
                AdomdCommand cmd1 = connection.CreateCommand();
                cmd1.CommandText = queryString1;
                AdomdDataAdapter ad1 = new AdomdDataAdapter(queryString1, connection);
                ad1.Fill(dt2);

                dt2.PrimaryKey = new DataColumn[] {
                    dt2.Columns["Report Name"],dt2.Columns["Dataset Name"],dt2.Columns["Column Name"] };


                dt.Merge(dt2);
                //  dt.DefaultView.Sort = "Dataset Name ASC";

                DataTable dt4 = new DataTable();
                string queryString3 = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace],  DATABASE_NAME as [Report Name],'['+[TABLE]+']' AS [Dataset Name],OBJECT AS [Column Name],EXPRESSION AS [Calculated Measure Expression] from $SYSTEM.DISCOVER_CALC_DEPENDENCY WHERE OBJECT_TYPE = 'MEASURE' ";
                AdomdCommand cmd3 = connection.CreateCommand();
                cmd3.CommandText = queryString3;
                AdomdDataAdapter ad3 = new AdomdDataAdapter(queryString3, connection);
                ad3.Fill(dt4);



                dt.Merge(dt4);
                //   dt.DefaultView.Sort = "Dataset Name ASC";




                DataTable dt3 = new DataTable();
                string queryString2 = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace],DATABASE_NAME as [Report Name],'['+[TABLE]+']' AS [Dataset Name],OBJECT AS [Column Name],EXPRESSION AS [Calculated Table Expression] from $SYSTEM.DISCOVER_CALC_DEPENDENCY WHERE OBJECT_TYPE = 'CALC_TABLE' ";
                AdomdCommand cmd2 = connection.CreateCommand();
                cmd2.CommandText = queryString2;
                AdomdDataAdapter ad2 = new AdomdDataAdapter(queryString2, connection);
                ad2.Fill(dt3);

                dt3.PrimaryKey = new DataColumn[] {
                    dt3.Columns["Report Name"],dt3.Columns["Dataset Name"],dt3.Columns["Column Name"] };


                dt.Merge(dt3);

                // dt.DefaultView.Sort = "DatasetName ASC";


                pos = workspacename.ToString().LastIndexOf("/") + 1;


                dt3 = new DataTable();
                queryString2 = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],TableID,QueryDefinition FROM $SYSTEM.TMSCHEMA_PARTITIONS ";
                cmd2 = connection.CreateCommand();
                cmd2.CommandText = queryString2;
                ad2 = new AdomdDataAdapter(queryString2, connection);
                ad2.Fill(dt3);


                dt4 = new DataTable();
                queryString3 = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],[ID] as [TableID],'['+[Name]+']' as [Table Name] FROM $SYSTEM.TMSCHEMA_TABLES ";
                AdomdCommand cmd4 = connection.CreateCommand();
                cmd4.CommandText = queryString3;
                AdomdDataAdapter ad4 = new AdomdDataAdapter(queryString3, connection);
                ad4.Fill(dt4);


                var JoinResult = (from p in dt3.AsEnumerable()
                                  join t in dt4.AsEnumerable()
                                  on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("TableID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("TableID") } into ps
                                  from tnew in ps.DefaultIfEmpty()
                                  select new
                                  {
                                      WorkspaceName = p.Field<string>("Workspace"),
                                      ReportName = p.Field<string>("Report Name"),
                                      TableName = tnew.Field<string>("Table Name"),
                                      //Query1 = p.Field<string>("QueryDefinition").Trim().Replace(" ","").Replace(@"\r\n?|\n",""),
                                      // Query2 = findNthOccur(p.Field<string>("QueryDefinition"),'"',2),
                                      QueryDef = p.Field<string>("QueryDefinition"),
                                      //Check1= p.Field<string>("QueryDefinition").IndexOf("Item") > 0  && p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).IndexOf(".")>0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : p.Field<string>("QueryDefinition").IndexOf("Item") > 0  ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 6) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) - 1) + "." + p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : "",
                                      Source = p.Field<string>("QueryDefinition").IndexOf("Database") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 2) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) - 1) : "File Source/Derived Table",
                                      Path = p.Field<string>("QueryDefinition").IndexOf("Contents") > 0 || p.Field<string>("QueryDefinition").IndexOf("Files") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 2) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 1) - 1) : p.Field<string>("QueryDefinition").IndexOf("Database") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 3) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 4) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 3) - 1) : p.Field<string>("QueryDefinition").IndexOf("Table.NestedJoin") > 0 || p.Field<string>("QueryDefinition").IndexOf("Table.FromRows") > 0 ? "Derived Table inside PBI" : "No Database or Path available",
                                      Query = p.Field<string>("QueryDefinition").IndexOf("Query=") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 6) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) - 1).Replace("#(lf)", "") : p.Field<string>("QueryDefinition").IndexOf("NativeQuery") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).Replace("#(lf)", "") : "No Query Available",
                                      DatabaseItem = p.Field<string>("QueryDefinition").IndexOf("Item") > 0 && p.Field<string>("QueryDefinition").IndexOf("Contents") <= 0 && p.Field<string>("QueryDefinition").IndexOf("Query") <= 0 && p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).IndexOf(".") > 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : p.Field<string>("QueryDefinition").IndexOf("Item") > 0 && p.Field<string>("QueryDefinition").IndexOf("Contents") <= 0 && p.Field<string>("QueryDefinition").IndexOf("Query") <= 0 && p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1).IndexOf(".") <= 0 ? p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 6) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 5) - 1) + "." + p.Field<string>("QueryDefinition").Substring(findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) + 1, findNthOccur(p.Field<string>("QueryDefinition"), '"', 8) - findNthOccur(p.Field<string>("QueryDefinition"), '"', 7) - 1) : "No Database Item available",
                                  }).ToList();

                dt4 = LINQResultToDataTable(JoinResult);
                dt4.Columns["WorkspaceName"].ColumnName = "Workspace";
                dt4.Columns["ReportName"].ColumnName = "Report Name";
                dt4.Columns["TableName"].ColumnName = "Dataset Name";
                dt4.Columns["Source"].ColumnName = "Source";
                dt4.Columns["Path"].ColumnName = "Database Or Path";
                // dt4.Columns["Query"].ColumnName = "Advance Editor Steps";


                var JoinResult1 = (from p in dt.AsEnumerable()
                                   join t in dt4.AsEnumerable()
                                   on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<string>("Dataset Name") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<string>("Dataset Name") } into ps
                                   from tnew in ps
                                   select new
                                   {

                                       WorkspaceName = p.Field<string>("Workspace"),
                                       DatasetName = p.Field<string>("Dataset Name"),
                                       ReportName = p.Field<string>("Report Name"),
                                       ColumnName = p.Field<string>("Column Name"),
                                       Source = tnew == null ? "" : tnew.Field<string>("Source"),
                                       Path = tnew == null ? "" : tnew.Field<string>("Database Or Path"),
                                       Query = tnew == null ? "" : tnew.Field<string>("Query"),
                                       DatabaseItem = tnew == null ? "" : tnew.Field<string>("DatabaseItem"),
                                       // Check1= tnew == null ? "" : tnew.Field<string>("Check1"),
                                       Steps = tnew == null ? "" : tnew.Field<string>("QueryDef")
                                       //Check= tnew == null ? "" : tnew.Field<string>("Check")

                                   }).ToList();

                dt4 = LINQResultToDataTable(JoinResult1);
                dt4.Columns["WorkspaceName"].ColumnName = "Workspace";
                dt4.Columns["ReportName"].ColumnName = "Report Name";
                dt4.Columns["ColumnName"].ColumnName = "Column Name";
                dt4.Columns["DatasetName"].ColumnName = "Dataset Name";
                dt4.Columns["Source"].ColumnName = "Source";
                dt4.Columns["Path"].ColumnName = "Database Or Path";
                //dt4.Columns["Query"].ColumnName = "Advance Editor Steps";
                dt4.PrimaryKey = new DataColumn[] {
                    dt4.Columns["Report Name"],dt4.Columns["Dataset Name"],dt4.Columns["Column Name"] };

                dt.PrimaryKey = new DataColumn[] {
                    dt.Columns["Report Name"],dt.Columns["Dataset Name"],dt.Columns["Column Name"] };

                dt.Merge(dt4);
                /*dt.Columns["WorkspaceName"].ColumnName = "Workspace";
                dt.Columns["ReportName"].ColumnName = "Report Name";
                dt.Columns["ColumnName"].ColumnName = "Column Name";
                dt.Columns["DatasetName"].ColumnName = "Dataset Name";*/
                //dt.Columns["Source"].ColumnName = "Source";
                //dt.Columns["Path"].ColumnName = "Database Or Path";
                //dt.Columns["Query"].ColumnName = "Advance Editor Steps";

                //dt.DefaultView.Sort = "DatasetName ASC";



                pos = workspacename.ToString().LastIndexOf("/") + 1;


                dt3 = new DataTable();
                queryString2 = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],FromTableID,FromColumnID,ToTableID,ToColumnID,RefreshedTime FROM $SYSTEM.TMSCHEMA_RELATIONSHIPS ";
                cmd2 = connection.CreateCommand();
                cmd2.CommandText = queryString2;
                ad2 = new AdomdDataAdapter(queryString2, connection);
                ad2.Fill(dt3);

                if (dt3.Rows.Count > 0)
                {



                    DataTable dt4Master = new DataTable();
                    string queryStringMaster = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],[ID] AS [Dataset ID] ,'['+[Name]+']'  AS [Dataset Name] FROM $SYSTEM.TMSCHEMA_TABLES";
                    AdomdCommand cmd4Master = connection.CreateCommand();
                    cmd4Master.CommandText = queryStringMaster;
                    AdomdDataAdapter ad4Master = new AdomdDataAdapter(queryStringMaster, connection);
                    ad4Master.Fill(dt4Master);

                    DataTable dt4ColumnMaster = new DataTable();
                    string queryStringColumnMaster = "select DISTINCT " + "'" + workspacename.ToString().Substring(pos, workspacename.ToString().Length - pos).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],[ID] AS [Column ID],ExplicitName AS [Column Name],InferredName FROM $SYSTEM.TMSCHEMA_COLUMNS";
                    AdomdCommand cmd4ColumnMaster = connection.CreateCommand();
                    cmd4ColumnMaster.CommandText = queryStringColumnMaster;
                    AdomdDataAdapter ad4ColumnMaster = new AdomdDataAdapter(queryStringColumnMaster, connection);
                    ad4ColumnMaster.Fill(dt4ColumnMaster);

                    //MessageBox.Show(dt3.Columns["RefreshedTime"].DataType.ToString());


                    var JoinResult4 = (from p in dt3.AsEnumerable()
                                       join t in dt4Master.AsEnumerable()
                                       on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("FromTableID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Dataset ID") } into ps
                                       from tnew in ps.DefaultIfEmpty()
                                       select new
                                       {
                                           WorkspaceName = p.Field<string>("Workspace"),
                                           ReportName = p.Field<string>("Report Name"),
                                           FromTableID = p.Field<System.UInt64>("FromTableID"),
                                           ToTableID = p.Field<System.UInt64>("ToTableID"),
                                           FromColumnID = p.Field<System.UInt64>("FromColumnID"),
                                           ToColumnID = p.Field<System.UInt64>("ToColumnID"),
                                           RefreshedTime = p.Field<System.DateTime>("RefreshedTime"),
                                           FromTableName = tnew.Field<string>("Dataset Name")

                                       }).ToList();

                    dt3 = LINQResultToDataTable(JoinResult4);
                    dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                    dt3.Columns["ReportName"].ColumnName = "Report Name";
                    dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                    dt3.Columns["ToTableID"].ColumnName = "To Table ID";
                    dt3.Columns["FromTableID"].ColumnName = "From Table ID";

                    var JoinResult2 = (from p in dt3.AsEnumerable()
                                       join t in dt4Master.AsEnumerable()
                                       on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("To Table ID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Dataset ID") } into ps
                                       from tnew in ps.DefaultIfEmpty()
                                       select new
                                       {
                                           WorkspaceName = p.Field<string>("Workspace"),
                                           ReportName = p.Field<string>("Report Name"),
                                           FromTableName = p.Field<string>("From Table Name"),
                                           FromColumnID = p.Field<System.UInt64>("FromColumnID"),
                                           ToTableName = tnew.Field<string>("Dataset Name"),
                                           ToColumnID = p.Field<System.UInt64>("ToColumnID"),
                                           RefreshedTime = p.Field<System.DateTime>("RefreshedTime")

                                       }).ToList();
                    dt3 = LINQResultToDataTable(JoinResult2);
                    dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                    dt3.Columns["ReportName"].ColumnName = "Report Name";
                    dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                    dt3.Columns["FromColumnID"].ColumnName = "From Column ID";
                    dt3.Columns["ToTableName"].ColumnName = "To Table Name";
                    dt3.Columns["ToColumnID"].ColumnName = "To Column ID";

                    var JoinResult3 = (from p in dt3.AsEnumerable()
                                       join t in dt4ColumnMaster.AsEnumerable()
                                       on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("From Column ID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Column ID") } into ps
                                       from tnew in ps.DefaultIfEmpty()
                                       select new
                                       {
                                           WorkspaceName = p.Field<string>("Workspace"),
                                           ReportName = p.Field<string>("Report Name"),
                                           FromTableName = p.Field<string>("From Table Name"),
                                           FromColumnID = p.Field<System.UInt64>("From Column ID"),
                                           FromColumnName = tnew.Field<string>("Column Name") == null ? tnew.Field<string>("InferredName") : tnew.Field<string>("Column Name"),
                                           ToTableName = p.Field<string>("To Table Name"),
                                           ToColumnID = p.Field<System.UInt64>("To Column ID"),
                                           RefreshedTime = p.Field<System.DateTime>("RefreshedTime")

                                       }).ToList();
                    dt3 = LINQResultToDataTable(JoinResult3);
                    dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                    dt3.Columns["ReportName"].ColumnName = "Report Name";
                    dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                    dt3.Columns["FromColumnID"].ColumnName = "From Column ID";
                    dt3.Columns["FromColumnName"].ColumnName = "From Column Name";
                    dt3.Columns["ToTableName"].ColumnName = "To Table Name";
                    dt3.Columns["ToColumnID"].ColumnName = "To Column ID";

                    var JoinResultTemp = (from p in dt3.AsEnumerable()
                                          join t in dt4ColumnMaster.AsEnumerable()
                                          on new { X0 = p.Field<string>("Workspace"), X1 = p.Field<string>("Report Name"), X2 = p.Field<System.UInt64>("To Column ID") } equals new { X0 = t.Field<string>("Workspace"), X1 = t.Field<string>("Report Name"), X2 = t.Field<System.UInt64>("Column ID") } into ps
                                          from tnew in ps.DefaultIfEmpty()
                                          select new
                                          {
                                              WorkspaceName = p.Field<string>("Workspace"),
                                              ReportName = p.Field<string>("Report Name"),
                                              DatasetName = p.Field<string>("From Table Name"),
                                              ColumnName = p.Field<string>("From Column Name"),
                                              FromTableName = p.Field<string>("From Table Name"),
                                              FromColumnName = p.Field<string>("From Column Name"),
                                              ToTableName = p.Field<string>("To Table Name"),
                                              ToColumnName = tnew.Field<string>("Column Name") == null ? tnew.Field<string>("InferredName") : tnew.Field<string>("Column Name"),
                                              RefreshedTime = p.Field<System.DateTime>("RefreshedTime")



                                          }).ToList();
                    dt3 = LINQResultToDataTable(JoinResultTemp);

                    dt3.Columns["WorkspaceName"].ColumnName = "Workspace";
                    dt3.Columns["ReportName"].ColumnName = "Report Name";
                    dt3.Columns["ColumnName"].ColumnName = "Column Name";
                    dt3.Columns["DatasetName"].ColumnName = "Dataset Name";
                    dt3.Columns["FromTableName"].ColumnName = "From Table Name";
                    dt3.Columns["FromColumnName"].ColumnName = "From Column Name";
                    dt3.Columns["ToTableName"].ColumnName = "To Table Name";
                    dt3.Columns["ToColumnName"].ColumnName = "To Column Name";
                    dt3.Columns["RefreshedTime"].ColumnName = "Refreshed Time";

                    dt.Merge(dt3);

                }




                int posUsage = workspacename.ToString().LastIndexOf("/") + 1;


                string queryUsage = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME,COLUMN_ID,ATTRIBUTE_NAME AS COLUMN_NAME,DATATYPE AS [Data Type],DICTIONARY_SIZE AS DICTIONARY_SIZE_BYTES,COLUMN_ENCODING AS COLUMN_ENCODING_INT from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMNS WHERE COLUMN_TYPE='BASIC_DATA' ";
                AdomdCommand cmdUsage = connection.CreateCommand();
                cmdUsage.CommandText = queryUsage;
                AdomdDataAdapter ad4Usage = new AdomdDataAdapter(queryUsage, connection);
                ad4Usage.Fill(dtUsage);




                string queryUsage1 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME,COLUMN_ID AS STRUCTURE_NAME,USED_SIZE,TABLE_ID AS HIERARCHY_ID from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE LEFT( TABLE_ID,2 )='U$' ";
                AdomdCommand cmdUsage1 = connection.CreateCommand();
                cmdUsage1.CommandText = queryUsage1;
                AdomdDataAdapter ad4Usage1 = new AdomdDataAdapter(queryUsage1, connection);
                ad4Usage1.Fill(dtUsage1);





                string queryUsage2 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME,COLUMN_ID AS STRUCTURE_NAME,SEGMENT_NUMBER,TABLE_PARTITION_NUMBER,USED_SIZE,TABLE_ID AS COLUMN_HIERARCHY_ID from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE LEFT( TABLE_ID,2 )='H$' ";
                AdomdCommand cmdUsage2 = connection.CreateCommand();
                cmdUsage2.CommandText = queryUsage2;
                AdomdDataAdapter ad4Usage2 = new AdomdDataAdapter(queryUsage2, connection);
                ad4Usage2.Fill(dtUsage2);




                string queryUsage3 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME, PARTITION_NAME,COLUMN_ID AS COLUMN_NAME , SEGMENT_NUMBER,TABLE_PARTITION_NUMBER,RECORDS_COUNT AS SEGMENT_ROWS,USED_SIZE,COMPRESSION_TYPE,BITS_COUNT,BOOKMARK_BITS_COUNT,VERTIPAQ_STATE from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE RIGHT(LEFT( TABLE_ID,2 ),1)<>'$' ";
                AdomdCommand cmdUsage3 = connection.CreateCommand();
                cmdUsage3.CommandText = queryUsage3;
                AdomdDataAdapter ad4Usage3 = new AdomdDataAdapter(queryUsage3, connection);
                ad4Usage3.Fill(dtUsage3);




                string queryUsage4 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],DIMENSION_NAME AS TABLE_NAME, TABLE_ID AS RELATIONSHIP_ID,USED_SIZE from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMN_SEGMENTS WHERE  LEFT( TABLE_ID,2 )='R$' ";
                AdomdCommand cmdUsage4 = connection.CreateCommand();
                cmdUsage4.CommandText = queryUsage4;
                AdomdDataAdapter ad4Usage4 = new AdomdDataAdapter(queryUsage4, connection);
                ad4Usage4.Fill(dtUsage4);




                string queryUsage5 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],[NAME] AS TABLE_NAME,[RefreshedTime] FROM  $SYSTEM.TMSCHEMA_PARTITIONS  ";
                AdomdCommand cmdUsage5 = connection.CreateCommand();
                cmdUsage5.CommandText = queryUsage5;
                AdomdDataAdapter ad4Usage5 = new AdomdDataAdapter(queryUsage5, connection);
                ad4Usage5.Fill(dtUsage5);




                string queryUsage6 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],[ID] AS [Table ID],[Name] AS [Table Name] FROM  $SYSTEM.TMSCHEMA_TABLES ";
                AdomdCommand cmdUsage6 = connection.CreateCommand();
                cmdUsage6.CommandText = queryUsage6;
                AdomdDataAdapter ad4Usage6 = new AdomdDataAdapter(queryUsage6, connection);
                ad4Usage6.Fill(dtUsage6);




                string queryUsage7 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],TABLEID AS [Table ID], [ID] AS [Column ID],ExplicitName AS [Column Name] FROM $SYSTEM.TMSCHEMA_COLUMNS ";
                AdomdCommand cmdUsage7 = connection.CreateCommand();
                cmdUsage7.CommandText = queryUsage7;
                AdomdDataAdapter ad4Usage7 = new AdomdDataAdapter(queryUsage7, connection);
                ad4Usage7.Fill(dtUsage7);




                string queryUsage8 = "select DISTINCT " + "'" + workspacename.ToString().Substring(posUsage, workspacename.ToString().Length - posUsage).Replace("%20", " ").Replace("'", "''").Replace("\"", "") + "' AS [Workspace]," + "'" + itemCombo.ToString() + "' as [Report Name],[ID] AS [Relationship ID],[FromTableID],[FromColumnID],[FromCardinality],[ToTableID],[ToColumnID],[ToCardinality],[IsActive],CrossFilteringBehavior FROM $System.TMSCHEMA_RELATIONSHIPS";
                AdomdCommand cmdUsage8 = connection.CreateCommand();
                cmdUsage8.CommandText = queryUsage8;
                AdomdDataAdapter ad4Usage8 = new AdomdDataAdapter(queryUsage8, connection);
                ad4Usage8.Fill(dtUsage8);








                createsqltable(dt, "Metadata");

                createsqltableUsage(dtUsage, "Dictionary_Usage");
                createsqltableUsage(dtUsage1, "User_Hierarchy");
                createsqltableUsage(dtUsage2, "Hierarchy");
                createsqltableUsage(dtUsage3, "Data_Size");
                createsqltableUsage(dtUsage4, "Relationships_Size");
                createsqltableUsage(dtUsage5, "Last_Update");
                createsqltableUsage(dtUsage6, "TMSchema_Table");
                createsqltableUsage(dtUsage7, "TMSchema_Columns");
                createsqltableUsage(dtUsage8, "TMSchema_Relationships");



                SqlConnection SQLConnection = new SqlConnection();
                SQLConnection.ConnectionString = "Data Source =" + serverlabel.ToString() + "; Initial Catalog =Power BI Metadata; " + "Integrated Security=true;";

                string script = File.ReadAllText(@"C:\Users\Rakesh.P\B4Bi-V1\GetMetaData\GetMetaData\GetMetaData\GetMetaData\bin\Debug\Scripts\vw_Metadata.sql");

                SqlCommand cmdView = new SqlCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = script;
                SQLConnection.Open();
                cmd.ExecuteNonQuery();
                SQLConnection.Close();

                string QueryReport = "select count(DISTINCT [Report Name]) from dbo.Metadata";
                //Execute Queries and save results into variables
                SqlCommand CmdCntReport = SQLConnection.CreateCommand();
                CmdCntReport.CommandText = QueryReport;
                SQLConnection.Open();
                ReportCnt = (Int32)CmdCntReport.ExecuteScalar();
                SQLConnection.Close();


                string QueryColumns = "select count(DISTINCT [Column Name]) from dbo.Metadata";
                //Execute Queries and save results into variables
                SqlCommand CmdCntColumns = SQLConnection.CreateCommand();
                CmdCntColumns.CommandText = QueryColumns;
                SQLConnection.Open();

                ColumnsCnt = (Int32)CmdCntColumns.ExecuteScalar();
                SQLConnection.Close();

                string QueryCalc = "SELECT SUM([Calc 1]) FROM ";
                QueryCalc += "\n (";
                QueryCalc += "\n select COUNT(DISTINCT [Calculated Column Expression]) [Calc 1] from dbo.Metadata";
                QueryCalc += "\n where [Calculated Column Expression] is not null";
                QueryCalc += "\n UNION ALL ";
                QueryCalc += "\n select COUNT(DISTINCT [Calculated Measure Expression]) [Calc 2] from dbo.Metadata";
                QueryCalc += "\n where [Calculated Measure Expression] is not null";
                QueryCalc += "\n UNION ALL";
                QueryCalc += "\n select COUNT(DISTINCT [Calculated Table Expression]) [Calc 3] from dbo.Metadata";
                QueryCalc += "\n where [Calculated Table Expression] is not null";
                QueryCalc += "\n ) A";
                //Execute Queries and save results into variables
                SqlCommand CmdCntCalc = SQLConnection.CreateCommand();
                CmdCntCalc.CommandText = QueryCalc;
                SQLConnection.Open();
                CalcCnt = (Int32)CmdCntCalc.ExecuteScalar();
                SQLConnection.Close();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
            }




        }



        void dataGrid_AutoGeneratingColumn(object sender,
                                               DataGridAutoGeneratingColumnEventArgs e)
        {
            e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
        }

        private void AddServer_Click(object sender, RoutedEventArgs e)
        {
            //ServerStack.Height = 500;
            ResultText2.Visibility = Visibility.Visible;

            Server2Bord.Visibility = Visibility.Visible;
            Show_by_Report.Visibility = Visibility.Collapsed;
            WrapCheck.Visibility = Visibility.Collapsed;
            BorderBox.Visibility = Visibility.Collapsed;
            RemoveServer.Visibility = Visibility.Visible;
            AddServer.Visibility = Visibility.Collapsed;

            //AddServer2.Visibility = Visibility.Visible;

            RemoveServer.Visibility = Visibility.Visible;
            Get_Database.Margin = new Thickness(460, -22, 450, 0);
            ComboBoxZone1.Visibility = Visibility.Visible;
            LabelReport.Content = "Reports in the order of above text boxes";
            GenerateMetadata.Margin = new Thickness(-700, 40, 897.333, 0);

            //  GenerateMetadata.IsChecked = false;
            Output.IsChecked = false;
            CallGraphButton.Visibility = Visibility.Visible;
            Show_by_Report.Visibility = Visibility.Visible;
            button1.Visibility = Visibility.Visible;
            ReqButton.Visibility = Visibility.Visible;
            SignOutButton.Visibility = Visibility.Visible;



        }

        private async void RemoveServer_Click(object sender, RoutedEventArgs e)
        {
            // ServerStack.Height = 350;
            // Animation.Visibility = Visibility.Visible;
            // ServerStack.Visibility = Visibility.Hidden;
            //StackGrid.Visibility = Visibility.Hidden;


            //dataGrid1.Visibility = Visibility.Collapsed;
            ComboBoxZone.DataContext = null;
            ComboBoxZone.ItemsSource = null;
            ComboBoxZone.Items.Clear();

            ResultText2.Clear();
            ResultText2.Visibility = Visibility.Collapsed;
            Server2Bord.Visibility = Visibility.Collapsed;
            RemoveServer.Visibility = Visibility.Collapsed;
            AddServer.Visibility = Visibility.Visible;
            AddServer2.Visibility = Visibility.Collapsed;
            //MessageBox.Show("Result :" + ResultText2.Text);
            // ComboBoxZone.DataContext = null;
            ComboBoxZone1.Visibility = Visibility.Collapsed;
            ComboBoxZone1.ItemsSource = null;
            ComboBoxZone1.Items.Clear();
            LabelReport.Content = "Reports";
            //dataGrid1.Visibility = Visibility.Collapsed;
            //Animation.Visibility = Visibility.Collapsed;
            //StackGrid.Visibility = Visibility.Visible;
            //  ServerStack.Visibility = Visibility.Visible;
            Get_Database.Margin = new Thickness(240, -22, 450, 0);
            GenerateMetadata.Margin = new Thickness(-700, 20, 897.333, 0);
            //  GenerateMetadata.IsChecked = false;
            Output.IsChecked = false;
            CallGraphButton.Visibility = Visibility.Visible;
            Show_by_Report.Visibility = Visibility.Visible;
            button1.Visibility = Visibility.Visible;
            ReqButton.Visibility = Visibility.Visible;
            SignOutButton.Visibility = Visibility.Visible;

        }

        private void Light_Click(object sender, RoutedEventArgs e)
        {
            Light_Call();
        }

        private void Dark_Click(object sender, RoutedEventArgs e)
        {
            string query = "";
            string queryString = check(query);

            // Dark_Call();
            MessageBox.Show(queryString);


        }
        public void Dark_Call()
        {
            //Dark.Visibility = Visibility.Collapsed;
            Light.Visibility = Visibility.Visible;
            //Application.Current.Resources.Source = new Uri("/Themes/Dark.xaml", UriKind.RelativeOrAbsolute);

            //CallGraphButton.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            SignOutButton.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            //AddServer.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            AddServer2.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            RemoveServer.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            RemoveServer2.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            Get_Database.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            Show_by_Report.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));



            ResultText.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            ResultText2.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            ResultText3.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));


            //  dataGrid1.RowBackground = new SolidColorBrush(Color.FromRgb(60, 60, 61));

            StackBG.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));

            if (button1.IsEnabled == false)
            {
                button1.Background = new SolidColorBrush(Color.FromRgb(204, 204, 204));
                button1.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));

            }
            else
            {
                button1.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));
                button1.Foreground = new SolidColorBrush(Colors.White);
            }
            if (CallGraphButton.IsEnabled == false)
            {
                CallGraphButton.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
                CallGraphButton.Background = new SolidColorBrush(Color.FromRgb(204, 204, 204));
            }
            else
            {
                CallGraphButton.Background = new SolidColorBrush(Color.FromRgb(60, 60, 61));
                CallGraphButton.Foreground = new SolidColorBrush(Colors.White);
            }


            LabelServer.Foreground = new SolidColorBrush(Colors.White);

            LabelReport.Foreground = new SolidColorBrush(Colors.White);

            SignOutButton.Foreground = new SolidColorBrush(Colors.White);
            // AddServer.Foreground = new SolidColorBrush(Colors.White);
            AddServer2.Foreground = new SolidColorBrush(Colors.White);
            RemoveServer.Foreground = new SolidColorBrush(Colors.White);
            RemoveServer2.Foreground = new SolidColorBrush(Colors.White);
            Get_Database.Foreground = new SolidColorBrush(Colors.White);
            Show_by_Report.Foreground = new SolidColorBrush(Colors.White);
            // button1.Foreground = new SolidColorBrush(Colors.White);

            ResultText.Foreground = new SolidColorBrush(Colors.White);
            ResultText2.Foreground = new SolidColorBrush(Colors.White);
            ResultText3.Foreground = new SolidColorBrush(Colors.White);


            // dataGrid1.Foreground = new SolidColorBrush(Colors.White);

        }

        public void Light_Call()
        {
            Light.Visibility = Visibility.Collapsed;
            // Dark.Visibility = Visibility.Visible;
            //Application.Current.Resources.Source = new Uri("/Themes/Dark.xaml", UriKind.RelativeOrAbsolute);

            // CallGraphButton.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            SignOutButton.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            //AddServer.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            AddServer2.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            RemoveServer.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            RemoveServer2.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            Get_Database.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            Show_by_Report.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            //button1.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));

            ResultText.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            ResultText2.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
            ResultText3.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));


            // dataGrid1.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));


            LabelServer.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));

            LabelReport.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));

            // StackBG.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));

            if (button1.IsEnabled == false)
            {
                button1.Background = new SolidColorBrush(Colors.Gray);
                button1.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));

            }
            else
            {
                button1.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
                button1.Background = new SolidColorBrush(Colors.White);
            }
            if (CallGraphButton.IsEnabled == false)
            {
                CallGraphButton.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
                CallGraphButton.Background = new SolidColorBrush(Colors.Gray);
            }
            else
            {
                CallGraphButton.Foreground = new SolidColorBrush(Color.FromRgb(60, 60, 61));
                CallGraphButton.Background = new SolidColorBrush(Colors.White);
            }

            // CallGraphButton.Background = new SolidColorBrush(Colors.White);
            SignOutButton.Background = new SolidColorBrush(Colors.White);
            // AddServer.Background = new SolidColorBrush(Colors.White);
            AddServer2.Background = new SolidColorBrush(Colors.White);
            RemoveServer.Background = new SolidColorBrush(Colors.White);
            RemoveServer2.Background = new SolidColorBrush(Colors.White);
            Get_Database.Background = new SolidColorBrush(Colors.White);
            Show_by_Report.Background = new SolidColorBrush(Colors.White);
            // button1.Background = new SolidColorBrush(Colors.White);

            ResultText.Background = new SolidColorBrush(Colors.White);
            ResultText2.Background = new SolidColorBrush(Colors.White);
            ResultText3.Background = new SolidColorBrush(Colors.White);


            // dataGrid1.RowBackground = new SolidColorBrush(Colors.White);
            StackBG.Background = new SolidColorBrush(Colors.White);

        }

        private void AddServer2_Click(object sender, RoutedEventArgs e)
        {
            /*
            //ServerStack.Height = 477;
            ResultText3.Visibility = Visibility.Visible;
            Server3Bord.Visibility = Visibility.Visible;

            RemoveServer2.Visibility = Visibility.Visible;

            RemoveServer.Visibility = Visibility.Collapsed;

            AddServer.Visibility = Visibility.Collapsed;

            AddServer2.Visibility = Visibility.Collapsed;

            ComboBoxServer3.Visibility = Visibility.Visible;

            Show_by_Report.Visibility = Visibility.Collapsed;
            WrapCheck.Visibility = Visibility.Collapsed;
            BorderBox.Visibility = Visibility.Collapsed;
            Get_Database.Visibility = Visibility.Visible;

            CallGraphButton.IsEnabled = false;
            CallGraphButton.Visibility = Visibility.Hidden;

            LabelReport.Content = "Reports in the order of above text boxes";

            Get_Database.Margin = new Thickness(693, -22, 450, 0);
            Get_Database.Width = 135;
            */

        }

        private void RemoveServer2_Click(object sender, RoutedEventArgs e)
        {
            //Animation.Visibility = Visibility.Visible;
            /*

            // dataGrid1.Visibility = Visibility.Collapsed;
            StackGrid.Visibility = Visibility.Collapsed;
            // ScrollViewer.Visibility = Visibility.Collapsed;


            ResultText3.Clear();
            ResultText3.Visibility = Visibility.Collapsed;
            Server3Bord.Visibility = Visibility.Collapsed;


            RemoveServer.Visibility = Visibility.Visible;

            RemoveServer2.Visibility = Visibility.Collapsed;

            AddServer.Visibility = Visibility.Collapsed;

            AddServer2.Visibility = Visibility.Visible;


            //MessageBox.Show("Result :" + ResultText2.Text);
            // ComboBoxZone.DataContext = null;
            ComboBoxServer3.Visibility = Visibility.Collapsed;
            ComboBoxServer3.ItemsSource = null;
            ComboBoxServer3.Items.Clear();
            LabelReport.Content = "Reports in the order of above text boxes";

            Show_by_Report.Visibility = Visibility.Collapsed;
            WrapCheck.Visibility = Visibility.Collapsed;
            BorderBox.Visibility = Visibility.Collapsed;
            Get_Database.Visibility = Visibility.Visible;
            button1.IsEnabled = false;
            button1.Visibility = Visibility.Collapsed;
            ReqButton.Visibility = Visibility.Collapsed;
            Get_Database.Margin = new Thickness(460, -22, 450, 0);
            //Animation.Visibility = Visibility.Collapsed;
            */
        }
        public DataTable LINQResultToDataTable<T>(IEnumerable<T> Linqlist)
        {
            DataTable dt = new DataTable();
            PropertyInfo[] columns = null;
            if (Linqlist == null) return dt;
            foreach (T Record in Linqlist)
            {
                if (columns == null)
                {
                    columns = ((Type)Record.GetType()).GetProperties();
                    foreach (PropertyInfo GetProperty in columns)
                    {
                        Type IcolType = GetProperty.PropertyType;
                        if ((IcolType.IsGenericType) && (IcolType.GetGenericTypeDefinition()
                        == typeof(Nullable<>)))
                        {
                            IcolType = IcolType.GetGenericArguments()[0];
                        }
                        dt.Columns.Add(new DataColumn(GetProperty.Name, IcolType));
                    }
                }
                DataRow dr = dt.NewRow();
                foreach (PropertyInfo p in columns)
                {
                    dr[p.Name] = p.GetValue(Record, null) == null ? DBNull.Value : p.GetValue(Record, null);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        private void ResultText2_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void ResultText3_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void CheckBoxZone_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            // ShowInTaskbar = true;
            this.Close();
            //this.Hide();
            Application.Current.Shutdown();
        }

        private void Maximize_Click(object sender, RoutedEventArgs e)
        {
            if (this.WindowState == WindowState.Maximized)
            {
                this.WindowState = WindowState.Normal;
            }
            else
            {
                this.WindowState = WindowState.Maximized;

            }
        }

        private void Minimize_Click(object sender, RoutedEventArgs e)
        {
            this.ShowInTaskbar = false;
            MyNotifyIcon.BalloonTipTitle = "Minimize Sucessful";
            MyNotifyIcon.BalloonTipText = "Minimized the app ";
            MyNotifyIcon.ShowBalloonTip(400);
            MyNotifyIcon.Visible = true;
            this.WindowState = WindowState.Minimized;
            //PopText.Visibility = Visibility.Collapsed;
            //popup1.Visibility = Visibility.Collapsed;
            //ShowInTaskbar = true;
        }
        void MyNotifyIcon_MouseDoubleClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            this.WindowState = WindowState.Normal;
        }

        private void CalcTablesCheck_Checked(object sender, RoutedEventArgs e)
        {
            Workspace.IsChecked = true;
            DatasetCheck.IsChecked = true;
            ReportsCheck.IsChecked = true;
            ColumnsCheck.IsChecked = true;
        }

        private void MeasuresCheck_Checked(object sender, RoutedEventArgs e)
        {
            Workspace.IsChecked = true;
            DatasetCheck.IsChecked = true;
            ReportsCheck.IsChecked = true;
            ColumnsCheck.IsChecked = true;
        }

        private void CalcColumnsCheck_Checked(object sender, RoutedEventArgs e)
        {
            Workspace.IsChecked = true;
            DatasetCheck.IsChecked = true;
            ReportsCheck.IsChecked = true;
            ColumnsCheck.IsChecked = true;
        }
        static int findNthOccur(String str,
                    char ch, int N)
        {
            int occur = 0;

            // Loop to find the Nth
            // occurrence of the character
            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] == ch)
                {
                    occur += 1;
                }
                if (occur == N)
                    return i;
            }
            return -1;
        }

        private void Source_Checked(object sender, RoutedEventArgs e)
        {
            Workspace.IsChecked = true;
            DatasetCheck.IsChecked = true;
            ReportsCheck.IsChecked = true;
            ColumnsCheck.IsChecked = true;
        }

        private void Relationships_Checked(object sender, RoutedEventArgs e)
        {
            Workspace.IsChecked = true;
            DatasetCheck.IsChecked = true;
            ReportsCheck.IsChecked = true;
            ColumnsCheck.IsChecked = true;
        }
        public void createsqlDatabase()
        {
            string connectionString = @"Data Source = " + serverlabel.ToString().Replace("\\\\", "\\") + "; Integrated Security=true";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string strconnection = "Data Source = " + serverlabel.ToString().ToString() + "; Integrated Security=true";

            string table = "IF NOT EXISTS(SELECT name FROM master.dbo.sysdatabases WHERE Name='Power BI Metadata') CREATE DATABASE[Power BI Metadata]";
            InsertQuery(table, strconnection);
        }

        public void createsqltable(DataTable dt, string tablename)
        {
            createsqlDatabase();
            string connectionString = @"Data Source = " + serverlabel.ToString().Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=Power BI Metadata";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string strconnection = "Data Source = " + serverlabel.ToString().ToString() + "; Integrated Security=true; Initial Catalog=Power BI Metadata";
            // MessageBox.Show(strconnection.ToString());
            /* string table = "";
             table += "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" + tablename + "]') AND type in (N'U'))";
             table += "BEGIN ";
             table +=" DROP TABLE " + tablename + " ";
             table += " create table " + tablename + "";
             table += "(";
             for (int i = 0; i < dt.Columns.Count; i++)
             {
                 if (i != dt.Columns.Count - 1)
                     table += "[" + dt.Columns[i].ColumnName+ "]" + " " + "varchar(max)" + ",";
                 else
                     table += "[" + dt.Columns[i].ColumnName + "]" + " " + "varchar(max)";
             }
             table += ") ";
             table += "END";*/

            string table = "\n IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='TMSchema_Relationships') BEGIN ";
            table += "\n CREATE TABLE [dbo].[TMSchema_Relationships](";
            table += "\n 	[Workspace] [varchar](max) NULL,";
            table += "\n 	[Report Name] [varchar](max) NULL,";
            table += "\n 	[Relationship ID] [varchar](max) NULL,";
            table += "\n 	[FromTableID] [varchar](max) NULL,";
            table += "\n 	[FromColumnID] [varchar](max) NULL,";
            table += "\n 	[FromCardinality] [varchar](max) NULL,";
            table += "\n 	[ToTableID] [varchar](max) NULL,";
            table += "\n 	[ToColumnID] [varchar](max) NULL,";
            table += "\n 	[ToCardinality] [varchar](max) NULL,";
            table += "\n 	[IsActive] [varchar](max) NULL,";
            table += "\n 	[CrossFilteringBehavior] [varchar](max) NULL";
            table += "\n ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] END";
            table += "\n IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='TMSchema_Columns') BEGIN";
            table += "\n CREATE TABLE [dbo].[TMSchema_Columns](";
            table += "\n 	[Workspace] [varchar](max) NULL,";
            table += "\n 	[Report Name] [varchar](max) NULL,";
            table += "\n 	[Table ID] [varchar](max) NULL,";
            table += "\n 	[Column ID] [varchar](max) NULL,";
            table += "\n 	[Column Name] [varchar](max) NULL";
            table += "\n ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] END";
            table += "\n IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='TMSchema_Table') BEGIN ";
            table += "\n CREATE TABLE [dbo].[TMSchema_Table](";
            table += "\n 	[Workspace] [varchar](max) NULL,";
            table += "\n 	[Report Name] [varchar](max) NULL,";
            table += "\n 	[Table ID] [varchar](max) NULL,";
            table += "\n 	[Table Name] [varchar](max) NULL";
            table += "\n ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] END";
            table += "\n IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Last_Update') BEGIN ";
            table += "\n CREATE TABLE [dbo].[Last_Update](";
            table += "\n 	[Workspace] [varchar](max) NULL,";
            table += "\n 	[Report Name] [varchar](max) NULL,";
            table += "\n 	[Table Name] [varchar](max) NULL,";
            table += "\n 	[RefreshedTime] [varchar](max) NULL";
            table += "\n ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] END";
            table += "\n IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Relationships_Size') BEGIN";
            table += "\n CREATE TABLE [dbo].[Relationships_Size](";
            table += "\n 	[Workspace] [varchar](max) NULL,";
            table += "\n 	[Report Name] [varchar](max) NULL,";
            table += "\n 	[TABLE_NAME] [varchar](max) NULL,";
            table += "\n 	[RELATIONSHIP_ID] [varchar](max) NULL,";
            table += "\n 	[USED_SIZE] [varchar](max) NULL";
            table += "\n ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] END";
            table += "\n IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Data_Size') BEGIN";
            table += "\n CREATE TABLE [dbo].[Data_Size](";
            table += "\n 	[Workspace] [varchar](max) NULL,";
            table += "\n 	[Report Name] [varchar](max) NULL,";
            table += "\n 	[TABLE_NAME] [varchar](max) NULL,";
            table += "\n 	[PARTITION_NAME] [varchar](max) NULL,";
            table += "\n 	[COLUMN_NAME] [varchar](max) NULL,";
            table += "\n 	[SEGMENT_NUMBER] [varchar](max) NULL,";
            table += "\n 	[TABLE_PARTITION_NUMBER] [varchar](max) NULL,";
            table += "\n 	[SEGMENT_ROWS] [varchar](max) NULL,";
            table += "\n 	[USED_SIZE] [varchar](max) NULL,";
            table += "\n 	[COMPRESSION_TYPE] [varchar](max) NULL,";
            table += "\n 	[BITS_COUNT] [varchar](max) NULL,";
            table += "\n 	[BOOKMARK_BITS_COUNT] [varchar](max) NULL,";
            table += "\n 	[VERTIPAQ_STATE] [varchar](max) NULL";
            table += "\n ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] END";
            table += "\n IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Hierarchy') BEGIN";
            table += "\n CREATE TABLE [dbo].[Hierarchy](";
            table += "\n 	[Workspace] [varchar](max) NULL,";
            table += "\n 	[Report Name] [varchar](max) NULL,";
            table += "\n 	[TABLE_NAME] [varchar](max) NULL,";
            table += "\n 	[STRUCTURE_NAME] [varchar](max) NULL,";
            table += "\n 	[SEGMENT_NUMBER] [varchar](max) NULL,";
            table += "\n 	[TABLE_PARTITION_NUMBER] [varchar](max) NULL,";
            table += "\n 	[USED_SIZE] [varchar](max) NULL,";
            table += "\n 	[COLUMN_HIERARCHY_ID] [varchar](max) NULL";
            table += "\n ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] END";
            table += "\n IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='User_Hierarchy') BEGIN ";
            table += "\n CREATE TABLE [dbo].[User_Hierarchy](";
            table += "\n 	[Workspace] [varchar](max) NULL,";
            table += "\n 	[Report Name] [varchar](max) NULL,";
            table += "\n 	[TABLE_NAME] [varchar](max) NULL,";
            table += "\n 	[STRUCTURE_NAME] [varchar](max) NULL,";
            table += "\n 	[USED_SIZE] [varchar](max) NULL,";
            table += "\n 	[HIERARCHY_ID] [varchar](max) NULL";
            table += "\n ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] END";
            table += "\n IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Dictionary_Usage') BEGIN";
            table += "\n CREATE TABLE [dbo].[Dictionary_Usage](";
            table += "\n 	[Workspace] [varchar](max) NULL,";
            table += "\n 	[Report Name] [varchar](max) NULL,";
            table += "\n 	[TABLE_NAME] [varchar](max) NULL,";
            table += "\n 	[COLUMN_ID] [varchar](max) NULL,";
            table += "\n 	[COLUMN_NAME] [varchar](max) NULL,";
            table += "\n 	[Data Type] [varchar](max) NULL,";
            table += "\n 	[DICTIONARY_SIZE_BYTES] [varchar](max) NULL,";
            table += "\n 	[COLUMN_ENCODING_INT] [varchar](max) NULL";
            table += "\n ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] END";
            table += "\n IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Metadata') BEGIN";
            table += "\n CREATE TABLE [dbo].[Metadata](";
            table += "\n 	[Workspace] [varchar](max) NULL,";
            table += "\n 	[Report Name] [varchar](max) NULL,";
            table += "\n 	[Dataset Name] [varchar](max) NULL,";
            table += "\n 	[Column Name] [varchar](max) NULL,";
            table += "\n 	[Calculated Column Expression] [varchar](max) NULL,";
            table += "\n 	[Calculated Measure Expression] [varchar](max) NULL,";
            table += "\n 	[Calculated Table Expression] [varchar](max) NULL,";
            table += "\n 	[Source] [varchar](max) NULL,";
            table += "\n 	[Database Or Path] [varchar](max) NULL,";
            table += "\n 	[Query] [varchar](max) NULL,";
            table += "\n 	[DatabaseItem] [varchar](max) NULL,";
            table += "\n 	[Steps] [varchar](max) NULL,";
            table += "\n 	[From Table Name] [varchar](max) NULL,";
            table += "\n 	[From Column Name] [varchar](max) NULL,";
            table += "\n 	[To Table Name] [varchar](max) NULL,";
            table += "\n 	[To Column Name] [varchar](max) NULL,";
            table += "\n 	[Refreshed Time] [varchar](max) NULL";
            table += "\n ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] END";

            //string table = "TRUNCATE TABLE [Power BI Metadata].[dbo].[Metadata]";
            InsertQuery(table, strconnection.ToString());
            CopyData(strconnection, dt, tablename);
        }
        public void createsqltableUsage(DataTable dt, string tablename)
        {
            string connectionString = @"Data Source = " + serverlabel.ToString().Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=Power BI Metadata";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string strconnection = "Data Source = " + serverlabel.ToString().ToString() + "; Integrated Security=true; Initial Catalog=Power BI Metadata";


            string table = "";
            /* string table = "";
             table += "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" + tablename + "]') AND type in (N'U'))";
             table += "BEGIN ";
             table +=" DROP TABLE " + tablename + " ";
             table += " create table " + tablename + "";
             table += "(";
             for (int i = 0; i < dt.Columns.Count; i++)
             {
                 if (i != dt.Columns.Count - 1)
                     table += "[" + dt.Columns[i].ColumnName+ "]" + " " + "varchar(max)" + ",";
                 else
                     table += "[" + dt.Columns[i].ColumnName + "]" + " " + "varchar(max)";
             }
             table += ") ";
             table += "END";*/

       /*   if (tablename.Equals("Dictionary_Usage"))
            {
                table = "TRUNCATE TABLE [Power BI Metadata].[dbo].[Dictionary_Usage]";
            }
            if (tablename.Equals("User_Hierarchy"))
            {
                table = "TRUNCATE TABLE [Power BI Metadata].[dbo].[User_Hierarchy]";
            }
            if (tablename.Equals("Hierarchy"))
            {
                table = "TRUNCATE TABLE [Power BI Metadata].[dbo].[Hierarchy]";
            }
            if (tablename.Equals("Data_Size"))
            {
                table = "TRUNCATE TABLE [Power BI Metadata].[dbo].[Data_Size]";
            }

            if (tablename.Equals("Relationships_Size"))
            {
                table = "TRUNCATE TABLE [Power BI Metadata].[dbo].[Relationships_Size]";
            }
            if (tablename.Equals("Last_Update"))
            {
                table = "TRUNCATE TABLE [Power BI Metadata].[dbo].[Last_Update]";
            }
            if (tablename.Equals("TMSchema_Table"))
            {
                table = "TRUNCATE TABLE [Power BI Metadata].[dbo].[TMSchema_Table]";
            }
            if (tablename.Equals("TMSchema_Columns"))
            {
                table = "TRUNCATE TABLE [Power BI Metadata].[dbo].[TMSchema_Columns]";
            }
            if (tablename.Equals("TMSchema_Relationships"))
            {
                table = "TRUNCATE TABLE [Power BI Metadata].[dbo].[TMSchema_Relationships]";
            } */

           // InsertQuery(table, strconnection);
            CopyData(strconnection, dt, tablename);
        }
        public void InsertQuery(string qry, string connection)
        {


            SqlConnection _connection = new SqlConnection(connection);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = qry;
            cmd.Connection = _connection;
            _connection.Open();
            cmd.ExecuteNonQuery();
            _connection.Close();
        }
        public void CopyData(string connStr, DataTable dt, string tablename)
        {
            try
            {
                using (SqlBulkCopy bulkCopy =
                new SqlBulkCopy(connStr, SqlBulkCopyOptions.TableLock))
                {
                    bulkCopy.DestinationTableName = tablename;
                    bulkCopy.WriteToServer(dt);
                }
                MessageBox.Show("Data loaded to the " + Server.Text.ToString());
            }
            catch (Exception ex)
            {
               
            }
           
        }

        private void ReqButton_Click(object sender, RoutedEventArgs e)
        {
            int result = 0;

            string connectionstring = "Data Source=" + serverlabel.ToString() + "; Integrated Security=true; Initial Catalog=Power BI Metadata"; ; //your connectionstring    


            if (serverlabel.ToString().Equals(""))
            {
                MessageBox.Show("Click Load Data Options to populate the data base-> Then click the document generator");
            }
            else
            {

                using (SqlConnection conn = new SqlConnection(connectionstring))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("select COUNT(*) from dbo.Metadata", conn);
                    result = (int)cmd.ExecuteScalar();
                    conn.Close();
                }
                if (serverlabel.ToString().ToString().Equals("") || result == 0)
                {
                    MessageBox.Show("Either the Metadata is not extracted or the SQL Server details is blank");
                }
                else
                {
                    Document_Generator objWelcome = new Document_Generator();
                    objWelcome.SQLTB.Text = serverlabel.ToString();
                    objWelcome.Show(); //Sending value from one form to another form.
                    Close();
                }
            }
        }

        private void GenerateMetadata_Checked(object sender, RoutedEventArgs e)
        {
            //button1.Visibility = Visibility.Collapsed;
            //ReqButton.Visibility = Visibility.Collapsed;
            // Show_by_Report.Visibility = Visibility.Visible;
            // CallGraphButton.Visibility = Visibility.Visible;
        }

        private void Output_Checked(object sender, RoutedEventArgs e)
        {
            //   Show_by_Report.Visibility = Visibility.Collapsed;
            // CallGraphButton.Visibility = Visibility.Collapsed;
            //  button1.Visibility = Visibility.Visible;
            //  ReqButton.Visibility = Visibility.Visible;
        }

        private void ComboBoxZone_DropDownClosed(object sender, EventArgs e)
        {

        }

        private void ComboBoxZone1_DropDownClosed(object sender, EventArgs e)
        {

        }
        private void Logout_Click(object sender, RoutedEventArgs e)
        {
            MainOptions objWelcome = new MainOptions();
            objWelcome.Show();
            this.Close();
        }

        private void Server_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void TextPython_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }

    public class DDL_Report
    {

        public string Country_Name
        {
            get;
            set;
        }
        public Boolean Check_Status
        {
            get;
            set;
        }
    }


}