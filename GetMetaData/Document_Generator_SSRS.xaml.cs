using IronPython.Hosting;
using Microsoft.Scripting.Hosting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace GetMetaData
{
    /// <summary>
    /// Interaction logic for Document_Generator_SSRS.xaml
    /// </summary>
    public partial class Document_Generator_SSRS : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;
        private static string PythonPath1;
        private static string TemplatePathString;
        private static string DestinationPathString;
        public Document_Generator_SSRS()
        {
            //Document_Generator_Load();
            InitializeComponent();
            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            MyNotifyIcon.Icon = new System.Drawing.Icon(
                            @"Final.ico");
            MyNotifyIcon.MouseDoubleClick +=
                new System.Windows.Forms.MouseEventHandler(MyNotifyIcon_MouseDoubleClick);
            TokenInfo_DGMS.Text = "Click on Get Reports to get Started";
            GenerateDocAll_DGMS.Visibility = Visibility.Collapsed;
        }
        public static string RunFromCmd(string rCodeFilePath)
        {
            string file = rCodeFilePath;
            string result = string.Empty;

            try
            {


                var info = new ProcessStartInfo(PythonPath1 + @"\python.exe");
                info.Arguments = rCodeFilePath;

                info.RedirectStandardInput = false;
                info.RedirectStandardOutput = true;
                info.UseShellExecute = false;
                info.CreateNoWindow = true;

                using (var proc = new Process())
                {
                    proc.StartInfo = info;
                    proc.Start();
                    proc.WaitForExit();
                    if (proc.ExitCode == 0)
                    {
                        result = proc.StandardOutput.ReadToEnd();
                    }
                }
                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Either the Python path is invalid or the Tableau Server connection was not done. Please contact the administrator for further info.");
                return null;
            }
        }
        private void BindComboBox()
        {
            LabelServer_DGMS.Visibility = Visibility.Collapsed;
            LabelServer_DGMS1.Visibility = Visibility.Collapsed;
            ComboBoxZone_DGMS.Visibility = Visibility.Collapsed;
            ComboBoxZone_DGMS1.Visibility = Visibility.Collapsed;
            LabelPythonPath_DGMS.Visibility = Visibility.Collapsed;
            BorderPythonPath_DGMS.Visibility = Visibility.Collapsed;
            TemplatePath_DGMS.Visibility = Visibility.Collapsed;
            BorderTemplatePAth_DGMS.Visibility = Visibility.Collapsed;
            DestinationPath_DGMS.Visibility = Visibility.Collapsed;
            DestinationPathText_DGMS.Visibility = Visibility.Collapsed;
            Browse_Copy_DGMS.Visibility = Visibility.Collapsed;
            Template_Browse_DGMS.Visibility = Visibility.Collapsed;
            DestPath_Browse_DGMS.Visibility = Visibility.Collapsed;
            SignOutButton_DGMS.Visibility = Visibility.Collapsed;
            GetReports_DGMS.Visibility = Visibility.Collapsed;
            Info_DGMS.Visibility = Visibility.Collapsed;
            GetReports_DGMS1.Visibility = Visibility.Collapsed;
            GenerateDoc_DGMS.Visibility = Visibility.Collapsed;
            GenerateDocAll_DGMS.Visibility = Visibility.Collapsed;
            Animation_DGMS.Visibility = Visibility.Visible;

            MessageBox.Show("Report loading in drop down. Average Wait Time less than 1 minute.");

            string connectionString = @"Data Source = " + SQLTB_DGMS.Text.ToString().Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=SSRS Metadata";
            //String connectionString = "Data Source=IN3040866W1\\SQLEXPRESS; Integrated Security=true; Initial Catalog=Power BI COE";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand("SELECT distinct [Name] as [Report] FROM [SSRS Metadata].[dbo].[ReportInventory] where [TypeName] in ('Report','Grid Report','Dossier','Grid and Graph Report','Managed Grid Report','SQL Report') and [Path]='" + ComboBoxZone_DGMS1.Text.ToString() + "'", con);
            con.Open();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            adapter.Fill(ds, "t");
            ComboBoxZone_DGMS.ItemsSource = ds.Tables["t"].DefaultView;
            ComboBoxZone_DGMS.DisplayMemberPath = ds.Tables[0].Columns["Report"].ToString();
            ComboBoxZone_DGMS.SelectedValuePath = ds.Tables[0].Columns["Report"].ToString();
            LabelServer_DGMS.Visibility = Visibility.Visible;
            LabelServer_DGMS1.Visibility = Visibility.Visible;
            ComboBoxZone_DGMS.Visibility = Visibility.Visible;
            ComboBoxZone_DGMS1.Visibility = Visibility.Visible;
            LabelPythonPath_DGMS.Visibility = Visibility.Visible;
            BorderPythonPath_DGMS.Visibility = Visibility.Visible;
            TemplatePath_DGMS.Visibility = Visibility.Visible;
            BorderTemplatePAth_DGMS.Visibility = Visibility.Visible;
            DestinationPath_DGMS.Visibility = Visibility.Visible;
            DestinationPathText_DGMS.Visibility = Visibility.Visible;
            Browse_Copy_DGMS.Visibility = Visibility.Visible;
            Template_Browse_DGMS.Visibility = Visibility.Visible;
            DestPath_Browse_DGMS.Visibility = Visibility.Visible;
            SignOutButton_DGMS.Visibility = Visibility.Visible;
            GetReports_DGMS.Visibility = Visibility.Visible;
            GetReports_DGMS1.Visibility = Visibility.Visible;
            Info_DGMS.Visibility = Visibility.Visible;
            GenerateDoc_DGMS.Visibility = Visibility.Visible;
            GenerateDocAll_DGMS.Visibility = Visibility.Collapsed;
            Animation_DGMS.Visibility = Visibility.Collapsed;
        }

        private void BindComboBoxProject()
        {
            LabelServer_DGMS.Visibility = Visibility.Collapsed;
            LabelServer_DGMS1.Visibility = Visibility.Collapsed;
            ComboBoxZone_DGMS.Visibility = Visibility.Collapsed;
            ComboBoxZone_DGMS1.Visibility = Visibility.Collapsed;
            LabelPythonPath_DGMS.Visibility = Visibility.Collapsed;
            BorderPythonPath_DGMS.Visibility = Visibility.Collapsed;
            TemplatePath_DGMS.Visibility = Visibility.Collapsed;
            BorderTemplatePAth_DGMS.Visibility = Visibility.Collapsed;
            DestinationPath_DGMS.Visibility = Visibility.Collapsed;
            DestinationPathText_DGMS.Visibility = Visibility.Collapsed;
            Browse_Copy_DGMS.Visibility = Visibility.Collapsed;
            Template_Browse_DGMS.Visibility = Visibility.Collapsed;
            DestPath_Browse_DGMS.Visibility = Visibility.Collapsed;
            SignOutButton_DGMS.Visibility = Visibility.Collapsed;
            GetReports_DGMS.Visibility = Visibility.Collapsed;
            GetReports_DGMS1.Visibility = Visibility.Collapsed;
            Info_DGMS.Visibility = Visibility.Collapsed;
            GenerateDoc_DGMS.Visibility = Visibility.Collapsed;
            GenerateDocAll_DGMS.Visibility = Visibility.Collapsed;
            Animation_DGMS.Visibility = Visibility.Visible;

            MessageBox.Show("Project loading in drop down. Average Wait Time less than 1 minute.");

            string connectionString = @"Data Source = " + SQLTB_DGMS.Text.ToString().Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=SSRS Metadata";
            //String connectionString = "Data Source=IN3040866W1\\SQLEXPRESS; Integrated Security=true; Initial Catalog=Power BI COE";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand("SELECT distinct [Path] as [Project] FROM [SSRS Metadata].[dbo].[ReportInventory]", con);
            con.Open();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            adapter.Fill(ds, "t");
            ComboBoxZone_DGMS1.ItemsSource = ds.Tables["t"].DefaultView;
            ComboBoxZone_DGMS1.DisplayMemberPath = ds.Tables[0].Columns["Project"].ToString();
            ComboBoxZone_DGMS1.SelectedValuePath = ds.Tables[0].Columns["Project"].ToString();
            LabelServer_DGMS.Visibility = Visibility.Visible;
            LabelServer_DGMS1.Visibility = Visibility.Visible;
            ComboBoxZone_DGMS1.Visibility = Visibility.Visible;
            ComboBoxZone_DGMS.Visibility = Visibility.Visible;
            LabelPythonPath_DGMS.Visibility = Visibility.Visible;
            BorderPythonPath_DGMS.Visibility = Visibility.Visible;
            TemplatePath_DGMS.Visibility = Visibility.Visible;
            BorderTemplatePAth_DGMS.Visibility = Visibility.Visible;
            DestinationPath_DGMS.Visibility = Visibility.Visible;
            DestinationPathText_DGMS.Visibility = Visibility.Visible;
            Browse_Copy_DGMS.Visibility = Visibility.Visible;
            Template_Browse_DGMS.Visibility = Visibility.Visible;
            DestPath_Browse_DGMS.Visibility = Visibility.Visible;
            SignOutButton_DGMS.Visibility = Visibility.Visible;
            GetReports_DGMS.Visibility = Visibility.Visible;
            GetReports_DGMS1.Visibility = Visibility.Visible;
            Info_DGMS.Visibility = Visibility.Visible;
            GenerateDoc_DGMS.Visibility = Visibility.Visible;
            GenerateDocAll_DGMS.Visibility = Visibility.Collapsed;
            Animation_DGMS.Visibility = Visibility.Collapsed;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            // ShowInTaskbar = true;
            this.Close();
            Application.Current.Shutdown();
            //this.Hide();
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
            //ShowInTaskbar = true;
        }
        void MyNotifyIcon_MouseDoubleClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            this.WindowState = WindowState.Normal;
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

        public void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.ShowDialog();
            PythonPathText_DGMS.Text = dialog.SelectedPath;
            PythonPath1 = PythonPathText_DGMS.Text;
        }


        private void Template_Browse_Click(object sender, RoutedEventArgs e)
        {

            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();

            // Launch OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = openFileDlg.ShowDialog();
            // Get the selected file name and display in a TextBox.
            // Load content of file in a TextBlock
            if (result == true)
            {
                TemplatePathText_DGMS.Text = openFileDlg.FileName;
                //TextBlock1.Text = System.IO.File.ReadAllText(openFileDlg.FileName);
            }
        }

        private void DestPath_Browse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.ShowDialog();
            DestPath_DGMS.Text = dialog.SelectedPath;
            DestinationPathString = DestPath_DGMS.Text;
        }

        private void GenerateDoc_Click(object sender, RoutedEventArgs e)
        {
            try { 
            if (String.IsNullOrEmpty(ComboBoxZone_DGMS1.Text) || PythonPathText_DGMS.Text.Equals("") || TemplatePathText_DGMS.Text.Equals("") || DestPath_DGMS.Text.Equals("") || String.IsNullOrEmpty(ComboBoxZone_DGMS.Text))
            {
                MessageBox.Show("Please enter the mandatory fields and try again");
            }
            else
            {
                string TemplatePathvar = TemplatePathText_DGMS.Text.ToString().Replace("\\", "\\\\");
                string DestinationPathvar = DestPath_DGMS.Text.ToString().Replace("\\", "\\\\");
                string SQLServervar = SQLTB_DGMS.Text.ToString();//.Replace("\\", "\\\\");
                string path = Directory.GetCurrentDirectory() + @"\PythonFile\Document_Generator_SSRS.py";

                LabelServer_DGMS1.Visibility = Visibility.Collapsed;
                SQLTB_DGMS1.Visibility = Visibility.Collapsed;
                ComboBoxZone_DGMS1.Visibility = Visibility.Collapsed;
                BorderPasswordShow_DGMS1.Visibility = Visibility.Collapsed;
                PasswordShow_DGMS1.Visibility = Visibility.Collapsed;
                GetReports_DGMS1.Visibility = Visibility.Collapsed;
                LabelServer_DGMS.Visibility = Visibility.Collapsed;
                ComboBoxZone_DGMS.Visibility = Visibility.Collapsed;
                LabelPythonPath_DGMS.Visibility = Visibility.Collapsed;
                BorderPythonPath_DGMS.Visibility = Visibility.Collapsed;
                TemplatePath_DGMS.Visibility = Visibility.Collapsed;
                BorderTemplatePAth_DGMS.Visibility = Visibility.Collapsed;
                DestinationPath_DGMS.Visibility = Visibility.Collapsed;
                DestinationPathText_DGMS.Visibility = Visibility.Collapsed;
                Browse_Copy_DGMS.Visibility = Visibility.Collapsed;
                Template_Browse_DGMS.Visibility = Visibility.Collapsed;
                DestPath_Browse_DGMS.Visibility = Visibility.Collapsed;
                SignOutButton_DGMS.Visibility = Visibility.Collapsed;
                GetReports_DGMS.Visibility = Visibility.Collapsed;
                Info_DGMS.Visibility = Visibility.Collapsed;
                GenerateDoc_DGMS.Visibility = Visibility.Collapsed;
                GenerateDocAll_DGMS.Visibility = Visibility.Collapsed;
                Animation_DGMS.Visibility = Visibility.Visible;
                MessageBox.Show("Requirement document generation in process please wait for sometime.");

                String script = "import os                                                                                                                                                                   ";
                script += "\nimport random                                                                                                                                                                   ";
                script += "\nfrom docx import Document                                                                                                                                                       ";
                script += "\nfrom docx.shared import  Inches,Cm,Pt                                                                                                                                           ";
                script += "\nimport pandas as pd                                                                                                                                                             ";
                script += "\nimport pyodbc as py                                                                                                                                                             ";
                script += "\nfrom docx.oxml.shared import OxmlElement,qn                                                                                                                                     ";

                script += "\nheader_name =[\"Document History\", \"Index\", \"Introduction\", \"Metadata Summary\", \"Report Summary\", \"Data Field\", \"Query\", \"Parameter and Prompts\", \"Conclusion\", \"Appendix\"]      ";
                script += "\nfor i in header_name:                                                                                                                                                           ";
                script += "\n    header_1 = header_name[0]                                                                                                                                                   ";
                script += "\n    header_2 = header_name[1]                                                                                                                                                   ";
                script += "\n    header_3 = header_name[2]                                                                                                                                                   ";
                script += "\n    header_4 = header_name[3]                                                                                                                                                   ";
                script += "\n    header_5 = header_name[4]                                                                                                                                                   ";
                script += "\n    header_6 = header_name[5]                                                                                                                                                   ";
                script += "\n    header_7 = header_name[6]                                                                                                                                                   ";
                script += "\n    header_8 = header_name[7]                                                                                                                                                   ";
                script += "\n    header_9 = header_name[8]                                                                                                                                                   ";
                script += "\n    header_10 = header_name[9]                                                                                                                                                  ";
                script += "\nfile_path = '" + TemplatePathvar + "'";
                script += "\ndocument = Document(file_path)                                                                                                                                                  ";
                script += "\n## Set up a column widths function                                                                                                                                              ";
                script += "\ndef set_column_width(column, width):                                                                                                                                            ";
                script += "\n    for cell in column.cells:                                                                                                                                                   ";
                script += "\n        cell.width = width                                                                                                                                                      ";
                script += "\ndef set_repeat_table_header(row):                                                                                                                                               ";
                script += "\n    tr = row._tr                                                                                                                                                                ";
                script += "\n    trPr = tr.get_or_add_trPr()                                                                                                                                                 ";
                script += "\n    tblHeader = OxmlElement('w:tblHeader')                                                                                                                                      ";
                script += "\n    tblHeader.set(qn('w:val'), \"true\")                                                                                                                                          ";
                script += "\n    trPr.append(tblHeader)                                                                                                                                                      ";
                script += "\n    return row                                                                                                                                                                  ";


                script += "\nconn_str = (                                                                                                                                                                    ";
                script += "\n           r'Driver={SQL Server};'                                                                                                                                              ";
                script += "\n           r'Server=" + SQLServervar + ";'";
                script += "\n           r'Database=SSRS Metadata;'                                                                                                                                           ";
                script += "\n           r'Trusted_Connection=yes;'                                                                                                                                           ";
                script += "\n                   )                                                                                                                                                            ";
                script += "\ncnxn = py.connect(conn_str)                                                                                                                                                     ";
                script += "\ncursor = cnxn.cursor()                                                                                                                                                          ";
                script += "\nreport_name =" + "'" + ComboBoxZone_DGMS.Text.ToString() + "'";


                script += "\nreport_details_df = pd.read_sql(\"select Name as 'Report Name', Path, Created_User_Name, Modified_User_Name, LocalDataSourceName, DataProvider, ConnectionString from ReportInventory where Name = '{}'\".format(report_name), cnxn) ";
                script += "\nreport_details_df = report_details_df.drop_duplicates()";

                script += "\nflatenned_result = pd.read_sql(\"select * from flattened_ssrs_report where [Report Name] = '{}'\".format(report_name), cnxn)                                                        ";

                script += "\npara_and_prompts = pd.read_sql(\"SELECT [ParameterName], [Prompt] FROM ReportInventory where Name='{}'\".format(report_name), cnxn)                                                 ";

                script += "\ndata_source = []                                                                                                                                                                  ";
                script += "\ndata_set = []                                                                                                                                                                     ";
                script += "\ndata_set_type = []                                                                                                                                                                ";
                script += "\ndatabase_name = []                                                                                                                                                                ";
                script += "\nschema = []                                                                                                                                                                       ";
                script += "\ntable_name = []                                                                                                                                                                   ";
                script += "\ndata_field_name = []                                                                                                                                                              ";
                script += "\nquery = []                                                                                                                                                                        ";
                script += "\nfor _, each_row in flatenned_result.iterrows():                                                                                                                                   ";
                script += "\n    data_source.append(each_row['DataSource'])                                                                                                                                    ";
                script += "\n    data_set.append(each_row['DataSet'])                                                                                                                                          ";
                script += "\n    data_set_type.append(each_row['DataSet Type'])                                                                                                                                ";
                script += "\n    database_name.append(each_row['Database Name'])                                                                                                                               ";
                script += "\n    schema.append(each_row['Schema'])                                                                                                                                             ";
                script += "\n    table_name.append(each_row['Table Name'])                                                                                                                                     ";
                script += "\n    data_field_name.append(each_row['DataField'])                                                                                                                                 ";
                script += "\n    query.append(each_row['Query'])                                                                                                                                               ";


                script += "\nreport_summary = pd.DataFrame([list(set(data_source)), list(set(data_set)), list(set(data_set_type)),                                                                             ";
                script += "\n              list(set(database_name)), list(set(schema)), list(set(table_name))],                                                                                                ";
                script += "\n              index =['DataSource', 'DataSet', 'DataSet Type', 'Database Name', 'Schema', 'Table Name']).transpose()                                                              ";
                script += "\nreport_summary = report_summary.drop_duplicates()                                                                                                                                 ";

                script += "\ndata_field = pd.DataFrame(list(set(data_field_name)), columns =['Data Field'])                                                                                                    ";

                script += "\nquery_df = pd.DataFrame(list(set(query)), columns =['Query'])                                                                                                                     ";

                script += "\ndocument.add_heading(f' {header_1}', 0)                                                                                                                                           ";
                script += "\ndata = ['', '', '', '']                                                                                                                                                           ";
                script += "\ntable = document.add_table(rows = 5, cols = 4)                                                                                                                                    ";
                script += "\nrow = table.rows[0].cells                                                                                                                                                         ";
                script += "\nrow[0].text = 'Date'                                                                                                                                                              ";
                script += "\nrow[1].text = 'Version'                                                                                                                                                           ";
                script += "\nrow[2].text = 'Description'                                                                                                                                                       ";
                script += "\nrow[3].text = 'Used by'                                                                                                                                                           ";
                script += "\nfor i in data:                                                                                                                                                                    ";
                script += "\n    row = table.add_row().cells                                                                                                                                                   ";
                script += "\n    table.style = 'TableGrid'                                                                                                                                                     ";
                script += "\ndocument.add_page_break()                                                                                                                                                         ";

                script += "\ndocument.add_heading(f' {header_2}', 0)                                                                                                                                           ";
                script += "\nparagraph = document.add_paragraph()                                                                                                                                              ";
                script += "\nrun = paragraph.add_run()                                                                                                                                                         ";
                script += "\nfldChar = OxmlElement('w:fldChar')  # creates a new element                                                                                                                       ";
                script += "\nfldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element                                                                                                            ";
                script += "\ninstrText = OxmlElement('w:instrText')                                                                                                                                            ";
                script += "\ninstrText.set(qn('xml:space'), 'preserve')  # sets attribute on element                                                                                                           ";

                script += "\ninstrText.text = 'TOC \\\\\\\\o \"1-3\" \\\\\\\\h \\\\\\\\z \\\\\\\\u'   # change 1-3 depending on heading levels you need";
                script += "\nfldChar2 = OxmlElement('w:fldChar')                                                                                                                                               ";
                script += "\nfldChar2.set(qn('w:fldCharType'), 'separate')                                                                                                                                     ";
                script += "\nfldChar3 = OxmlElement('w:t')                                                                                                                                                     ";
                script += "\nfldChar3.text = 'Right - click to update field.'";
                script += "\nfldChar2.append(fldChar3)                                                                                                                                                         ";
                script += "\nfldChar4 = OxmlElement('w:fldChar')                                                                                                                                               ";
                script += "\nfldChar4.set(qn('w:fldCharType'), 'end')                                                                                                                                          ";
                script += "\nr_element = run._r                                                                                                                                                                ";
                script += "\nr_element.append(fldChar)                                                                                                                                                         ";
                script += "\nr_element.append(instrText)                                                                                                                                                       ";
                script += "\nr_element.append(fldChar2)                                                                                                                                                        ";
                script += "\nr_element.append(fldChar4)                                                                                                                                                        ";
                script += "\np_element = paragraph._p                                                                                                                                                          ";
                script += "\ndocument.add_page_break()                                                                                                                                                         ";

                script += "\ndocument.add_heading(f' {header_3}', 1)                                                                                                                                           ";
                script += "\np = document.add_paragraph('This document gives us an idea on the Metadata of the reports in scope. Using this document the audience would be able to identify the Metadata information along with the calculations and Source Target Mapping which will be handy in migration.') ";
                script += "\ndocument.add_page_break()                                                                                                                                          ";

                script += "\ndocument.add_heading(f' {header_4}', 1)                                                                                                                                            ";
                script += "\np = document.add_paragraph('Below is a summary of Metadata extracted')                                                                                                             ";
                script += "\ntable2 = document.add_table(report_details_df.shape[0] + 1, report_details_df.shape[1])                                                                                            ";
                script += "\ntable2.style = 'TableGrid'                                                                                                                                                         ";
                script += "\ntable2.autofit = False                                                                                                                                                             ";
                script += "\nfor j in range(report_details_df.shape[-1]):                                                                                                                                       ";
                script += "\n    table2.cell(0, j).text = report_details_df.columns[j]                                                                                                                          ";
                script += "\nfor i in range(report_details_df.shape[0]):                                                                                                                                        ";
                script += "\n    for j in range(report_details_df.shape[-1]):                                                                                                                                   ";
                script += "\n        table2.cell(i + 1, j).text = str(report_details_df.values[i, j])                                                                                                           ";
                script += "\nfor cell in table2.columns[0].cells:                                                                                                                                               ";
                script += "\n    cell.width = Inches(1)                                                                                                                                                         ";
                script += "\ndocument.add_page_break()                                                                                                                                                          ";

                script += "\ndocument.add_heading(f' {header_5}', 1)                                                                                                                                            ";
                script += "\np = document.add_paragraph('Below is a summary of the report')                                                                                                                     ";
                script += "\ntable1 = document.add_table(report_summary.shape[0] + 1, report_summary.shape[1])                                                                                                  ";
                script += "\ntable1.style = 'TableGrid'                                                                                                                                                         ";
                script += "\ntable1.autofit = False                                                                                                                                                             ";
                script += "\nfor j in range(report_summary.shape[-1]):                                                                                                                                          ";
                script += "\n    table1.cell(0, j).text = report_summary.columns[j]                                                                                                                             ";
                script += "\nfor i in range(report_summary.shape[0]):                                                                                                                                           ";
                script += "\n    for j in range(report_summary.shape[-1]):                                                                                                                                      ";
                script += "\n        table1.cell(i + 1, j).text = str(report_summary.values[i, j])                                                                                                              ";
                script += "\nfor cell in table1.columns[0].cells:                                                                                                                                               ";
                script += "\n    cell.width = Inches(1)                                                                                                                                                         ";
                script += "\nset_repeat_table_header(table1.rows[0])                                                                                                                                            ";
                script += "\ndocument.add_page_break()                                                                                                                                                          ";

                script += "\ndocument.add_heading(f' {header_6}', 1)                                                                                                                                            ";
                script += "\np = document.add_paragraph('Below is the Data Field Names of the report')                                                                                                          ";
                script += "\ntable4 = document.add_table(data_field.shape[0] + 1, data_field.shape[1])                                                                                                          ";
                script += "\ntable4.style = 'TableGrid'                                                                                                                                                         ";
                script += "\ntable4.autofit = True                                                                                                                                                              ";
                script += "\ntable_cells2 = table4._cells                                                                                                                                                       ";
                script += "\nfor i in range(data_field.shape[0]):                                                                                                                                               ";
                script += "\n    for j in range(data_field.shape[-1]):                                                                                                                                          ";
                script += "\n        table_cells2[j].text = str(data_field.columns[j])                                                                                                                          ";
                script += "\n    for j in range(data_field.shape[1]):                                                                                                                                           ";
                script += "\n        table_cells2[j + i * data_field.shape[1]].text = str(data_field.values[i][j])                                                                                              ";
                script += "\ndocument.add_page_break()                                                                                                                                                          ";

                script += "\ndocument.add_heading(f' {header_7}', 1)                                                                                                                                            ";
                script += "\np = document.add_paragraph('Below is a summary of Query for the report')                                                                                                           ";
                script += "\ntable1 = document.add_table(query_df.shape[0] + 1, query_df.shape[1])                                                                                                              ";
                script += "\ntable1.style = 'TableGrid'                                                                                                                                                         ";
                script += "\ntable1.autofit = False                                                                                                                                                             ";
                script += "\nfor j in range(query_df.shape[-1]):                                                                                                                                                ";
                script += "\n    table1.cell(0, j).text = query_df.columns[j]                                                                                                                                   ";
                script += "\nfor i in range(query_df.shape[0]):                                                                                                                                                 ";
                script += "\n    for j in range(query_df.shape[-1]):                                                                                                                                            ";
                script += "\n        table1.cell(i + 1, j).text = str(query_df.values[i, j])                                                                                                                    ";
                script += "\nfor cell in table1.columns[0].cells:                                                                                                                                               ";
                script += "\n    cell.width = Inches(2)                                                                                                                                                         ";
                script += "\nset_repeat_table_header(table1.rows[0])                                                                                                                                            ";
                script += "\ndocument.add_page_break()                                                                                                                                                          ";

                script += "\ndocument.add_heading(f' {header_8}', 1)                                                                                                                                            ";
                script += "\np = document.add_paragraph('Below is a summary of Parameters and Prompts for the report')                                                                                          ";
                script += "\ntable1 = document.add_table(para_and_prompts.shape[0] + 1, para_and_prompts.shape[1])                                                                                              ";
                script += "\ntable1.style = 'TableGrid'                                                                                                                                                         ";
                script += "\ntable1.autofit = False                                                                                                                                                             ";
                script += "\nfor j in range(para_and_prompts.shape[-1]):                                                                                                                                        ";
                script += "\n    table1.cell(0, j).text = para_and_prompts.columns[j]                                                                                                                           ";
                script += "\nfor i in range(para_and_prompts.shape[0]):                                                                                                                                         ";
                script += "\n    for j in range(para_and_prompts.shape[-1]):                                                                                                                                    ";
                script += "\n        table1.cell(i + 1, j).text = str(para_and_prompts.values[i, j])                                                                                                            ";
                script += "\nfor cell in table1.columns[0].cells:                                                                                                                                               ";
                script += "\n    cell.width = Inches(2)                                                                                                                                                         ";
                script += "\nset_repeat_table_header(table1.rows[0])                                                                                                                                            ";
                script += "\ndocument.add_page_break()                                                                                                                                                          ";

                script += "\ndocument.add_heading(f' {header_9}', 1)                                                                                                                                            ";
                script += "\np = document.add_paragraph('The Metadata summary of the reports in scope are defined in this document. This can be further leveraged for the migration and Rationalization.')      ";
                script += "\ndocument.add_page_break()                                                                                                                                                          ";

                script += "\ndocument.add_heading(f' {header_10}', 1)                                                                                                                                           ";
                script += "\ndocument.save(\"" + DestinationPathvar + "\\\\Requirement Document For " + ComboBoxZone_DGMS.SelectedValue + ".docx \")";

                //System.Threading.Thread.Sleep(120000);

                //  File.SetAttributes(path, FileAttributes.Normal);

                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                using (StreamWriter writer = File.CreateText(path))
                {
                    writer.WriteLine(script);
                }

                //System.Threading.Thread.Sleep(30000);
                run_cmd();
                LabelServer_DGMS1.Visibility = Visibility.Visible;
                SQLTB_DGMS1.Visibility = Visibility.Collapsed;
                ComboBoxZone_DGMS1.Visibility = Visibility.Visible;
                BorderPasswordShow_DGMS1.Visibility = Visibility.Collapsed;
                PasswordShow_DGMS1.Visibility = Visibility.Collapsed;
                GetReports_DGMS1.Visibility = Visibility.Visible;

                LabelServer_DGMS.Visibility = Visibility.Visible;
                ComboBoxZone_DGMS.Visibility = Visibility.Visible;
                LabelPythonPath_DGMS.Visibility = Visibility.Visible;
                BorderPythonPath_DGMS.Visibility = Visibility.Visible;
                TemplatePath_DGMS.Visibility = Visibility.Visible;
                BorderTemplatePAth_DGMS.Visibility = Visibility.Visible;
                DestinationPath_DGMS.Visibility = Visibility.Visible;
                DestinationPathText_DGMS.Visibility = Visibility.Visible;
                Browse_Copy_DGMS.Visibility = Visibility.Visible;
                Template_Browse_DGMS.Visibility = Visibility.Visible;
                DestPath_Browse_DGMS.Visibility = Visibility.Visible;
                SignOutButton_DGMS.Visibility = Visibility.Visible;
                GetReports_DGMS.Visibility = Visibility.Visible;
                Info_DGMS.Visibility = Visibility.Visible;
                GenerateDoc_DGMS.Visibility = Visibility.Visible;
                GenerateDocAll_DGMS.Visibility = Visibility.Collapsed;
                Animation_DGMS.Visibility = Visibility.Collapsed;

                MessageBox.Show("Requirement document generation process completed. please check below path:-" + DestPath_DGMS.Text.ToString());
            }
            }catch (Exception ex)
            {
                MessageBox.Show("Please entenr the valid input Exception:-" + ex.Message);
            }
         }


            private async void run_cmd1()
        {
            //MessageBox.Show("Report Generation in process. Average Wait Time less than 1 minute.");

            try
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
                        // Batch script to activate Anaconda
                        sw.WriteLine(PythonPathText_DGMS.Text + @"\Scripts\activate.bat");
                        // Activate your environment
                        // sw.WriteLine("conda activate py3.9.7");
                        // run your script. You can also pass in arguments
                        sw.WriteLine("python Document_Generator_SSRS_ALL.py");
                    }
                }
                process.WaitForExit();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }

        }


        private async void run_cmd()
        {
            //MessageBox.Show("Report Generation in process. Average Wait Time less than 1 minute.");
            
            try
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
                        // Batch script to activate Anaconda
                        sw.WriteLine(PythonPathText_DGMS.Text + @"\Scripts\activate.bat");
                        // Activate your environment
                        // sw.WriteLine("conda activate py3.9.7");
                        // run your script. You can also pass in arguments
                        sw.WriteLine("python Document_Generator_SSRS.py");
                    }
                }
                process.WaitForExit();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }

        }
        private void GenerateDocAll_Click(object sender, RoutedEventArgs e)
        {
            //string TemplatePathvar = TemplatePathText_DGMS.Text.ToString().Replace("\\", "\\\\");
            string DestinationPathvar = DestPath_DGMS.Text.ToString().Replace("\\", "\\\\");
            string SQLServervar = SQLTB_DGMS.Text.ToString();//.Replace("\\", "\\\\");
            string path = Directory.GetCurrentDirectory() + @"\PythonFile\Document_Generator_SSRS_ALL.py";
            /*if (PythonPathText_DGMS.Text.Equals("") || TemplatePathText_DGMS.Text.Equals("") || DestPath_DGMS.Text.Equals("") || String.IsNullOrEmpty(ComboBoxZone_DGMS.Text))
            {
                //MessageBox.Show("Please enter the mandatory fields and try again");
            }
            else if (SQLTB_DGMS.Text.Equals(""))
            {
                //MessageBox.Show("Data is missing- Please load the data again");
            }
            else
            {*/
            LabelServer_DGMS.Visibility = Visibility.Visible;
            ComboBoxZone_DGMS.Visibility = Visibility.Visible;
            LabelPythonPath_DGMS.Visibility = Visibility.Visible;
            BorderPythonPath_DGMS.Visibility = Visibility.Visible;
            TemplatePath_DGMS.Visibility = Visibility.Visible;
            BorderTemplatePAth_DGMS.Visibility = Visibility.Visible;
            DestinationPath_DGMS.Visibility = Visibility.Visible;
            DestinationPathText_DGMS.Visibility = Visibility.Visible;
            Browse_Copy_DGMS.Visibility = Visibility.Visible;
            Template_Browse_DGMS.Visibility = Visibility.Visible;
            DestPath_Browse_DGMS.Visibility = Visibility.Visible;
            SignOutButton_DGMS.Visibility = Visibility.Visible;
            GetReports_DGMS.Visibility = Visibility.Visible;
            Info_DGMS.Visibility = Visibility.Visible;
            GenerateDoc_DGMS.Visibility = Visibility.Visible;
            GenerateDocAll_DGMS.Visibility = Visibility.Collapsed;
            Animation_DGMS.Visibility = Visibility.Collapsed;
            ComboBoxZone_DGMS.Text = "";

            MessageBox.Show("Requirement Document Generation in process. Average Wait Time less than 5 minutes.");
               string script = "import pandas as pd                                                                                                                                                ";
               script += "\nimport pyodbc                                                                                                                                                      ";
               script += "\n                                                                                                                                                                   ";
               script += "\nconn_str = (                                                                                                                                                       ";
               script += "\n           r'Driver={SQL Server};'                                                                                                                                 ";
               script += "\n           r'Server=" + SQLServervar + ";'";
               script += "\n           r'Database=SSRS Metadata;'                                                                                                                              ";
               script += "\n           r'Trusted_Connection=yes;'                                                                                                                              ";
               script += "\n                   )                                                                                                                                               ";
               script += "\ncnxn = pyodbc.connect(conn_str)                                                                                                                                    ";
               script += "\ncursor = cnxn.cursor()                                                                                                                                             ";
                                                                                                                                                                              
               script += "\ncount_df = pd.read_sql(\"select Name as [Report Name], count(ParameterName) as ParameterName, count(Prompt) as Prompt from ReportInventory group by Name\", cnxn)    ";
               script += "\nfield_count_df = pd.read_sql(\"select [Report Name], count(DataField) as DataField from flattened_ssrs_report group by [Report Name]\", cnxn)                        ";
               script += "\nresult_df = pd.merge(field_count_df, count_df, on = 'Report Name')                                                                                                 ";
               script += "\nresult_df.to_csv(\"" + DestinationPathvar + "\\\\All Reports Document" + ".csv\")";

            //   File.SetAttributes(path, FileAttributes.Normal);

            if (File.Exists(path))
            {
                File.Delete(path);
            }
            // Create the file and use streamWriter to write text to it.
            //If the file existence is not check, this will overwrite said file.
            //Use the using block so the file can close and vairable disposed correctly

            using (StreamWriter writer = File.CreateText(path))
            {
                writer.WriteLine(script);
            }
            System.Threading.Thread.Sleep(5000);
            run_cmd1();
            // string fileName = "Tableau Metadata.pbix";
            //string path1 = System.IO.Path.Combine(Environment.CurrentDirectory, @"Report\", fileName);
            //Process.Start(path1);

            LabelServer_DGMS.Visibility = Visibility.Visible;
            ComboBoxZone_DGMS.Visibility = Visibility.Visible;
            LabelPythonPath_DGMS.Visibility = Visibility.Visible;
            BorderPythonPath_DGMS.Visibility = Visibility.Visible;
            TemplatePath_DGMS.Visibility = Visibility.Visible;
            BorderTemplatePAth_DGMS.Visibility = Visibility.Visible;
            DestinationPath_DGMS.Visibility = Visibility.Visible;
            DestinationPathText_DGMS.Visibility = Visibility.Visible;
            Browse_Copy_DGMS.Visibility = Visibility.Visible;
            Template_Browse_DGMS.Visibility = Visibility.Visible;
            DestPath_Browse_DGMS.Visibility = Visibility.Visible;
            SignOutButton_DGMS.Visibility = Visibility.Visible;
            GetReports_DGMS.Visibility = Visibility.Visible;
            Info_DGMS.Visibility = Visibility.Visible;
            GenerateDoc_DGMS.Visibility = Visibility.Visible;
            GenerateDocAll_DGMS.Visibility = Visibility.Collapsed;
            Animation_DGMS.Visibility = Visibility.Collapsed;
            //}
            MessageBox.Show("Requirement Document Generation Process completed.");
        }

        private void GetReports_Click(object sender, RoutedEventArgs e)
        {
            Animation_DGMS.Visibility = Visibility.Visible;
            BindComboBox();
            Animation_DGMS.Visibility = Visibility.Collapsed;
            TokenInfo_DGMS.Text = "Generate Requirement Document for Selected Report -> To view the document for the selected report in the drop down";
            TokenInfo_DGMS.AppendText(Environment.NewLine);
            TokenInfo_DGMS.AppendText("Generate Requirement Document for All Reports -> To view the document for all reports");
        }

        private void GetProjects_Click(object sender, RoutedEventArgs e)
        {
            Animation_DGMS.Visibility = Visibility.Visible;
            BindComboBoxProject();
            Animation_DGMS.Visibility = Visibility.Collapsed;
            TokenInfo_DGMS.Text = "Generate Requirement Document for Selected Report -> To view the document for the selected report in the drop down";
            TokenInfo_DGMS.AppendText(Environment.NewLine);
            TokenInfo_DGMS.AppendText("Generate Requirement Document for All Reports -> To view the document for all reports");
        }



        private void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            SSRS navigation = new SSRS();
            navigation.ShowDialog();

        }
    }
}
