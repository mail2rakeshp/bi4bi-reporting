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
    /// Interaction logic for Document_Generator_Ms.xaml
    /// </summary>
    public partial class Document_Generator_Ms : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;
        private static string PythonPath1;
        private static string TemplatePathString;
        private static string DestinationPathString;
        public Document_Generator_Ms()
        {
            //Document_Generator_Load();
            InitializeComponent();
            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            MyNotifyIcon.Icon = new System.Drawing.Icon(
                            @"Final.ico");
            MyNotifyIcon.MouseDoubleClick +=
                new System.Windows.Forms.MouseEventHandler(MyNotifyIcon_MouseDoubleClick);
            TokenInfo_DGMS.Text = "Click on Get Reports to get Started";
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

            string connectionString = @"Data Source = " + SQLTB_DGMS.Text.ToString().Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=Microstrategy Metadata";
            //String connectionString = "Data Source=IN3040866W1\\SQLEXPRESS; Integrated Security=true; Initial Catalog=Power BI COE";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand("SELECT distinct [object_name] as [Report] FROM [Microstrategy Metadata].[dbo].[mstr_report_master] where [object_type_desc] in ('Grid Report','Dossier','Grid and Graph Report','Managed Grid Report','SQL Report') and [project_name]='" + ComboBoxZone_DGMS1.Text.ToString() + "'", con);
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
            GenerateDocAll_DGMS.Visibility = Visibility.Visible;
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

            string connectionString = @"Data Source = " + SQLTB_DGMS.Text.ToString().Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=Microstrategy Metadata";
            //String connectionString = "Data Source=IN3040866W1\\SQLEXPRESS; Integrated Security=true; Initial Catalog=Power BI COE";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand("SELECT distinct [project_name] as [Project] FROM [Microstrategy Metadata].[dbo].[mstr_report_master]", con);
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
            GenerateDocAll_DGMS.Visibility = Visibility.Visible;
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
            string TemplatePathvar = TemplatePathText_DGMS.Text.ToString().Replace("\\", "\\\\");
            string DestinationPathvar = DestPath_DGMS.Text.ToString().Replace("\\", "\\\\");
            
            string SQLServervar = SQLTB_DGMS.Text.ToString();//.Replace("\\", "\\\\");
            string path = Directory.GetCurrentDirectory() + @"\PythonFile\Document_Generator_MS.py";
            if (PythonPathText_DGMS.Text.Equals("") || TemplatePathText_DGMS.Text.Equals("") || DestPath_DGMS.Text.Equals("") || String.IsNullOrEmpty(ComboBoxZone_DGMS.Text))
            {
                MessageBox.Show("Please enter the mandatory fields and try again");
            }
            else if (SQLTB_DGMS.Text.Equals(""))
            {
                MessageBox.Show("Data is missing- Please load the data again");
            }
            else
            {
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
                GenerateDocAll_DGMS.Visibility = Visibility.Visible;
                Animation_DGMS.Visibility = Visibility.Collapsed;

                MessageBox.Show("Requirement Document Generation in process. Average Wait Time less than 1 minute.");

                string script = "import os";
                script += "\nimport random";
                script += "\nfrom docx import Document";
                script += "\nfrom docx.shared import  Inches,Cm,Pt";
                script += "\nimport pandas as pd";
                script += "\nimport pyodbc";
                script += "\nfrom docx.oxml.shared import OxmlElement,qn";
                script += "\nheader_name =[\"Document History\", \"Index\", \"Introduction\", \"Report details\", \"Report Columns\", \"Report Filters\", \"Report Prompts\", \"Report Source Tables\", \"Conclusion\", \"Appendix\" ]";
                script += "\nfor i in header_name:";
                script += "\n    header_1=header_name[0]"; ;
                script += "\n    header_2=header_name[1]";
                script += "\n    header_3=header_name[2]";
                script += "\n    header_4=header_name[3]";
                script += "\n    header_5=header_name[4]";
                script += "\n    header_6=header_name[5]";
                script += "\n    header_7=header_name[6]";
                script += "\n    header_8=header_name[7]";
                script += "\n    header_9=header_name[8]";
                script += "\n    header_10=header_name[9]";
                script += "\nfile_path = '" + TemplatePathvar + "'";
                script += "\ndocument = Document(file_path)";
                script += "\n## Set up a column widths function";
                script += "\ndef set_column_width(column, width):";
                script += "\n    for cell in column.cells:";
                script += "\n        cell.width = width";
                script += "\ndef set_repeat_table_header(row):";
                script += "\n    tr = row._tr";
                script += "\n    trPr = tr.get_or_add_trPr()";
                script += "\n    tblHeader = OxmlElement('w:tblHeader')";
                script += "\n    tblHeader.set(qn('w:val'), \"true\")";
                script += "\n    trPr.append(tblHeader)";
                script += "\n    return row";
                script += "\nconn_str = (";
                script += "\n           r'Driver={SQL Server};'";
                script += "\n           r'Server=" + SQLServervar + ";'";
                script += "\n           r'Database=Microstrategy Metadata;'";
                script += "\n           r'Trusted_Connection=yes;'";
                script += "\n                   )   ";
                script += "\ncnxn = pyodbc.connect(conn_str)";
                script += "\ncursor = cnxn.cursor()";

                script += "\nreport_name =" + "'" + ComboBoxZone_DGMS.Text.ToString()+ "'";
                script += "\nmaster_report_df = pd.read_sql(\"select* from mstr_report_master where object_name = '{}'\".format(report_name), cnxn )";
                script += "\nflattened_df = pd.read_sql(\"select* from flattened_mstr_object_component_list where l0_object_name = '{}'\".format(report_name), cnxn )";

                script += "\nmaster_report_df = master_report_df[['object_name', 'object_type_desc', 'mstr_user_name', 'object_location']]";
                script += "\nmaster_report_df.columns.values[:] = ['Report Name', 'Report Type', 'Report Owner', 'Report Location']";
                script += "\ncolumns_l1_df = flattened_df[['l1_object_name', 'l1_object_type_desc']][(flattened_df['l1_object_type_desc'] == 'Attribute') | (flattened_df['l1_object_type_desc'] == 'Metric')]";
                script += "\ncolumns_l1_df = columns_l1_df.drop_duplicates()";
                script += "\ncolumns_l1_df.columns.values[:] = ['Column Name', 'Column Type']";
                script += "\nfilters = []";
                script += "\nfor each_row in flattened_df['all_columns']:";
                script += "\n    for each in eval(each_row):";
                script += "\n        if each[0] == 'Filter':";
                script += "\n            filters.append(each[1])";
                script += "\nfilters_df = pd.DataFrame(filters, columns=['Filters'])";
                script += "\nfilters_df = filters_df.drop_duplicates()";
                script += "\nprompts_l1_df = flattened_df[['l1_object_name', 'l1_object_type_desc']][flattened_df['l1_object_type_desc'] == 'Prompt']";
                script += "\nprompts_l1_df = prompts_l1_df.drop_duplicates()";
                script += "\nprompts_l1_df.columns.values[:] = ['Prompt Name', 'Prompt Type']";
                script += "\nlogical_table = []";
                script += "\nfor each_row in flattened_df['all_columns']:";
                script += "\n    for each in eval(each_row):";
                script += "\n        if each[0] == 'Logical Table':";
                script += "\n            logical_table.append(each[1])";
                script += "\nlogical_table_df = pd.DataFrame(logical_table, columns=['Logical Table'])";
                script += "\nlogical_table_df = logical_table_df.drop_duplicates()";
                script += "\ndocument.add_heading(f' {header_1}', 0)";
                script += "\ndata = ['','', '', '']  ";
                script += "\ntable = document.add_table(rows=5, cols=4)";
                script += "\nrow = table.rows[0].cells  ";
                script += "\nrow[0].text = 'Date'";
                script += "\nrow[1].text = 'Version'";
                script += "\nrow[2].text = 'Description'";
                script += "\nrow[3].text = 'Used by'";
                script += "\nfor i in data: ";
                script += "\n    row = table.add_row().cells";
                script += "\n    table.style = 'TableGrid'   ";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_2}', 0)";
                script += "\nparagraph = document.add_paragraph()";
                script += "\nrun = paragraph.add_run()";
                script += "\nfldChar = OxmlElement('w:fldChar')  # creates a new element";
                script += "\nfldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element";
                script += "\ninstrText = OxmlElement('w:instrText')";
                script += "\ninstrText.set(qn('xml:space'), 'preserve')  # sets attribute on element";
                script += "\ninstrText.text = 'TOC \\\\\\\\o \"1-3\" \\\\\\\\h \\\\\\\\z \\\\\\\\u'   # change 1-3 depending on heading levels you need";
                script += "\nfldChar2 = OxmlElement('w:fldChar')";
                script += "\nfldChar2.set(qn('w:fldCharType'), 'separate')";
                script += "\nfldChar3 = OxmlElement('w:t')";
                script += "\nfldChar3.text = 'Right - click to update field.'";
                script += "\nfldChar2.append(fldChar3)";
                script += "\nfldChar4 = OxmlElement('w:fldChar')";
                script += "\nfldChar4.set(qn('w:fldCharType'), 'end')";
                script += "\nr_element = run._r";
                script += "\nr_element.append(fldChar)";
                script += "\nr_element.append(instrText)";
                script += "\nr_element.append(fldChar2)";
                script += "\nr_element.append(fldChar4)";
                script += "\np_element = paragraph._p";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_3}', 1)";
                script += "\np = document.add_paragraph('This document gives us an idea on the Metadata of the reports in scope. Using this document the audience would be able to identify the Metadata information along with the calculations and Source Target Mapping which will be handy in migration.')";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_4}', 1)";
                script += "\np = document.add_paragraph('Below are the report details')";
                script += "\ntable2 = document.add_table(master_report_df.shape[0]+1, master_report_df.shape[1])";
                script += "\ntable2.style = 'TableGrid'";
                script += "\ntable2.autofit = False";
                script += "\nfor j in range(master_report_df.shape[-1]):";
                script += "\n    table2.cell(0,j).text = master_report_df.columns[j]";
                script += "\nfor i in range(master_report_df.shape[0]):";
                script += "\n    for j in range(master_report_df.shape[-1]):";
                script += "\n        table2.cell(i+1,j).text = str(master_report_df.values[i,j])   ";
                script += "\nfor cell in table2.columns[0].cells:";
                script += "\n    cell.width = Inches(1)   ";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_5}', 1)";
                script += "\np = document.add_paragraph('Below are the report columns')";
                script += "\ntable2 = document.add_table(columns_l1_df.shape[0]+1, columns_l1_df.shape[1])";
                script += "\ntable2.style = 'TableGrid'";
                script += "\ntable2.autofit = False";
                script += "\nfor j in range(columns_l1_df.shape[-1]):";
                script += "\n    table2.cell(0,j).text = columns_l1_df.columns[j]";
                script += "\nfor i in range(columns_l1_df.shape[0]):";
                script += "\n    for j in range(columns_l1_df.shape[-1]):";
                script += "\n        table2.cell(i+1,j).text = str(columns_l1_df.values[i,j])";
                script += "\nfor cell in table2.columns[0].cells:";
                script += "\n    cell.width = Inches(2)   ";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_6}', 1)";
                script += "\np = document.add_paragraph('Below are the report filters')";
                script += "\ntable2 = document.add_table(filters_df.shape[0]+1, filters_df.shape[1])";
                script += "\ntable2.style = 'TableGrid'";
                script += "\ntable2.autofit = False";
                script += "\nfor j in range(filters_df.shape[-1]):";
                script += "\n    table2.cell(0,j).text = filters_df.columns[j]";
                script += "\nfor i in range(filters_df.shape[0]):";
                script += "\n    for j in range(filters_df.shape[-1]):";
                script += "\n        table2.cell(i+1,j).text = str(filters_df.values[i,j])   ";
                script += "\nfor cell in table2.columns[0].cells:";
                script += "\n    cell.width = Inches(6)   ";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_7}', 1)";
                script += "\np = document.add_paragraph('Below are the report prompts')";
                script += "\ntable2 = document.add_table(prompts_l1_df.shape[0]+1, prompts_l1_df.shape[1])";
                script += "\ntable2.style = 'TableGrid'";
                script += "\ntable2.autofit = False";
                script += "\nfor j in range(prompts_l1_df.shape[-1]):";
                script += "\n    table2.cell(0,j).text = prompts_l1_df.columns[j]";
                script += "\nfor i in range(prompts_l1_df.shape[0]):";
                script += "\n    for j in range(prompts_l1_df.shape[-1]):";
                script += "\n        table2.cell(i+1,j).text = str(prompts_l1_df.values[i,j])   ";
                script += "\nfor cell in table2.columns[0].cells:";
                script += "\n    cell.width = Inches(3)  ";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_8}', 1)";
                script += "\np = document.add_paragraph('Below are the report sourace tables')";
                script += "\ntable2 = document.add_table(logical_table_df.shape[0]+1, logical_table_df.shape[1])";
                script += "\ntable2.style = 'TableGrid'";
                script += "\ntable2.autofit = False";
                script += "\nfor j in range(logical_table_df.shape[-1]):";
                script += "\n    table2.cell(0,j).text = logical_table_df.columns[j]";
                script += "\nfor i in range(logical_table_df.shape[0]):";
                script += "\n    for j in range(logical_table_df.shape[-1]):";
                script += "\n        table2.cell(i+1,j).text = str(logical_table_df.values[i,j])   ";
                script += "\nfor cell in table2.columns[0].cells:";
                script += "\n    cell.width = Inches(6)   ";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_9}', 1)";
                script += "\np = document.add_paragraph('The Metadata summary of the reports in scope are defined in this document. This can be further leveraged for the migration and Rationalization.')";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_10}', 1)";
                script += "\ndocument.save(\"" + DestinationPathvar + "\\\\Requirement Document For " + ComboBoxZone_DGMS.SelectedValue + ".docx \")";
                
                //System.Threading.Thread.Sleep(120000);

                //  File.SetAttributes(path, FileAttributes.Normal);

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

                //System.Threading.Thread.Sleep(30000);
                run_cmd();
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
                GenerateDocAll_DGMS.Visibility = Visibility.Visible;
                Animation_DGMS.Visibility = Visibility.Collapsed;
            }
            MessageBox.Show("Requirement Document Generation Process completed.");


        }
        private async void run_cmd()
        {
            //MessageBox.Show("Report Generation in process. Average Wait Time less than 1 minute.");
            /*
                        ProcessStartInfo start = new ProcessStartInfo();
                        start.FileName = @"C:\Users\UT481LN\Anaconda3\python.exe";//python path
                        start.WorkingDirectory = Environment.CurrentDirectory + @"\PythonFile";//python file in the directoty
                        start.Arguments = "Test.py 2020-1-1";
                        start.UseShellExecute = false;
                        start.CreateNoWindow = true;
                        start.RedirectStandardOutput = true;
                        string result = "";
                        using (Process process = Process.Start(start))
                        {
                            result = process.StandardOutput.ReadToEnd();
                        }

                        MessageBox.Show(result.ToString());*/
            /*
            try
            {
                string workingDirectory = Environment.CurrentDirectory + @"\PythonFile";
                string fileName = workingDirectory + @"\Document_Generator.py";

                Process p = new Process();
                p.StartInfo = new ProcessStartInfo(@"C:\Users\UT481LN\Anaconda3\python.exe" , fileName)
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                p.Start();

                string output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();
            }*/
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
                        sw.WriteLine("python Document_Generator_MS.py");
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
            string path = Directory.GetCurrentDirectory() + @"\PythonFile\Document_Generator_MS.py";
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
                GenerateDocAll_DGMS.Visibility = Visibility.Visible;
                Animation_DGMS.Visibility = Visibility.Collapsed;
                ComboBoxZone_DGMS.Text = "";

                MessageBox.Show("Requirement Document Generation in process. Average Wait Time less than 5 minutes.");

                
                string script = "\nimport pandas as pd";
                script += "\nimport pyodbc";
                script += "\nconn_str = (";
                script += "\n           r'Driver={SQL Server};'";
                script += "\n           r'Server=" + SQLServervar + ";'";
                script += "\n           r'Database=Microstrategy Metadata;'";
                script += "\n           r'Trusted_Connection=yes;'";
                script += "\n                   )   ";
                script += "\ncnxn = pyodbc.connect(conn_str)";
                script += "\ncursor = cnxn.cursor()";
                script += "\ndf1 = pd.read_sql(\"select project_name, object_name, object_type_desc, object_location, mstr_user_name from mstr_report_master\", cnxn )";
                script += "\ndf2 = pd.read_sql(\"select[Report A],  max([Report A Attribute]), max([Report A Metric]), max([Report A Logical Table]) from report_match_percentage group by [Report A]\", cnxn )";
                script += "\nnew_df = pd.merge(df1, df2, how='left', left_on=['object_name'], right_on=['Report A'])";
                script += "\nnew_df = new_df.drop('Report A', axis=1)";
                script += "\nnew_df.columns.values[:] = ['project_name', 'report_name', 'report_type_desc', 'report_location', 'mstr_user_name', 'report_attribute', 'report_metric', 'report_logical_table']";
                script += "\nnew_df.to_csv(\"" + DestinationPathvar + "\\\\All Reports Document" + ".csv\")";
                

            //System.Threading.Thread.Sleep(12000);

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
                run_cmd();
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
                GenerateDocAll_DGMS.Visibility = Visibility.Visible;
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
            MicStr navigation = new MicStr();
            navigation.ShowDialog();

        }
    }
}
