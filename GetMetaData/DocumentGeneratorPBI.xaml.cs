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
    /// Interaction logic for Document_Generator.xaml
    /// </summary>
    public partial class Document_Generator : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;

        private static string PythonPath1;
        private static string TemplatePathString;
        private static string DestinationPathString;
        public Document_Generator()
        {
            //Document_Generator_Load();
            InitializeComponent();
            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            MyNotifyIcon.Icon = new System.Drawing.Icon(
                            @"Final.ico");
            MyNotifyIcon.MouseDoubleClick +=
                new System.Windows.Forms.MouseEventHandler(MyNotifyIcon_MouseDoubleClick);
            TokenInfo.Text = "Click on Get Reports to get Started";
            Home.Visibility = Visibility.Visible;
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
            LabelServer.Visibility = Visibility.Collapsed;
            ComboBoxZone.Visibility = Visibility.Collapsed;
            LabelPythonPath.Visibility = Visibility.Collapsed;
            BorderPythonPath.Visibility = Visibility.Collapsed;
            TemplatePath.Visibility = Visibility.Collapsed;
            BorderTemplatePAth.Visibility = Visibility.Collapsed;
            DestinationPath.Visibility = Visibility.Collapsed;
            DestinationPathText.Visibility = Visibility.Collapsed;
            Browse_Copy.Visibility = Visibility.Collapsed;
            Template_Browse.Visibility = Visibility.Collapsed;
            DestPath_Browse.Visibility = Visibility.Collapsed;
            Home.Visibility = Visibility.Collapsed;
            GetReports.Visibility = Visibility.Collapsed;
            Info.Visibility = Visibility.Collapsed;
            GenerateDoc.Visibility = Visibility.Collapsed;
            GenerateDocAll.Visibility = Visibility.Collapsed;
            Animation.Visibility = Visibility.Visible;

            MessageBox.Show("Report loading in drop down. Average Wait Time less than 1 minute.");

            //string connectionString = @"Data Source = " + SQLTB.Text.ToString().Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=Power BI Metadata";
             String connectionString = "Data Source=" + SQLTB.Text.ToString().Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=Power BI Metadata";
            //String connectionString = "Data Source=IN3040866W1\\SQLEXPRESS; Integrated Security=true; Initial Catalog=Power BI Metadata";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand("select DISTINCT [Report Name] from dbo.Metadata", con);
            con.Open();
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            adapter.Fill(ds, "t");
            ComboBoxZone.ItemsSource = ds.Tables["t"].DefaultView;
            ComboBoxZone.DisplayMemberPath = ds.Tables[0].Columns["Report Name"].ToString();
            ComboBoxZone.SelectedValuePath = ds.Tables[0].Columns["Report Name"].ToString();
            LabelServer.Visibility = Visibility.Visible;
            ComboBoxZone.Visibility = Visibility.Visible;
            LabelPythonPath.Visibility = Visibility.Visible;
            BorderPythonPath.Visibility = Visibility.Visible;
            TemplatePath.Visibility = Visibility.Visible;
            BorderTemplatePAth.Visibility = Visibility.Visible;
            DestinationPath.Visibility = Visibility.Visible;
            DestinationPathText.Visibility = Visibility.Visible;
            Browse_Copy.Visibility = Visibility.Visible;
            Template_Browse.Visibility = Visibility.Visible;
            DestPath_Browse.Visibility = Visibility.Visible;
            Home.Visibility = Visibility.Visible;
            GetReports.Visibility = Visibility.Visible;
            Info.Visibility = Visibility.Visible;
            GenerateDoc.Visibility = Visibility.Visible;
            GenerateDocAll.Visibility = Visibility.Visible;
            Animation.Visibility = Visibility.Collapsed;
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

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.ShowDialog();
            PythonPathText.Text = dialog.SelectedPath;
            PythonPath1 = PythonPathText.Text;
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
                TemplatePathText.Text = openFileDlg.FileName;
                //TextBlock1.Text = System.IO.File.ReadAllText(openFileDlg.FileName);
            }
        }

        private void DestPath_Browse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.ShowDialog();
            DestPath.Text = dialog.SelectedPath;
            DestinationPathString = DestPath.Text;
        }

        private void GenerateDoc_Click(object sender, RoutedEventArgs e)
        {
            string TemplatePathvar = TemplatePathText.Text.ToString().Replace("\\", "\\\\");
            string DestinationPathvar = DestPath.Text.ToString().Replace("\\","\\\\");
            //string SQLServervar = "IN3040866W1\\SQLEXPRESS";
            string SQLServervar = SQLTB.Text.ToString();//.Replace("\\", "\\\\");
            string path = Directory.GetCurrentDirectory() + @"\PythonFile\Document_Generator.py";
            if (PythonPathText.Text.Equals("") || TemplatePathText.Text.Equals("") || DestPath.Text.Equals("") || String.IsNullOrEmpty(ComboBoxZone.Text))
            {
                MessageBox.Show("Please enter the mandatory fields and try again");
            }
            else
            {
                LabelServer.Visibility = Visibility.Collapsed;
                ComboBoxZone.Visibility = Visibility.Collapsed;
                LabelPythonPath.Visibility = Visibility.Collapsed;
                BorderPythonPath.Visibility = Visibility.Collapsed;
                TemplatePath.Visibility = Visibility.Collapsed;
                BorderTemplatePAth.Visibility = Visibility.Collapsed;
                DestinationPath.Visibility = Visibility.Collapsed;
                DestinationPathText.Visibility = Visibility.Collapsed;
                Browse_Copy.Visibility = Visibility.Collapsed;
                Template_Browse.Visibility = Visibility.Collapsed;
                DestPath_Browse.Visibility = Visibility.Collapsed;
                Home.Visibility = Visibility.Collapsed;
                GetReports.Visibility = Visibility.Collapsed;
                Info.Visibility = Visibility.Collapsed;
                GenerateDoc.Visibility = Visibility.Collapsed;
                GenerateDocAll.Visibility = Visibility.Collapsed;
                Animation.Visibility = Visibility.Visible;

                MessageBox.Show("Requirement Document Generation in process. Average Wait Time less than 1 minute.");

                string script = "import os";
                script += "\nimport random";
                script += "\nfrom docx import Document";
                script += "\nfrom docx.shared import  Inches,Cm,Pt";
                script += "\nimport pandas as pd";
                script += "\nimport pyodbc as py";
                script += "\nfrom docx.oxml.shared import OxmlElement,qn";
                script += "\nheader_name =[\"Document History\", \"Index\", \"Introduction\", \"Metadata Summary\", \"Data source overview\", \"Source Target Mapping\", \"Calculated Columns\", \"Calculated Measures\", \"Calculated Tables\", \"Columns\" ,\"Conclusion\", \"Appendix\" ]";
                script += "\nfor i in header_name:";
                script += "\n    header_1=header_name[0]";
                script += "\n    header_2=header_name[1]";
                script += "\n    header_3=header_name[2]";
                script += "\n    header_4=header_name[3]";
                script += "\n    header_5=header_name[4]";
                script += "\n    header_6=header_name[5]";
                script += "\n    header_7=header_name[6]";
                script += "\n    header_8=header_name[7]";
                script += "\n    header_9=header_name[8]";
                script += "\n    header_10=header_name[9]";
                script += "\n    header_11=header_name[10]";
                script += "\n    header_12=header_name[11]";
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
                script += "\n    return row      ";
                script += "\nconn_str = (";
                script += "\n           r'Driver={SQL Server};'";
                script += "\n           r'Server=" + SQLServervar + ";'";
                script += "\n           r'Database=Power BI Metadata;'";
                script += "\n           r'Trusted_Connection=yes;'";
                script += "\n                   )   ";
                script += "\ncnxn = py.connect(conn_str)";
                script += "\ncursor = cnxn.cursor()";
                script += "\ndftop = pd.read_sql(\"select distinct Workspace,[Report Name],Source,[Database or Path] from dbo.vw_Metadata WHERE [Report Name]='" + ComboBoxZone.SelectedValue +"' AND [Dataset Name] NOT IN ('Internal Date Table','Internal Date Table Template' )\" , cnxn )";
                script += "\ndf2=pd.read_sql(\"select  [Report Name],  [Number of Sources], [Number of Calculated Columns], [Number of Calculated Measures], [Number of Calculated Tables], [Number of Columns] from vw_Metadata_Calculations WHERE [Report Name]='" + ComboBoxZone.SelectedValue + "'\" , cnxn )";
                script += "\ndf3=pd.read_sql(\"select Workspace , [Report Name],COLUMN_NAME [Column Name],TABLE_NAME [Dataset],[Data Type] from dbo.vw_Metadata_Columns WHERE [Report Name]='" + ComboBoxZone.SelectedValue + "'\" , cnxn )";
                script += "\ndf4=pd.read_sql(\"select * FROM vw_Metadata_STM WHERE [Report Name]='" + ComboBoxZone.SelectedValue + "'\" , cnxn )";
                script += "\ndf5=pd.read_sql(\"select distinct Workspace,[Report Name],CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END AS [Dataset Name],[Column Name],[Calculated Column Expression] from dbo.Metadata where [Calculated Column Expression] is not null and [Report Name]='" + ComboBoxZone.SelectedValue + "'\" , cnxn )";
                script += "\ndf6=pd.read_sql(\"select distinct Workspace,[Report Name],CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END AS [Dataset Name],[Column Name],[Calculated Measure Expression] from dbo.Metadata where [Calculated Measure Expression] is not null and [Report Name]='" + ComboBoxZone.SelectedValue + "'\" , cnxn )";
                script += "\ndf7=pd.read_sql(\"select distinct Workspace,[Report Name],CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END AS [Dataset Name],[Calculated Table Expression] from dbo.Metadata where [Calculated Table Expression] is not null and [Report Name]='" + ComboBoxZone.SelectedValue + "'\" , cnxn )";
                script += "\ndocument.save(file_path)";
                script += "\ndocument.add_heading(f' {header_1}', 0)";
                script += "\ndata = ['','', '', '']  ";
                script += "\ntable = document.add_table(rows=5, cols=4)";
                script += "\nrow = table.rows[0].cells      ";
                script += "\nrow[0].text = 'Date'";
                script += "\nrow[1].text = 'Version'";
                script += "\nrow[2].text = 'Description'";
                script += "\nrow[3].text = 'Used by'";
                script += "\nfor i in data: ";
                script += "\n   row = table.add_row().cells";
                script += "\n   table.style = 'TableGrid'   ";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_2}', 0)";
                script += "\nparagraph = document.add_paragraph()";
                script += "\nrun = paragraph.add_run()";
                script += "\nfldChar = OxmlElement('w:fldChar')  # creates a new element";
                script += "\nfldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element";
                script += "\ninstrText = OxmlElement('w:instrText')";
                script += "\ninstrText.set(qn('xml:space'), 'preserve')  # sets attribute on element";
                script += "\ninstrText.text = 'TOC \\\\o \"1-3\" \\\\h \\\\z \\\\u'   # change 1-3 depending on heading levels you need";
                script += "\nfldChar2 = OxmlElement('w:fldChar')";
                script += "\nfldChar2.set(qn('w:fldCharType'), 'separate')";
                script += "\nfldChar3 = OxmlElement('w:t')";
                script += "\nfldChar3.text = \"Right-click to update field.\"";
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
                script += "\np = document.add_paragraph('Below is a summary of Metadata extracted')";
                script += "\ntable2 = document.add_table(df2.shape[0]+1, df2.shape[1])";
                script += "\ntable2.style = 'TableGrid'";
                script += "\ntable2.autofit = False";
                script += "\nfor j in range(df2.shape[-1]):";
                script += "\n    table2.cell(0,j).text = df2.columns[j]";
                script += "\nfor i in range(df2.shape[0]):";
                script += "\n    for j in range(df2.shape[-1]):";
                script += "\n        table2.cell(i+1,j).text = str(df2.values[i,j])   ";
                script += "\nfor cell in table2.columns[0].cells:";
                script += "\n    cell.width = Inches(2)   ";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_5}', 1)";
                script += "\np = document.add_paragraph('Below is a summary of Data sources for the reports')";
                script += "\ntable1 = document.add_table(dftop.shape[0]+1, dftop.shape[1])";
                script += "\ntable1.style = 'TableGrid' ";
                script += "\ntable1.autofit = False";
                script += "\nfor j in range(dftop.shape[-1]):";
                script += "\n    table1.cell(0,j).text = dftop.columns[j]";
                script += "\nfor i in range(dftop.shape[0]):";
                script += "\n    for j in range(dftop.shape[-1]):";
                script += "\n        table1.cell(i+1,j).text = str(dftop.values[i,j]) ";
                script += "\nfor cell in table1.columns[0].cells:";
                script += "\n    cell.width = Inches(2)";
                script += "\nset_repeat_table_header(table1.rows[0])";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_6}', 1)";
                script += "\np = document.add_paragraph('Below is Source Target Mapping for the reports')";
                script += "\ntable4 = document.add_table(df4.shape[0]+1,df4.shape[1])";
                script += "\ntable4.style = 'TableGrid' ";
                script += "\ntable4.autofit = True";
                script += "\ntable_cells2 = table4._cells";
                script += "\nfor i in range(df4.shape[0]):";
                script += "\n    for j in range(df4.shape[-1]):";
                script += "\n        table_cells2[j].text =  str(df4.columns[j])";
                script += "\n    for j in range(df4.shape[1]):";
                script += "\n        table_cells2[j + i * df4.shape[1]].text =  str(df4.values[i][j])";
                script += "\nfor row in table4.rows:";
                script += "\n    for cell in row.cells:";
                script += "\n        paragraphs = cell.paragraphs";
                script += "\n        for paragraph in paragraphs:";
                script += "\n            for run in paragraph.runs:";
                script += "\n                font = run.font";
                script += "\n                font.size= Pt(6)";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_7}', 1)";
                script += "\np = document.add_paragraph('Below is a summary of Calculated Columns and its expressions for the reports')";
                script += "\ntable5 = document.add_table(df5.shape[0]+1,df5.shape[1])";
                script += "\ntable5.style = 'TableGrid' ";
                script += "\ntable5.autofit = False";
                script += "\ntable_cells5 = table5._cells";
                script += "\nfor i in range(df5.shape[0]):";
                script += "\n    for j in range(df5.shape[-1]):";
                script += "\n        table_cells5[j].text =  str(df5.columns[j])";
                script += "\n    for j in range(df5.shape[1]):";
                script += "\n        table_cells5[j + i * df5.shape[1]].text =  str(df5.values[i][j])";
                script += "\nfor cell in table5.columns[0].cells:";
                script += "\n    cell.width = Inches(2)";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_8}', 1)";
                script += "\np = document.add_paragraph('Below is a summary of Calculated Measures and its expressions for the reports')";
                script += "\ntable6 = document.add_table(df6.shape[0]+1,df6.shape[1])";
                script += "\ntable6.style = 'TableGrid' ";
                script += "\ntable6.autofit = False";
                script += "\ntable_cells6 = table6._cells";
                script += "\nfor i in range(df6.shape[0]):";
                script += "\n    for j in range(df6.shape[-1]):";
                script += "\n        table_cells6[j].text =  str(df6.columns[j])";
                script += "\n    for j in range(df6.shape[1]):";
                script += "\n        table_cells6[j + i * df6.shape[1]].text =  str(df6.values[i][j])";
                script += "\nfor cell in table6.columns[0].cells:";
                script += "\n    cell.width = Inches(2)";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_9}', 1)";
                script += "\np = document.add_paragraph('Below is a summary of Calculated Tables and its expressions for the reports')";
                script += "\ntable7 = document.add_table(df7.shape[0]+1,df7.shape[1])";
                script += "\ntable7.style = 'TableGrid' ";
                script += "\ntable7.autofit = False";
                script += "\ntable_cells7 = table7._cells";
                script += "\nfor i in range(df7.shape[0]):";
                script += "\n    for j in range(df7.shape[-1]):";
                script += "\n        table_cells7[j].text =  str(df7.columns[j])";
                script += "\n    for j in range(df7.shape[1]):";
                script += "\n        table_cells7[j + i * df7.shape[1]].text =  str(df7.values[i][j])";
                script += "\nfor cell in table7.columns[0].cells:";
                script += "\n    cell.width = Inches(2)";
                script += "\ndocument.add_page_break()        ";
                script += "\ndocument.add_heading(f' {header_10}', 1)";
                script += "\np = document.add_paragraph('The below is a summary of Columns for the reports')";
                script += "\ntable3 = document.add_table(df3.shape[0]+1,df3.shape[1])";
                script += "\ntable3.style = 'TableGrid' ";
                script += "\ntable3.autofit = False";
                script += "\ntable_cells = table3._cells";
                script += "\nfor i in range(df3.shape[0]):";
                script += "\n    for j in range(df3.shape[-1]):";
                script += "\n        table_cells[j].text =  str(df3.columns[j])";
                script += "\n    for j in range(df3.shape[1]):";
                script += "\n        table_cells[j + i * df3.shape[1]].text =  str(df3.values[i][j])";
                script += "\nfor cell in table3.columns[0].cells:";
                script += "\n    cell.width = Inches(2)";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_11}', 1)";
                script += "\np = document.add_paragraph('The Metadata summary of the reports in scope are defined in this document. This can be further leveraged for the migration and Rationalization.')";
                script += "\ndocument.add_page_break()";
                script += "\ndocument.add_heading(f' {header_12}', 1)";
                script += "\ndocument.save(\"" + DestinationPathvar + "\\\\Requirement Document For "+ComboBoxZone.SelectedValue+".docx \")";
                script += "\nprint('Done')";
                //System.Threading.Thread.Sleep(120000);

                File.SetAttributes(path, FileAttributes.Normal);

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

                LabelServer.Visibility = Visibility.Visible;
                ComboBoxZone.Visibility = Visibility.Visible;
                LabelPythonPath.Visibility = Visibility.Visible;
                BorderPythonPath.Visibility = Visibility.Visible;
                TemplatePath.Visibility = Visibility.Visible;
                BorderTemplatePAth.Visibility = Visibility.Visible;
                DestinationPath.Visibility = Visibility.Visible;
                DestinationPathText.Visibility = Visibility.Visible;
                Browse_Copy.Visibility = Visibility.Visible;
                Template_Browse.Visibility = Visibility.Visible;
                DestPath_Browse.Visibility = Visibility.Visible;
                Home.Visibility = Visibility.Visible;
                GetReports.Visibility = Visibility.Visible;
                Info.Visibility = Visibility.Visible;
                GenerateDoc.Visibility = Visibility.Visible;
                GenerateDocAll.Visibility = Visibility.Visible;
                Animation.Visibility = Visibility.Collapsed;
            }
            
        
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
                        sw.WriteLine(PythonPath1 + @"\Scripts\activate.bat");
                        // Activate your environment
                        // sw.WriteLine("conda activate py3.9.7");
                        // run your script. You can also pass in arguments
                        sw.WriteLine("python Document_Generator.py");
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
            string TemplatePathvar = TemplatePathText.Text.ToString().Replace("\\", "\\\\");
            string DestinationPathvar = DestPath.Text.ToString().Replace("\\", "\\\\");
            //string SQLServervar = "IN3040866W1\\SQLEXPRESS";
            string SQLServervar = SQLTB.Text.ToString();//.Replace("\\", "\\\\");
            string path = Directory.GetCurrentDirectory() + @"\PythonFile\Document_Generator.py";
            if (PythonPathText.Text.Equals("") || TemplatePathText.Text.Equals("") || DestPath.Text.Equals("") || String.IsNullOrEmpty(ComboBoxZone.Text))
            {
                MessageBox.Show("Please enter the mandatory fields and try again");
            }
            else
            {
                LabelServer.Visibility = Visibility.Collapsed;
                ComboBoxZone.Visibility = Visibility.Collapsed;
                LabelPythonPath.Visibility = Visibility.Collapsed;
                BorderPythonPath.Visibility = Visibility.Collapsed;
                TemplatePath.Visibility = Visibility.Collapsed;
                BorderTemplatePAth.Visibility = Visibility.Collapsed;
                DestinationPath.Visibility = Visibility.Collapsed;
                DestinationPathText.Visibility = Visibility.Collapsed;
                Browse_Copy.Visibility = Visibility.Collapsed;
                Template_Browse.Visibility = Visibility.Collapsed;
                DestPath_Browse.Visibility = Visibility.Collapsed;
                Home.Visibility = Visibility.Collapsed;
                GetReports.Visibility = Visibility.Collapsed;
                Info.Visibility = Visibility.Collapsed;
                GenerateDoc.Visibility = Visibility.Collapsed;
                GenerateDocAll.Visibility = Visibility.Collapsed;
                Animation.Visibility = Visibility.Visible;
                ComboBoxZone.Text = "";
                MessageBox.Show("Requirement Document Generation in process. Average Wait Time less than 5 minutes.");

            /* string script = "import os";
             script += "\nimport random";
             script += "\nfrom docx import Document";
             script += "\nfrom docx.shared import  Inches,Cm,Pt";
             script += "\nimport pandas as pd";
             script += "\nimport pyodbc as py";
             script += "\nfrom docx.oxml.shared import OxmlElement,qn";
             script += "\nheader_name =[\"Document History\", \"Index\", \"Introduction\", \"Metadata Summary\", \"Data source overview\", \"Source Target Mapping\", \"Calculated Columns\", \"Calculated Measures\", \"Calculated Tables\", \"Columns\" ,\"Conclusion\", \"Appendix\" ]";
             script += "\nfor i in header_name:";
             script += "\n    header_1=header_name[0]";
             script += "\n    header_2=header_name[1]";
             script += "\n    header_3=header_name[2]";
             script += "\n    header_4=header_name[3]";
             script += "\n    header_5=header_name[4]";
             script += "\n    header_6=header_name[5]";
             script += "\n    header_7=header_name[6]";
             script += "\n    header_8=header_name[7]";
             script += "\n    header_9=header_name[8]";
             script += "\n    header_10=header_name[9]";
             script += "\n    header_11=header_name[10]";
             script += "\n    header_12=header_name[11]";
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
             script += "\n    return row      ";
             script += "\nconn_str = (";
             script += "\n           r'Driver={SQL Server};'";
             script += "\n           r'Server=" + SQLServervar + ";'";
             script += "\n           r'Database=Power BI Metadata;'";
             script += "\n           r'Trusted_Connection=yes;'";
             script += "\n                   )   ";
             script += "\ncnxn = py.connect(conn_str)";
             script += "\ncursor = cnxn.cursor()";
             script += "\ndftop = pd.read_sql(\"select distinct Workspace,[Report Name],Source,[Database or Path] from dbo.vw_Metadata WHERE  [Dataset Name] NOT IN ('Internal Date Table','Internal Date Table Template' )\" , cnxn )";
             script += "\ndf2=pd.read_sql(\"select  [Report Name],  [Number of Sources], [Number of Calculated Columns], [Number of Calculated Measures], [Number of Calculated Tables], [Number of Columns] from vw_Metadata_Calculations \" , cnxn )";
             script += "\ndf3=pd.read_sql(\"select Workspace , [Report Name],COLUMN_NAME [Column Name],TABLE_NAME [Dataset],[Data Type] from dbo.vw_Metadata_Columns \" , cnxn )";
             script += "\ndf4=pd.read_sql(\"select * FROM vw_Metadata_STM \" , cnxn )";
             script += "\ndf5=pd.read_sql(\"select distinct Workspace,[Report Name],CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END AS [Dataset Name],[Column Name],[Calculated Column Expression] from dbo.Metadata where [Calculated Column Expression] is not null \" , cnxn )";
             script += "\ndf6=pd.read_sql(\"select distinct Workspace,[Report Name],CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END AS [Dataset Name],[Column Name],[Calculated Measure Expression] from dbo.Metadata where [Calculated Measure Expression] is not null \" , cnxn )";
             script += "\ndf7=pd.read_sql(\"select distinct Workspace,[Report Name],CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END AS [Dataset Name],[Calculated Table Expression] from dbo.Metadata where [Calculated Table Expression] is not null \" , cnxn )";
             script += "\ndocument.save(file_path)";
             script += "\ndocument.add_heading(f' {header_1}', 0)";
             script += "\ndata = ['','', '', '']  ";
             script += "\ntable = document.add_table(rows=5, cols=4)";
             script += "\nrow = table.rows[0].cells      ";
             script += "\nrow[0].text = 'Date'";
             script += "\nrow[1].text = 'Version'";
             script += "\nrow[2].text = 'Description'";
             script += "\nrow[3].text = 'Used by'";
             script += "\nfor i in data: ";
             script += "\n   row = table.add_row().cells";
             script += "\n   table.style = 'TableGrid'   ";
             script += "\ndocument.add_page_break()";
             script += "\ndocument.add_heading(f' {header_2}', 0)";
             script += "\nparagraph = document.add_paragraph()";
             script += "\nrun = paragraph.add_run()";
             script += "\nfldChar = OxmlElement('w:fldChar')  # creates a new element";
             script += "\nfldChar.set(qn('w:fldCharType'), 'begin')  # sets attribute on element";
             script += "\ninstrText = OxmlElement('w:instrText')";
             script += "\ninstrText.set(qn('xml:space'), 'preserve')  # sets attribute on element";
             script += "\ninstrText.text = 'TOC \\\\o \"1-3\" \\\\h \\\\z \\\\u'   # change 1-3 depending on heading levels you need";
             script += "\nfldChar2 = OxmlElement('w:fldChar')";
             script += "\nfldChar2.set(qn('w:fldCharType'), 'separate')";
             script += "\nfldChar3 = OxmlElement('w:t')";
             script += "\nfldChar3.text = \"Right-click to update field.\"";
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
             script += "\np = document.add_paragraph('Below is a summary of Metadata extracted')";
             script += "\ntable2 = document.add_table(df2.shape[0]+1, df2.shape[1])";
             script += "\ntable2.style = 'TableGrid'";
             script += "\ntable2.autofit = False";
             script += "\nfor j in range(df2.shape[-1]):";
             script += "\n    table2.cell(0,j).text = df2.columns[j]";
             script += "\nfor i in range(df2.shape[0]):";
             script += "\n    for j in range(df2.shape[-1]):";
             script += "\n        table2.cell(i+1,j).text = str(df2.values[i,j])   ";
             script += "\nfor cell in table2.columns[0].cells:";
             script += "\n    cell.width = Inches(2)   ";
             script += "\ndocument.add_page_break()";
             script += "\ndocument.add_heading(f' {header_5}', 1)";
             script += "\np = document.add_paragraph('Below is a summary of Data sources for the reports')";
             script += "\ntable1 = document.add_table(dftop.shape[0]+1, dftop.shape[1])";
             script += "\ntable1.style = 'TableGrid' ";
             script += "\ntable1.autofit = False";
             script += "\nfor j in range(dftop.shape[-1]):";
             script += "\n    table1.cell(0,j).text = dftop.columns[j]";
             script += "\nfor i in range(dftop.shape[0]):";
             script += "\n    for j in range(dftop.shape[-1]):";
             script += "\n        table1.cell(i+1,j).text = str(dftop.values[i,j]) ";
             script += "\nfor cell in table1.columns[0].cells:";
             script += "\n    cell.width = Inches(2)";
             script += "\nset_repeat_table_header(table1.rows[0])";
             script += "\ndocument.add_page_break()";
             script += "\ndocument.add_heading(f' {header_6}', 1)";
             script += "\np = document.add_paragraph('Below is Source Target Mapping for the reports')";
             script += "\ntable4 = document.add_table(df4.shape[0]+1,df4.shape[1])";
             script += "\ntable4.style = 'TableGrid' ";
             script += "\ntable4.autofit = True";
             script += "\ntable_cells2 = table4._cells";
             script += "\nfor i in range(df4.shape[0]):";
             script += "\n    for j in range(df4.shape[-1]):";
             script += "\n        table_cells2[j].text =  str(df4.columns[j])";
             script += "\n    for j in range(df4.shape[1]):";
             script += "\n        table_cells2[j + i * df4.shape[1]].text =  str(df4.values[i][j])";
             //script += "\nfor row in table4.rows:";
             //script += "\n    for cell in row.cells:";
             //script += "\n        paragraphs = cell.paragraphs";
             //script += "\n        for paragraph in paragraphs:";
             //script += "\n            for run in paragraph.runs:";
             //script += "\n                font = run.font";
             //script += "\n                font.size= Pt(6)";
             script += "\ndocument.add_page_break()";
             script += "\ndocument.add_heading(f' {header_7}', 1)";
             script += "\np = document.add_paragraph('Below is a summary of Calculated Columns and its expressions for the reports')";
             script += "\ntable5 = document.add_table(df5.shape[0]+1,df5.shape[1])";
             script += "\ntable5.style = 'TableGrid' ";
             script += "\ntable5.autofit = False";
             script += "\ntable_cells5 = table5._cells";
             script += "\nfor i in range(df5.shape[0]):";
             script += "\n    for j in range(df5.shape[-1]):";
             script += "\n        table_cells5[j].text =  str(df5.columns[j])";
             script += "\n    for j in range(df5.shape[1]):";
             script += "\n        table_cells5[j + i * df5.shape[1]].text =  str(df5.values[i][j])";
             script += "\nfor cell in table5.columns[0].cells:";
             script += "\n    cell.width = Inches(2)";
             script += "\ndocument.add_page_break()";
             script += "\ndocument.add_heading(f' {header_8}', 1)";
             script += "\np = document.add_paragraph('Below is a summary of Calculated Measures and its expressions for the reports')";
             script += "\ntable6 = document.add_table(df6.shape[0]+1,df6.shape[1])";
             script += "\ntable6.style = 'TableGrid' ";
             script += "\ntable6.autofit = False";
             script += "\ntable_cells6 = table6._cells";
             script += "\nfor i in range(df6.shape[0]):";
             script += "\n    for j in range(df6.shape[-1]):";
             script += "\n        table_cells6[j].text =  str(df6.columns[j])";
             script += "\n    for j in range(df6.shape[1]):";
             script += "\n        table_cells6[j + i * df6.shape[1]].text =  str(df6.values[i][j])";
             script += "\nfor cell in table6.columns[0].cells:";
             script += "\n    cell.width = Inches(2)";
             script += "\ndocument.add_page_break()";
             script += "\ndocument.add_heading(f' {header_9}', 1)";
             script += "\np = document.add_paragraph('Below is a summary of Calculated Tables and its expressions for the reports')";
             script += "\ntable7 = document.add_table(df7.shape[0]+1,df7.shape[1])";
             script += "\ntable7.style = 'TableGrid' ";
             script += "\ntable7.autofit = False";
             script += "\ntable_cells7 = table7._cells";
             script += "\nfor i in range(df7.shape[0]):";
             script += "\n    for j in range(df7.shape[-1]):";
             script += "\n        table_cells7[j].text =  str(df7.columns[j])";
             script += "\n    for j in range(df7.shape[1]):";
             script += "\n        table_cells7[j + i * df7.shape[1]].text =  str(df7.values[i][j])";
             script += "\nfor cell in table7.columns[0].cells:";
             script += "\n    cell.width = Inches(2)";
             script += "\ndocument.add_page_break()        ";
             script += "\ndocument.add_heading(f' {header_10}', 1)";
             script += "\np = document.add_paragraph('The below is a summary of Columns for the reports')";
             script += "\ntable3 = document.add_table(df3.shape[0]+1,df3.shape[1])";
             script += "\ntable3.style = 'TableGrid' ";
             script += "\ntable3.autofit = False";
             script += "\ntable_cells = table3._cells";
             script += "\nfor i in range(df3.shape[0]):";
             script += "\n    for j in range(df3.shape[-1]):";
             script += "\n        table_cells[j].text =  str(df3.columns[j])";
             script += "\n    for j in range(df3.shape[1]):";
             script += "\n        table_cells[j + i * df3.shape[1]].text =  str(df3.values[i][j])";
             script += "\nfor cell in table3.columns[0].cells:";
             script += "\n    cell.width = Inches(2)";
             script += "\ndocument.add_page_break()";
             script += "\ndocument.add_heading(f' {header_11}', 1)";
             script += "\np = document.add_paragraph('The Metadata summary of the reports in scope are defined in this document. This can be further leveraged for the migration and Rationalization.')";
             script += "\ndocument.add_page_break()";
             script += "\ndocument.add_heading(f' {header_12}', 1)";
          //   script += "\ndocument.save(\"" + DestinationPathvar + "\\\\Requirement Document For " + ComboBoxZone.SelectedValue + ".docx \")";
             script += "\ndocument.save(\"" + DestinationPathvar + "\\\\Requirement Document for All Reports.docx\")"; */
            //System.Threading.Thread.Sleep(12000);

            string script = "import os";
            script += "\nimport random";
            script += "\nimport pandas as pd";
            script += "\nimport pyodbc as py";
            script += "\nconn_str = (";
            script += "\n           r'Driver={SQL Server};'";
            script += "\n           r'Server=" + SQLServervar + ";'";
            script += "\n           r'Database=Power BI Metadata;'";
            script += "\n           r'Trusted_Connection=yes;'";
            script += "\n                   )   ";
            script += "\ncnxn = py.connect(conn_str)";
            script += "\ncursor = cnxn.cursor()";
            script += "\ndftop = pd.read_sql(\"select distinct Workspace,[Report Name],Source,[Database or Path] from dbo.vw_Metadata WHERE  [Dataset Name] NOT IN ('Internal Date Table','Internal Date Table Template' )\" , cnxn )";
            script += "\ndf2=pd.read_sql(\"select  [Report Name],  [Number of Sources], [Number of Calculated Columns], [Number of Calculated Measures], [Number of Calculated Tables], [Number of Columns] from vw_Metadata_Calculations \" , cnxn )";
            script += "\ndf3=pd.read_sql(\"select Workspace , [Report Name],COLUMN_NAME [Column Name],TABLE_NAME [Dataset],[Data Type] from dbo.vw_Metadata_Columns \" , cnxn )";
            script += "\ndf4=pd.read_sql(\"select * FROM vw_Metadata_STM \" , cnxn )";
            script += "\ndf5=pd.read_sql(\"select distinct Workspace,[Report Name],CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END AS [Dataset Name],[Column Name],[Calculated Column Expression] from dbo.Metadata where [Calculated Column Expression] is not null \" , cnxn )";
            script += "\ndf6=pd.read_sql(\"select distinct Workspace,[Report Name],CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END AS [Dataset Name],[Column Name],[Calculated Measure Expression] from dbo.Metadata where [Calculated Measure Expression] is not null \" , cnxn )";
            script += "\ndf7=pd.read_sql(\"select distinct Workspace,[Report Name],CASE WHEN [Dataset Name] like '%LocalDateTable%' THEN 'Internal Date Table' WHEN [Dataset Name] like '%DateTableTemplate%' THEN 'Internal Date Table Template' ELSE REPLACE(REPLACE([Dataset Name],'[',''),']','') END AS [Dataset Name],[Calculated Table Expression] from dbo.Metadata where [Calculated Table Expression] is not null \" , cnxn )";
            script += "\nwriter = pd.ExcelWriter(\"" + DestinationPathvar + "\\\\Requirement Document for All Reports.xlsx\", engine = 'xlsxwriter')";
            script += "\ndf2.to_excel(writer, sheet_name = 'Metadata Summary')";
            script += "\ndftop.to_excel(writer, sheet_name = 'Data sources')";
            script += "\ndf4.to_excel(writer, sheet_name = 'Source Target Mapping')";
            script += "\ndf5.to_excel(writer, sheet_name = 'Calculated Columns')";
            script += "\ndf6.to_excel(writer, sheet_name = 'Calculated Measures')";
            script += "\ndf7.to_excel(writer, sheet_name = 'Calculated Tables')";
            script += "\ndf3.to_excel(writer, sheet_name = 'Column Summary')";
            script += "\nwriter.close()";

            File.SetAttributes(path, FileAttributes.Normal);

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
                System.Threading.Thread.Sleep(10000);
                run_cmd();
                LabelServer.Visibility = Visibility.Visible;
                ComboBoxZone.Visibility = Visibility.Visible;
                LabelPythonPath.Visibility = Visibility.Visible;
                BorderPythonPath.Visibility = Visibility.Visible;
                TemplatePath.Visibility = Visibility.Visible;
                BorderTemplatePAth.Visibility = Visibility.Visible;
                DestinationPath.Visibility = Visibility.Visible;
                DestinationPathText.Visibility = Visibility.Visible;
                Browse_Copy.Visibility = Visibility.Visible;
                Template_Browse.Visibility = Visibility.Visible;
                DestPath_Browse.Visibility = Visibility.Visible;
                Home.Visibility = Visibility.Visible;
                GetReports.Visibility = Visibility.Visible;
                Info.Visibility = Visibility.Visible;
                GenerateDoc.Visibility = Visibility.Visible;
                GenerateDocAll.Visibility = Visibility.Visible;
                Animation.Visibility = Visibility.Collapsed;
            }
        }

        private void GetReports_Click(object sender, RoutedEventArgs e)
        {
            Animation.Visibility = Visibility.Visible;
            BindComboBox();
            Animation.Visibility = Visibility.Collapsed;
            TokenInfo.Text = "Generate Requirement Document for Selected Report -> To view the document for the selected report in the drop down";
            TokenInfo.AppendText(Environment.NewLine);
            TokenInfo.AppendText("Generate Requirement Document for All Reports -> To view the document for all reports");
        }

        private void SignOut_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            PowerBi window1 = new PowerBi();
            window1.ShowDialog();
        }
    }
}
