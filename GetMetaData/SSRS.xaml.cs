
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml;

namespace GetMetaData
{
    /// <summary>
    /// Interaction logic for MicStr.xaml
    /// </summary>
    public partial class SSRS : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;

        private static string PythonPath1;
        private static string TemplatePathString;
        private static string DestinationPathString;
        string pythout = "";
        int XMLCnt = 0;
        public SSRS()
        {
            InitializeComponent();
            //  ProcessStart.Visibility = Visibility.Collapsed;
            GeneratePBI.Visibility = Visibility.Visible;
            GenerateDoc.Visibility = Visibility.Visible;
            Labelusername.Visibility = Visibility.Collapsed;
            Borderusername.Visibility = Visibility.Collapsed;
            Labelpasswd.Visibility = Visibility.Collapsed;
            Borderpasswd.Visibility = Visibility.Collapsed;
            PasswordChek.Visibility = Visibility.Collapsed;
            InsertXML.Visibility = Visibility.Visible;
            ProcessImage.Visibility = Visibility.Collapsed;
            OutputImage.Visibility = Visibility.Collapsed;
            DocImage.Visibility = Visibility.Collapsed;
            WindRad.IsChecked = true;
            DBConnection.IsChecked = true;
            CSVpath.IsEnabled = false;


            LabelUser.Margin = new Thickness(20, 90, 0, 0);
            Stack1.Margin = new Thickness(20, 100, 0, 0);
            LabelUser2.Visibility = Visibility.Collapsed;
            DBLabelUser2.Visibility = Visibility.Visible;
            PasswordChek1.Visibility = Visibility.Visible;


            HostName.IsReadOnly = false;
            UserName.IsReadOnly = true;
            DataBaseName.IsReadOnly = true;
            Password1.IsEnabled = false;
            PasswordShow1.IsReadOnly = true;
            PasswordChek1.IsEnabled = false;
            AuthRad.IsEnabled = false;


            InitializeComponent();
            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            MyNotifyIcon.Icon = new System.Drawing.Icon(
                            @"Final.ico");
            MyNotifyIcon.MouseDoubleClick +=
                new System.Windows.Forms.MouseEventHandler(MyNotifyIcon_MouseDoubleClick);

            MyNotifyIcon.BalloonTipClicked += new EventHandler(MyNotifyIcon_BalloonTipClicked);
            MyNotifyIcon.BalloonTipShown += new EventHandler(MyNotifyIcon_BalloonTipClicked);
            MyNotifyIcon.BalloonTipClosed += new EventHandler(MyNotifyIcon_BalloonTipClicked);

        }
        void MyNotifyIcon_BalloonTipClicked(object sender, EventArgs e)
        {
            this.WindowState = WindowState.Maximized;
            MyNotifyIcon.Visible = false;
        }
        private void SignOut_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            Window1 window1 = new Window1();
            window1.ShowDialog();
        }
        private void Close_Click(object sender, RoutedEventArgs e)
        {
            // ShowInTaskbar = true;
            this.Close();
            Application.Current.Shutdown();
            //this.Hide();
        }

        

        private void Minimize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
            this.ShowInTaskbar = true;
            MyNotifyIcon.Visible = true;
            //MyNotifyIcon.Text = "";
            MyNotifyIcon.BalloonTipTitle = "Minimize Sucessful";
            MyNotifyIcon.BalloonTipText = "Minimized the app ";
            MyNotifyIcon.ShowBalloonTip(5000);
            //Thread.Sleep(5000);
            //MyNotifyIcon.Dispose();
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
            TextPython.Text = dialog.SelectedPath;
            PythonPath1 = TextPython.Text;
        }

        private void Template_Browse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.ShowDialog();
            TextCSV.Text = dialog.SelectedPath;
            PythonPath1 = TextCSV.Text;
        }

        private void GeneratePBI_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Generating Report. Please Wait ....");
            // string file = @"Metadata Output.pbix";
            string fileName = "BI4BI - SSRS.pbix";
            string path = System.IO.Path.Combine(Environment.CurrentDirectory, @"Report\", fileName);
            Process.Start(path);
        }

        private void GenerateDoc_Click(object sender, RoutedEventArgs e)
        {
            int result = 0;

            string connectionstring = "Data Source=" + Source123.Text.ToString() + "; Integrated Security=true; Initial Catalog=SSRS Metadata"; ; //your connectionstring    

            if (Source123.Text.Equals(""))
            {
                MessageBox.Show("Click Load Data to populate the data base-> Then click the document generator");
            }
            else
            {
                using (SqlConnection conn = new SqlConnection(connectionstring))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("select COUNT(*) from dbo.ReportInventory", conn);
                    result = (int)cmd.ExecuteScalar();
                    conn.Close();
                }
                if (Source123.Text.ToString().Equals("") || result == 0)
                {
                    MessageBox.Show("Either the Metadata is not extracted or the SQL Server details is blank");
                }
                else
                {
                    Document_Generator_SSRS objWelcome = new Document_Generator_SSRS();
                    objWelcome.SQLTB_DGMS.Text = Source123.Text;
                    objWelcome.Show(); //Sending value from one form to another form.
                    Close();
                }
            }
        }

        private void GenerateMetadata_Checked(object sender, RoutedEventArgs e)
        {
            if (XMLCnt == 0)
            {
                InsertXML.Visibility = Visibility.Visible;
                GeneratePBI.Visibility = Visibility.Visible;
                GenerateDoc.Visibility = Visibility.Visible;
                //    ProcessStart.Visibility = Visibility.Collapsed;
            }
            else
            {
                InsertXML.Visibility = Visibility.Visible;
                GeneratePBI.Visibility = Visibility.Visible;
                GenerateDoc.Visibility = Visibility.Visible;
                //    ProcessStart.Visibility = Visibility.Visible;
            }
        }

        private void Output_Checked(object sender, RoutedEventArgs e)
        {
            InsertXML.Visibility = Visibility.Collapsed;
            //      ProcessStart.Visibility = Visibility.Collapsed;
            GeneratePBI.Visibility = Visibility.Visible;
            GenerateDoc.Visibility = Visibility.Visible;
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            PasswordShow.Text = Password.Password;
            BorderPasswordShow.Visibility = Visibility.Visible;
            Borderpasswd.Visibility = Visibility.Collapsed;
        }
        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Password.Password = PasswordShow.Text;
            BorderPasswordShow.Visibility = Visibility.Collapsed;
            Borderpasswd.Visibility = Visibility.Visible;
        }


        private void WindAuth_Checked(object sender, RoutedEventArgs e)
        {
            Labelusername.Visibility = Visibility.Collapsed;
            Borderusername.Visibility = Visibility.Collapsed;
            Labelpasswd.Visibility = Visibility.Collapsed;
            Borderpasswd.Visibility = Visibility.Collapsed;
            PasswordChek.Visibility = Visibility.Collapsed;
        }

        private void SQL_Checked(object sender, RoutedEventArgs e)
        {
            Labelusername.Visibility = Visibility.Visible;
            Borderusername.Visibility = Visibility.Visible;
            Labelpasswd.Visibility = Visibility.Visible;
            Borderpasswd.Visibility = Visibility.Visible;
            PasswordChek.Visibility = Visibility.Visible;
        }

        private void CSV_SQL_Checked(object sender, RoutedEventArgs e)
        {
            LabelUser.Margin = new Thickness(20, 7, 0, 0);
            Stack1.Margin = new Thickness(20, 15, 0, 0);
            LabelUser2.Visibility = Visibility.Visible;
            DBLabelUser2.Visibility = Visibility.Collapsed;
            PasswordChek1.Visibility = Visibility.Collapsed;

            LabelUser1.Visibility = Visibility.Collapsed;
            LabelHostName.Visibility = Visibility.Collapsed;
            BorderHosName.Visibility = Visibility.Collapsed;
            HostName.Visibility = Visibility.Collapsed;

            LabelDataBaseName.Visibility = Visibility.Collapsed;
            BorderDataBaseName.Visibility = Visibility.Collapsed;
            DataBaseName.Visibility = Visibility.Collapsed;

            LabelUserName.Visibility = Visibility.Collapsed;
            BorderUserName.Visibility = Visibility.Collapsed;
            UserName.Visibility = Visibility.Collapsed;

            LabelPassword.Visibility = Visibility.Collapsed;
            BorderPassword1.Visibility = Visibility.Collapsed;
            Password1.Visibility = Visibility.Collapsed;


            LabelUser.Visibility = Visibility.Visible;
            TemplatePath.Visibility = Visibility.Visible;
            BorderTemplatePAth.Visibility = Visibility.Visible;
            TextCSV.Visibility = Visibility.Visible;
            Template_Browse.Visibility = Visibility.Visible;

            HostName.Clear();
            UserName.Clear();
            DataBaseName.Clear();
            Password1.Clear();
            PasswordShow1.Clear();
            PasswordChek1.IsChecked = false;
            BorderPasswordShow1.Visibility = Visibility.Collapsed;
            BorderPassword1.Visibility = Visibility.Collapsed;
            TextCSV.Clear();
            Source123.Clear();
            username.Clear();
            Password.Clear();
            PasswordShow.Clear();
            TextPython.Clear();
            PasswordChek.IsChecked = false;
            WindRad.IsChecked = true;


        }

        private void DB_Checked(object sender, RoutedEventArgs e)
        {
            LabelUser.Margin = new Thickness(20, 100, 0, 0);
            Stack1.Margin = new Thickness(20, 100, 0, 0);
            LabelUser2.Visibility = Visibility.Collapsed;
            DBLabelUser2.Visibility = Visibility.Visible;
            PasswordChek1.Visibility = Visibility.Visible;

            LabelUser.Visibility = Visibility.Collapsed;
            TemplatePath.Visibility = Visibility.Collapsed;
            BorderTemplatePAth.Visibility = Visibility.Collapsed;
            TextCSV.Visibility = Visibility.Collapsed;
            Template_Browse.Visibility = Visibility.Collapsed;

            LabelUser1.Visibility = Visibility.Visible;
            LabelHostName.Visibility = Visibility.Visible;
            BorderHosName.Visibility = Visibility.Visible;
            HostName.Visibility = Visibility.Visible;

            LabelDataBaseName.Visibility = Visibility.Visible;
            BorderDataBaseName.Visibility = Visibility.Visible;
            DataBaseName.Visibility = Visibility.Visible;

            LabelUserName.Visibility = Visibility.Visible;
            BorderUserName.Visibility = Visibility.Visible;
            UserName.Visibility = Visibility.Visible;

            LabelPassword.Visibility = Visibility.Visible;
            BorderPassword1.Visibility = Visibility.Visible;
            Password1.Visibility = Visibility.Visible;

            TextCSV.Clear();
            Source123.Clear();
            username.Clear();
            Password.Clear();
            PasswordShow.Clear();
            TextPython.Clear();
            PasswordChek.IsChecked = false;
            WindRad.IsChecked = true;

        }

        private void DBCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            PasswordShow1.Text = Password1.Password;
            BorderPasswordShow1.Visibility = Visibility.Visible;
            BorderPassword1.Visibility = Visibility.Collapsed;
        }
        private void DBCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Password1.Password = PasswordShow1.Text;
            BorderPasswordShow1.Visibility = Visibility.Collapsed;
            BorderPassword1.Visibility = Visibility.Visible;
        }

        private void SignOutButton_Click(object sender, RoutedEventArgs e)
        {

            this.Close();
            Window1 window1 = new Window1();
            window1.ShowDialog();

        }

        private void InsertXML_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(HostName.Text.ToString()) || String.IsNullOrEmpty(Source123.Text.ToString()) || String.IsNullOrEmpty(TextPython.Text.ToString()))
            {
                MessageBox.Show("Please enter the mandatory fields and try again");
            }
            else { 
            string path = Directory.GetCurrentDirectory() + @"\PythonFile\SSRS_Process_Python.py";
            string password = "";
            if (PasswordChek.IsChecked == true)
            {
                password = PasswordShow.Text.ToString();
            }
            else
            {
                password = Password.Password;
            }
            if (WindRad.IsChecked == true)
            {

                TemplatePath.Visibility = Visibility.Collapsed;
                BorderTemplatePAth.Visibility = Visibility.Collapsed;
                Template_Browse.Visibility = Visibility.Collapsed;
                LabelSource.Visibility = Visibility.Collapsed;
                BorderSource.Visibility = Visibility.Collapsed;
                WindRad.Visibility = Visibility.Collapsed;
                AuthRad.Visibility = Visibility.Collapsed;
                Labelusername.Visibility = Visibility.Collapsed;
                Borderusername.Visibility = Visibility.Collapsed;
                Labelpasswd.Visibility = Visibility.Collapsed;
                Borderpasswd.Visibility = Visibility.Collapsed;
                BorderPasswordShow.Visibility = Visibility.Collapsed;
                PasswordChek.Visibility = Visibility.Collapsed;
                LabelDatabaseServer.Visibility = Visibility.Collapsed;
                BorderServer.Visibility = Visibility.Collapsed;
                Browse.Visibility = Visibility.Collapsed;
                //  GenerateMetadata.Visibility = Visibility.Collapsed;
                GeneratePBI.Visibility = Visibility.Collapsed;
                GenerateDoc.Visibility = Visibility.Collapsed;
                InsertXML.Visibility = Visibility.Collapsed;
                //  ProcessStart.Visibility = Visibility.Collapsed;
                ProcessImage.Visibility = Visibility.Collapsed;
                OutputImage.Visibility = Visibility.Collapsed;
                DocImage.Visibility = Visibility.Collapsed;
                Animation.Visibility = Visibility.Visible;

                LabelUser2.Visibility = Visibility.Collapsed;
                DBLabelUser2.Visibility = Visibility.Visible;

                PasswordChek1.Visibility = Visibility.Collapsed;

                LabelUser1.Visibility = Visibility.Collapsed;
                LabelHostName.Visibility = Visibility.Collapsed;
                BorderHosName.Visibility = Visibility.Collapsed;
                HostName.Visibility = Visibility.Collapsed;

                LabelDataBaseName.Visibility = Visibility.Collapsed;
                BorderDataBaseName.Visibility = Visibility.Collapsed;
                DataBaseName.Visibility = Visibility.Collapsed;

                LabelUserName.Visibility = Visibility.Collapsed;
                BorderUserName.Visibility = Visibility.Collapsed;
                UserName.Visibility = Visibility.Collapsed;

                LabelPassword.Visibility = Visibility.Collapsed;
                BorderPassword1.Visibility = Visibility.Collapsed;
                Password1.Visibility = Visibility.Collapsed;
                DBConnection.Visibility = Visibility.Collapsed;
                CSVpath.Visibility = Visibility.Collapsed;


                MessageBox.Show("Loading data into " + Source123.Text.ToString());
                
                    string script = "import urllib";
                    script += "\nimport pandas as pd                                                                                                                                                                                                           ";
                    script += "\nimport xmltodict                                                                                                                                                                                                              ";
                    script += "\nimport re                                                                                                                                                                                                                     ";
                    script += "\nimport numpy as np                                                                                                                                                                                                            ";
                    script += "\nfrom sqlalchemy import create_engine                                                                                                                                                                                          ";
                    script += "\nimport pyodbc                                                                                                                                                                                                                 ";
                    script += "\nconn_str = (\"DRIVER={SQL Server Native Client 11.0};SERVER=" + HostName.Text.ToString() + ";DATABASE=ReportServer;Trusted_Connection=yes;\")";
                    script += "\ncnxn = pyodbc.connect(conn_str)                                                                                                                                                                                               ";
                    script += "\ncursor = cnxn.cursor()                                                                                                                                                                                                        ";
                    script += "\ncatalog_data = pd.read_sql('select * from Catalog' , cnxn)                                                                                                                                                                    ";
                    script += "\ninventory_query = '''                                                                                                                                                                                                         ";
                    script += "\nSELECT ReportMetrics.*,ReportParameter.ParameterName,ReportParameter.Prompt,ReportParameter.DataType,                                                                                                                         ";
                    script += "\nReportConnectionString.LocalDataSourceName,ReportConnectionString.SharedDataSourceName,ReportConnectionString.SharedDataSource,                                                                                               ";
                    script += "\nReportConnectionString.DataProvider,ReportConnectionString.ConnectionString                                                                                                                                                   ";
                    script += "\nFROM                                                                                                                                                                                                                          ";
                    script += "\n(SELECT [Catalog].[ItemID],                                                                                                                                                                                                   ";
                    script += "\n[Catalog].[Path],                                                                                                                                                                                                             ";
                    script += "\n[Catalog].[Name],                                                                                                                                                                                                             ";
                    script += "\n[Catalog].[ParentID],                                                                                                                                                                                                         ";
                    script += "\n[Catalog].[Type],                                                                                                                                                                                                             ";
                    script += "\nCASE [Catalog].[Type] --Type, an int which can be converted using this case statement.                                                                                                                                        ";
                    script += "\n    WHEN 1 THEN 'Folder'                                                                                                                                                                                                      ";
                    script += "\n    WHEN 2 THEN 'Report'                                                                                                                                                                                                      ";
                    script += "\n    WHEN 3 THEN 'File'                                                                                                                                                                                                        ";
                    script += "\n    WHEN 4 THEN 'Linked Report'                                                                                                                                                                                               ";
                    script += "\n    WHEN 5 THEN 'Data Source'                                                                                                                                                                                                 ";
                    script += "\n    WHEN 6 THEN 'Report Model'                                                                                                                                                                                                ";
                    script += "\n    WHEN 7 THEN 'Report Part'                                                                                                                                                                                                 ";
                    script += "\n    WHEN 8 THEN 'Shared Data Set'                                                                                                                                                                                             ";
                    script += "\n    WHEN 9 THEN 'Image'                                                                                                                                                                                                       ";
                    script += "\n    ELSE CAST([Type] as varchar(100))                                                                                                                                                                                         ";
                    script += "\n  END AS TypeName,                                                                                                                                                                                                            ";
                    script += "\n[Catalog].[Description],                                                                                                                                                                                                      ";
                    script += "\n[Catalog].[Hidden],                                                                                                                                                                                                           ";
                    script += "\nCase [Catalog].[Hidden]                                                                                                                                                                                                       ";
                    script += "\n	  WHEN NULL THEN 'No'                                                                                                                                                                                                      ";
                    script += "\n	  ELSE 'Yes'END as HiddenDesc,                                                                                                                                                                                             ";
                    script += "\n[CreatedUser].[UserName] As Created_User_Name,                                                                                                                                                                                ";
                    script += "\n[Catalog].[CreationDate],                                                                                                                                                                                                     ";
                    script += "\n[ModifiedUser].[UserName] As Modified_User_Name,                                                                                                                                                                              ";
                    script += "\n[Catalog].[ModifiedDate],                                                                                                                                                                                                     ";
                    script += "\n[Catalog].[ExecutionFlag],                                                                                                                                                                                                    ";
                    script += "\nCASE [Catalog].[ExecutionFlag]                                                                                                                                                                                                ";
                    script += "\n	  WHEN 1 THEN 'Yes'                                                                                                                                                                                                        ";
                    script += "\n	  ELSE 'No' END as ExecutionFlagDescription,                                                                                                                                                                               ";
                    script += "\n[Catalog].[ContentSize],                                                                                                                                                                                                      ";
                    script += "\nCAST(CAST([Content] AS VARBINARY(MAX)) AS XML) AS ReportXML,                                                                                                                                                                  ";
                    script += "\n[DataSource].[DSID] As DataSourceID,                                                                                                                                                                                          ";
                    script += "\n[DataSource].[Name] AS DataSourceName,                                                                                                                                                                                        ";
                    script += "\n[DataSource].[Extension],                                                                                                                                                                                                     ";
                    script += "\n[DataSource].[Version],                                                                                                                                                                                                       ";
                    script += "\n[DataSets].[ID] AS DataSetID,                                                                                                                                                                                                 ";
                    script += "\n[DataSets].[Name] AS DataSetName                                                                                                                                                                                              ";
                    script += "\nFROM [ReportServer].[dbo].[Catalog]                                                                                                                                                                                           ";
                    script += "\nINNER JOIN (SELECT * FROM [ReportServer].[dbo].[Users]) CreatedUser ON CreatedUser.UserID=[Catalog].[CreatedByID]                                                                                                             ";
                    script += "\nINNER JOIN (SELECT * FROM [ReportServer].[dbo].[Users]) ModifiedUser ON ModifiedUser.UserID=[Catalog].[ModifiedByID]                                                                                                          ";
                    script += "\nLEFT JOIN [ReportServer].[dbo].[DataSource] ON [DataSource].[ItemID] = [Catalog].[ItemID]                                                                                                                                     ";
                    script += "\nLEFT JOIN [ReportServer].[dbo].[DataSets] ON [DataSets].[ItemID] = [Catalog].[ItemID]                                                                                                                                         ";
                    script += "\nWHERE [Catalog].[Type] = 2 /* exclude down to only showing reports */                                                                                                                                                         ";
                    script += "\n) ReportMetrics                                                                                                                                                                                                               ";
                    script += "\nLEFT OUTER JOIN                                                                                                                                                                                                               ";
                    script += "\n(                                                                                                                                                                                                                             ";
                    script += "\nSELECT                                                                                                                                                                                                                        ";
                    script += "\n        Cat.ItemID, cat.[Path], cat.Name                                                                                                                                                                                      ";
                    script += "\n        , p.ParameterName,p.Prompt,p.DataType                                                                                                                                                                                 ";
                    script += "\n    FROM ReportServer.dbo.Catalog cat                                                                                                                                                                                         ";
                    script += "\n        JOIN (                                                                                                                                                                                                                ";
                    script += "\n                SELECT ReportID = ItemID                                                                                                                                                                                      ";
                    script += "\n                                ,ParameterName = params.value('(Name/text())[1]', 'varchar(100)')                                                                                                                             ";
                    script += "\n                                ,Prompt = params.value('(Prompt/text())[1]', 'nvarchar(100)')                                                                                                                                 ";
                    script += "\n                                ,DataType = params.value('(Type/text())[1]', 'varchar(100)')                                                                                                                                  ";
                    script += "\n                FROM (                                                                                                                                                                                                        ";
                    script += "\n                                SELECT C.ItemID, C.Name,CONVERT(XML,C.Parameter) AS ParameterXML                                                                                                                              ";
                    script += "\n                                FROM  ReportServer.dbo.Catalog C                                                                                                                                                              ";
                    script += "\n                                WHERE  C.Content is not null                                                                                                                                                                  ";
                    script += "\n                                AND  C.Type  = 2                                                                                                                                                                              ";
                    script += "\n                                ) a                                                                                                                                                                                           ";
                    script += "\n                cross apply ParameterXML.nodes('//Parameters/Parameter') q (params)                                                                                                                                           ";
                    script += "\n        ) p on cat.ItemID = p.ReportID                                                                                                                                                                                        ";
                    script += "\n) ReportParameter ON ReportParameter.ItemID=ReportMetrics.ItemID                                                                                                                                                              ";
                    script += "\nLEFT OUTER JOIN                                                                                                                                                                                                               ";
                    script += "\n(                                                                                                                                                                                                                             ";
                    script += "\nSELECT * from  (                                                                                                                                                                                                              ";
                    script += "\n    SELECT r.ItemID,                                                                                                                                                                                                          ";
                    script += "\n        r.LocalDataSourceName, -- embedded data source's name or local name given to shared data source                                                                                                                       ";
                    script += "\n        sds.SharedDataSourceName,                                                                                                                                                                                             ";
                    script += "\n        SharedDataSource = CAST ((CASE WHEN sds.SharedDataSourceName IS NOT NULL THEN 1 ELSE 0 END) AS BIT),                                                                                                                  ";
                    script += "\n        DataProvider = ISNULL(r.DataProvider, sds.DataProvider),                                                                                                                                                              ";
                    script += "\n        ConnectionString = ISNULL(r.ConnectionString, sds.ConnectionString)                                                                                                                                                   ";
                    script += "\n    FROM (                                                                                                                                                                                                                    ";
                    script += "\n        SELECT c.*,                                                                                                                                                                                                           ";
                    script += "\n                LocalDataSourceName = DataSourceXml.value('@Name', 'NVARCHAR(260)'),                                                                                                                                          ";
                    script += "\n                DataProvider = DataSourceXml.value('(*:ConnectionProperties/*:DataProvider)[1]', 'NVARCHAR(260)'),                                                                                                            ";
                    script += "\n                ConnectionString = DataSourceXml.value('(*:ConnectionProperties/*:ConnectString)[1]', 'NVARCHAR(MAX)')                                                                                                        ";
                    script += "\n            FROM (                                                                                                                                                                                                            ";
                    script += "\n		SELECT *,                                                                                                                                                                                                              ";
                    script += "\n         ContentXml = (CONVERT(XML, CONVERT(VARBINARY(MAX), Content)))                                                                                                                                                        ";
                    script += "\n    FROM Catalog                                                                                                                                                                                                              ";
                    script += "\n		) c                                                                                                                                                                                                                    ";
                    script += "\n                CROSS APPLY ContentXml.nodes('/*:Report/*:DataSources/*:DataSource') DataSource(DataSourceXml)                                                                                                                ";
                    script += "\n            WHERE c.Type = 2 -- limit to reports only                                                                                                                                                                         ";
                    script += "\n        ) r                                                                                                                                                                                                                   ";
                    script += "\n        LEFT JOIN                                                                                                                                                                                                             ";
                    script += "\n		(                                                                                                                                                                                                                      ";
                    script += "\n    SELECT ds.ItemID,                                                                                                                                                                                                         ";
                    script += "\n        SharedDataSourceName = c.Name,                                                                                                                                                                                        ";
                    script += "\n        LocalDataSourceName = ds.Name,                                                                                                                                                                                        ";
                    script += "\n        DataProvider = ContentXML.value('(/*:DataSourceDefinition/*:Extension)[1]', 'NVARCHAR(260)'),                                                                                                                         ";
                    script += "\n        ConnectionString = ContentXML.value('(/*:DataSourceDefinition/*:ConnectString)[1]', 'NVARCHAR(MAX)')                                                                                                                  ";
                    script += "\n    FROM DataSource ds                                                                                                                                                                                                        ";
                    script += "\n        JOIN                                                                                                                                                                                                                  ";
                    script += "\n		(                                                                                                                                                                                                                      ";
                    script += "\n		SELECT *,                                                                                                                                                                                                              ";
                    script += "\n         ContentXml = (CONVERT(XML, CONVERT(VARBINARY(MAX), Content)))                                                                                                                                                        ";
                    script += "\n    FROM Catalog                                                                                                                                                                                                              ";
                    script += "\n		) c ON ds.Link = c.ItemID                                                                                                                                                                                              ";
                    script += "\n)sds ON r.ItemID = sds.ItemID AND r.LocalDataSourceName = sds.LocalDataSourceName                                                                                                                                             ";
                    script += "\n) AllDataSources                                                                                                                                                                                                              ";
                    script += "\n) ReportConnectionString on ReportMetrics.ItemID=ReportConnectionString.ItemID                                                                                                                                                ";
                    script += "\n'''                                                                                                                                                                                                                           ";
                    script += "\nexecution_metrics_query = '''                                                                                                                                                                                                 ";
                    script += "\nSELECT [ExecutionLogStorage].[InstanceName],                                                                                                                                                                                  ";
                    script += "\n[ExecutionLogStorage].[ReportID],                                                                                                                                                                                             ";
                    script += "\n[Catalog].[Name] As ReportName,                                                                                                                                                                                               ";
                    script += "\nCOALESCE(                                                                                                                                                                                                                     ";
                    script += "\nCASE [ExecutionLogStorage].[ReportAction]                                                                                                                                                                                     ";
                    script += "\n        WHEN 11 THEN AdditionalInfo.value('(AdditionalInfo/SourceReportUri)[1]', 'nvarchar(max)')                                                                                                                             ";
                    script += "\n        ELSE [Catalog].[Path]                                                                                                                                                                                                 ";
                    script += "\n        END                                                                                                                                                                                                                   ";
                    script += "\n		, 'Unknown')                                                                                                                                                                                                           ";
                    script += "\n		AS ItemPath,                                                                                                                                                                                                           ";
                    script += "\n[ExecutionLogStorage].[UserName],                                                                                                                                                                                             ";
                    script += "\n[ExecutionLogStorage].[ExecutionId],                                                                                                                                                                                          ";
                    script += "\n[ExecutionLogStorage].[RequestType],                                                                                                                                                                                          ";
                    script += "\nCASE [ExecutionLogStorage].[RequestType]                                                                                                                                                                                      ";
                    script += "\n        WHEN 0 THEN 'Interactive'                                                                                                                                                                                             ";
                    script += "\n        WHEN 1 THEN 'Subscription'                                                                                                                                                                                            ";
                    script += "\n        WHEN 2 THEN 'Refresh Cache'                                                                                                                                                                                           ";
                    script += "\n        ELSE 'Unknown'                                                                                                                                                                                                        ";
                    script += "\n        END AS RequestTypeDesc,                                                                                                                                                                                               ";
                    script += "\n[ExecutionLogStorage].[Format],                                                                                                                                                                                               ";
                    script += "\n[ExecutionLogStorage].[Parameters],                                                                                                                                                                                           ";
                    script += "\n[ExecutionLogStorage].[ReportAction],                                                                                                                                                                                         ";
                    script += "\n  CASE [ExecutionLogStorage].[ReportAction]                                                                                                                                                                                   ";
                    script += "\n        WHEN 1 THEN 'Render'                                                                                                                                                                                                  ";
                    script += "\n        WHEN 2 THEN 'BookmarkNavigation'                                                                                                                                                                                      ";
                    script += "\n        WHEN 3 THEN 'DocumentMapNavigation'                                                                                                                                                                                   ";
                    script += "\n        WHEN 4 THEN 'DrillThrough'                                                                                                                                                                                            ";
                    script += "\n        WHEN 5 THEN 'FindString'                                                                                                                                                                                              ";
                    script += "\n        WHEN 6 THEN 'GetDocumentMap'                                                                                                                                                                                          ";
                    script += "\n        WHEN 7 THEN 'Toggle'                                                                                                                                                                                                  ";
                    script += "\n        WHEN 8 THEN 'Sort'                                                                                                                                                                                                    ";
                    script += "\n        WHEN 9 THEN 'Execute'                                                                                                                                                                                                 ";
                    script += "\n        WHEN 10 THEN 'RenderEdit'                                                                                                                                                                                             ";
                    script += "\n        WHEN 11 THEN 'ExecuteDataShapeQuery'                                                                                                                                                                                  ";
                    script += "\n        WHEN 12 THEN 'RenderMobileReport'                                                                                                                                                                                     ";
                    script += "\n        WHEN 13 THEN 'ConceptualSchema'                                                                                                                                                                                       ";
                    script += "\n        WHEN 14 THEN 'QueryData'                                                                                                                                                                                              ";
                    script += "\n        WHEN 15 THEN 'ASModelStream'                                                                                                                                                                                          ";
                    script += "\n        WHEN 16 THEN 'RenderExcelWorkbook'                                                                                                                                                                                    ";
                    script += "\n        WHEN 17 THEN 'GetExcelWorkbookInfo'                                                                                                                                                                                   ";
                    script += "\n        WHEN 18 THEN 'SaveToCatalog'                                                                                                                                                                                          ";
                    script += "\n        WHEN 19 THEN 'DataRefresh'                                                                                                                                                                                            ";
                    script += "\n        ELSE 'Unknown'                                                                                                                                                                                                        ";
                    script += "\n        END AS ReportActionDesc,                                                                                                                                                                                              ";
                    script += "\n[ExecutionLogStorage].[TimeStart],                                                                                                                                                                                            ";
                    script += "\n[ExecutionLogStorage].[TimeEnd],                                                                                                                                                                                              ";
                    script += "\n[ExecutionLogStorage].[TimeDataRetrieval],                                                                                                                                                                                    ";
                    script += "\n[ExecutionLogStorage].[TimeProcessing],                                                                                                                                                                                       ";
                    script += "\n[ExecutionLogStorage].[TimeRendering],                                                                                                                                                                                        ";
                    script += "\n[ExecutionLogStorage].[Source],                                                                                                                                                                                               ";
                    script += "\nCASE [ExecutionLogStorage].[Source]                                                                                                                                                                                           ";
                    script += "\n        WHEN 1 THEN 'Live'                                                                                                                                                                                                    ";
                    script += "\n        WHEN 2 THEN 'Cache'                                                                                                                                                                                                   ";
                    script += "\n        WHEN 3 THEN 'Snapshot'                                                                                                                                                                                                ";
                    script += "\n        WHEN 4 THEN 'History'                                                                                                                                                                                                 ";
                    script += "\n        WHEN 5 THEN 'AdHoc'                                                                                                                                                                                                   ";
                    script += "\n        WHEN 6 THEN 'Session'                                                                                                                                                                                                 ";
                    script += "\n        WHEN 7 THEN 'Rdce'                                                                                                                                                                                                    ";
                    script += "\n        ELSE 'Unknown'                                                                                                                                                                                                        ";
                    script += "\n        END AS SourceDesc,                                                                                                                                                                                                    ";
                    script += "\n[ExecutionLogStorage].[Status],                                                                                                                                                                                               ";
                    script += "\n[ExecutionLogStorage].[ByteCount],                                                                                                                                                                                            ";
                    script += "\n[ExecutionLogStorage].[RowCount],                                                                                                                                                                                             ";
                    script += "\n[ExecutionLogStorage].[AdditionalInfo]                                                                                                                                                                                        ";
                    script += "\nFROM [ReportServer].[dbo].[ExecutionLogStorage]                                                                                                                                                                               ";
                    script += "\nLEFT JOIN [ReportServer].[dbo].[Catalog] ON [ExecutionLogStorage].[ReportID] = [Catalog].[ItemID]                                                                                                                             ";
                    script += "\nWHERE [Catalog].[Type] = 2 /* exclude down to only showing reports */                                                                                                                                                         ";
                    script += "\n'''                                                                                                                                                                                                                           ";
                    script += "\nexecution_metrics_data = pd.read_sql(execution_metrics_query, cnxn)                                                                                                                                                           ";
                    script += "\ninventory_data = pd.read_sql(inventory_query, cnxn)                                                                                                                                                                           ";
                    script += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=SSRS Metadata;Trusted_Connection=yes;\")";
                    script += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted), fast_executemany=True)                                                                                                                              ";
                    script += "\ncatalog_data.to_sql('Catalog', schema='dbo',if_exists = 'append', con = engine, index=False)                                                                                                                                  ";
                    script += "\ninventory_data.to_sql('ReportInventory', schema='dbo',if_exists = 'append', con = engine, index=False)                                                                                                                        ";
                    script += "\nexecution_metrics_data.to_sql('ExecutionMetrics', schema='dbo',if_exists = 'append', con = engine, index=False)                                                                                                               ";
                    script += "\ncnxn.close()                                                                                                                                                                                                                  ";
                    script += "\nconn_str = (\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=SSRS Metadata;Trusted_Connection=yes;\")";
                    script += "\ncnxn = pyodbc.connect(conn_str)                                                                                                                                                                                               ";
                    script += "\ncursor = cnxn.cursor()                                                                                                                                                                                                        ";
                    script += "\ncatalog_data = pd.read_sql('select * from Catalog', cnxn)                                                                                                                                                                     ";
                    script += "\ncatalog_data = catalog_data.drop_duplicates()                                                                                                                                                                                 ";
                    script += "\ncatalog_data.to_sql('Catalog', schema='dbo', if_exists = 'replace', con = engine, index=False)                                                                                                                                ";
                    script += "\ninventory_data = pd.read_sql('select * from ReportInventory', cnxn)                                                                                                                                                           ";
                    script += "\ninventory_data = inventory_data.drop_duplicates()                                                                                                                                                                             ";
                    script += "\ninventory_data.to_sql('ReportInventory', schema='dbo', if_exists = 'replace', con = engine, index=False)                                                                                                                      ";
                    script += "\nexecution_metrics_data = pd.read_sql('select * from ExecutionMetrics', cnxn)                                                                                                                                                  ";
                    script += "\nexecution_metrics_data = execution_metrics_data.drop_duplicates()                                                                                                                                                             ";
                    script += "\nexecution_metrics_data.to_sql('ExecutionMetrics', schema='dbo', if_exists = 'replace', con = engine, index=False)                                                                                                             ";
                    script += "\nbase_query = '''                                                                                                                                                                                                              ";
                    script += "\nselect Name, Type, Content , CAST(CAST(Content as VARBINARY(MAX)) AS XML) AS ReportXML                                                                                                                                        ";
                    script += "\n  from Catalog                                                                                                                                                                                                                ";
                    script += "\n  where Type=2                                                                                                                                                                                                                ";
                    script += "\n'''                                                                                                                                                                                                                           ";
                    script += "\ndf_data = pd.read_sql(base_query , cnxn)                                                                                                                                                                                      ";
                    script += "\ndef get_table(query):                                                                                                                                                                                                         ";
                    script += "\n    db,schema,table = None,None,None                                                                                                                                                                                          ";
                    script += "\n    if re.search(r'FROM\\s+\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?',query,re.IGNORECASE):                                                                                                                                        ";
                    script += "\n        db,schema,table = re.findall(r'FROM\\s+\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?',query,re.IGNORECASE)[-1]                                                                                                                 ";
                    script += "\n    elif re.search(r'FROM\\s+\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?',query,re.IGNORECASE):                                                                                                                                                   ";
                    script += "\n        db,schema = None,None                                                                                                                                                                                                 ";
                    script += "\n        table = re.findall(r'FROM\\s+\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?',query,re.IGNORECASE)[-1][-1]                                                                                                                                    ";
                    script += "\n    return db,schema,table                                                                                                                                                                                                    ";
                    script += "\ndef get_data(ds):                                                                                                                                                                                                             ";
                    script += "\n    ds_type, db_name, schema, table_name, query = None,None,None,None,None                                                                                                                                                    ";
                    script += "\n    if 'SharedDataSet' in ds:                                                                                                                                                                                                 ";
                    script += "\n        ds_type = 'Shared dataset'                                                                                                                                                                                            ";
                    script += "\n        db_name,schema,table_name = None,None,None                                                                                                                                                                            ";
                    script += "\n    elif ds['Query'].get('CommandType') and ds['Query']['CommandType'] == 'StoredProcedure':                                                                                                                                  ";
                    script += "\n        ds_type = 'StoredProcedure'                                                                                                                                                                                           ";
                    script += "\n        query = ds['Query']['CommandText']                                                                                                                                                                                    ";
                    script += "\n        table_name = query                                                                                                                                                                                                    ";
                    script += "\n    elif 'EXEC' in ds['Query']['CommandText']:                                                                                                                                                                                ";
                    script += "\n        query =ds['Query']['CommandText']                                                                                                                                                                                     ";
                    script += "\n        ds_type = 'StoredProcedure'                                                                                                                                                                                           ";
                    script += "\n        if re.search(r'EXEC\\s+\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?',query,re.IGNORECASE):                                                                                                                                    ";
                    script += "\n            db_name, schema,table_name = re.findall(r'EXEC\\s+\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?',query,re.IGNORECASE)[-1]                                                                                                  ";
                    script += "\n        elif re.search(r'EXEC\\s+\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?',query,re.IGNORECASE):                                                                                                                                               ";
                    script += "\n            table_name = re.findall(r'EXEC\\s+\\[?(\\w+)\\]?\\.\\[?(\\w+)\\]?',query,re.IGNORECASE)[-1][-1]                                                                                                                           ";
                    script += "\n    else:                                                                                                                                                                                                                     ";
                    script += "\n        ds_type = 'Dataset'                                                                                                                                                                                                   ";
                    script += "\n        query= ds['Query']['CommandText']                                                                                                                                                                                     ";
                    script += "\n        query = re.sub(r'\\n',' ',query)                                                                                                                                                                                       ";
                    script += "\n        query = re.sub(r'\\s+',' ',query)                                                                                                                                                                                      ";
                    script += "\n        db_name, schema, table_name = get_table(query)                                                                                                                                                                        ";
                    script += "\n    return ds_type, db_name, schema, table_name, query                                                                                                                                                                        ";
                    script += "\nreport = []                                                                                                                                                                                                                   ";
                    script += "\nfor _, each_row in df_data.iterrows():                                                                                                                                                                                        ";
                    script += "\n    query = None                                                                                                                                                                                                              ";
                    script += "\n    data = xmltodict.parse(each_row['ReportXML'], attr_prefix = '', dict_constructor=dict)                                                                                                                                    ";
                    script += "\n    if type(data['Report']['DataSources']['DataSource']) == list:                                                                                                                                                             ";
                    script += "\n        ds_list = [item ['Name'] for item in data['Report']['DataSources']['DataSource']]                                                                                                                                     ";
                    script += "\n    else:                                                                                                                                                                                                                     ";
                    script += "\n        ds_list = [data['Report']['DataSources']['DataSource']['Name']]                                                                                                                                                       ";
                    script += "\n                                                                                                                                                                                                                              ";
                    script += "\n    if type(data['Report']['DataSets']['DataSet']) == dict:                                                                                                                                                                   ";
                    script += "\n        ds_type, db_name, schema, table_name, query = get_data(data['Report']['DataSets']['DataSet'])                                                                                                                         ";
                    script += "\n        for each in data['Report']['DataSets']['DataSet']['Fields']['Field']:                                                                                                                                                 ";
                    script += "\n            report.append([each_row['Name'],data['Report']['rd:ReportID'], ds_list, data['Report']['DataSets']['DataSet']['Name'],                                                                                            ";
                    script += "\n                               ds_type,db_name,schema,table_name,each['Name'], each['DataField'],query])                                                                                                                      ";
                    script += "\n    else:                                                                                                                                                                                                                     ";
                    script += "\n        for each in data['Report']['DataSets']['DataSet']:                                                                                                                                                                    ";
                    script += "\n            ds_type, db_name, schema, table_name, query = get_data(each)                                                                                                                                                      ";
                    script += "\n            if type(each['Fields']['Field']) == list:                                                                                                                                                                         ";
                    script += "\n                for each_field in each['Fields']['Field']:                                                                                                                                                                    ";
                    script += "\n                    if 'DataField' in each_field:                                                                                                                                                                             ";
                    script += "\n                        report.append([each_row['Name'],data['Report']['rd:ReportID'], ds_list,                                                                                                                               ";
                    script += "\n                                       each['Name'],ds_type,db_name,schema,table_name,                                                                                                                                        ";
                    script += "\n                                       each_field['Name'], each_field['DataField'],query])                                                                                                                                    ";
                    script += "\n                    else:                                                                                                                                                                                                     ";
                    script += "\n                        report.append([each_row['Name'],data['Report']['rd:ReportID'], ds_list,                                                                                                                               ";
                    script += "\n                                       each['Name'],ds_type,db_name,schema,table_name,                                                                                                                                        ";
                    script += "\n                                       each_field['Name'], np.nan, query])                                                                                                                                                    ";
                    script += "\n            else:                                                                                                                                                                                                             ";
                    script += "\n                report.append([each_row['Name'],data['Report']['rd:ReportID'], ds_list,                                                                                                                                       ";
                    script += "\n                                each['Name'],ds_type,db_name,schema,table_name,                                                                                                                                               ";
                    script += "\n                                each['Fields']['Field']['Name'], each['Fields']['Field']['DataField'],query])                                                                                                                 ";
                    script += "\ndf = pd.DataFrame(report, columns = ['Report Name','Report ID', 'DataSource', 'DataSet','DataSet Type','Database Name','Schema','Table Name', 'Field Name', 'DataField','Query'])                                             ";
                    script += "\ndf['DataSource'] = df['DataSource'].astype(str)                                                                                                                                                                               ";
                    script += "\ndf.to_sql('flattened_ssrs_report', schema='dbo',if_exists = 'replace', con = engine)                                                                                                                                          ";
                    script += "\nreport_names = df['Report Name'].unique().tolist()                                                                                                                                                                            ";
                    script += "\nresult=[]                                                                                                                                                                                                                     ";
                    script += "\nfor each_rpt in report_names:                                                                                                                                                                                                 ";
                    script += "\n    inner_dataset = []                                                                                                                                                                                                        ";
                    script += "\n    inner_database = []                                                                                                                                                                                                       ";
                    script += "\n    inner_datafield = []                                                                                                                                                                                                      ";
                    script += "\n    inner_query =[]                                                                                                                                                                                                           ";
                    script += "\n    temp_data = df[(df['Report Name'] == each_rpt)]                                                                                                                                                                           ";
                    script += "\n    for _, each_row in temp_data.iterrows():                                                                                                                                                                                  ";
                    script += "\n        inner_dataset.append(each_row['DataSet'])                                                                                                                                                                             ";
                    script += "\n        inner_database.append(each_row['Database Name'])                                                                                                                                                                      ";
                    script += "\n        inner_datafield.append((each_row['Table Name']+'.'+each_row['DataField']) if each_row['Table Name'] else each_row['DataField'] )                                                                                      ";
                    script += "\n        inner_query.append(each_row['Query'])                                                                                                                                                                                 ";
                    script += "\n    inner_datafield = [i for i in inner_datafield if i is not None]                                                                                                                                                           ";
                    script += "\n    inner_dataset = [i for i in inner_dataset if i is not None]                                                                                                                                                               ";
                    script += "\n    inner_database = [i for i in inner_database if i is not None]                                                                                                                                                             ";
                    script += "\n    inner_query = [i for i in inner_query if i is not None]                                                                                                                                                                   ";
                    script += "\n    result.append([each_rpt, inner_datafield,inner_dataset,inner_database,inner_query])                                                                                                                                       ";
                    script += "\npercentage_table = []                                                                                                                                                                                                         ";
                    script += "\nfor i in range(0, len(result)):                                                                                                                                                                                               ";
                    script += "\n    for j in range(i+1, len(result)):                                                                                                                                                                                         ";
                    script += "\n        datafield_per = 0                                                                                                                                                                                                     ";
                    script += "\n        dataset_per = 0                                                                                                                                                                                                       ";
                    script += "\n        database_per = 0                                                                                                                                                                                                      ";
                    script += "\n        query_per = 0                                                                                                                                                                                                         ";
                    script += "\n        if result[i][1] == [] and result[j][1] == []:                                                                                                                                                                         ";
                    script += "\n            datafield_per = np.nan                                                                                                                                                                                            ";
                    script += "\n        elif result[i][1] != [] or result[j][1] != []:                                                                                                                                                                        ";
                    script += "\n            datafield_per = len(set(result[i][1]).intersection(set(result[j][1]))) / float(len(set(result[i][1] + result[j][1]))) * 100                                                                                       ";
                    script += "\n        if result[i][2] == [] and result[j][2] == []:                                                                                                                                                                         ";
                    script += "\n            dataset_per = np.nan                                                                                                                                                                                              ";
                    script += "\n        elif result[i][2] != [] or result[j][2] != []:                                                                                                                                                                        ";
                    script += "\n            dataset_per = len(set(result[i][2]).intersection(set(result[j][2]))) / float(len(set(result[i][2] + result[j][2]))) * 100                                                                                         ";
                    script += "\n        if result[i][3] == [] and result[j][3] == []:                                                                                                                                                                         ";
                    script += "\n             database_per = np.nan                                                                                                                                                                                            ";
                    script += "\n        elif result[i][3] != [] or result[j][3] != []:                                                                                                                                                                        ";
                    script += "\n             database_per = len(set(result[i][3]).intersection(set(result[j][3]))) / float(len(set(result[i][3] + result[j][3]))) * 100                                                                                       ";
                    script += "\n        if result[i][4] == [] and result[j][4] == []:                                                                                                                                                                         ";
                    script += "\n            query_per = np.nan                                                                                                                                                                                                ";
                    script += "\n        elif result[i][4] != [] or result[j][4] != []:                                                                                                                                                                        ";
                    script += "\n            query_per = len(set(result[i][4]).intersection(set(result[j][4]))) / float(len(set(result[i][4] + result[j][4]))) * 100                                                                                           ";
                    script += "\n        percentage_table.append([result[i][0], result[j][0], datafield_per,dataset_per,database_per,query_per])                                                                                                               ";
                    script += "\npercentage_table_df = pd.DataFrame(percentage_table, columns = ['Report A', 'Report B', 'DataField','DataSet','Database','Query'])                                                                                            ";
                    script += "\npercentage_table_df.to_sql('ssrs_report_match_percentage', schema='dbo',if_exists = 'replace', con = engine)                                                                                                                  ";
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          

                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                using (StreamWriter writer = File.CreateText(path))
                {
                    writer.WriteLine(script);
                }
                try
                {


                    createsqlDatabase();



                }
                catch (Exception ex)
                {

                }

            }

            MessageBox.Show("Data loaded to the " + Source123.Text.ToString());

            }

        }
        private async void run_cmd()
        {

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
                        sw.WriteLine(TextPython.Text + @"\Scripts\activate.bat");
                        sw.WriteLine("python SSRS_Process_Python.py");
                    }
                }
                //string output = process.StandardOutput.ReadToEnd();

                process.WaitForExit();
                Animation.Visibility = Visibility.Collapsed;
                LabelUser.Margin = new Thickness(20, 100, 0, 0);
                Stack1.Margin = new Thickness(20, 100, 0, 0);
                LabelUser2.Visibility = Visibility.Collapsed;
                DBLabelUser2.Visibility = Visibility.Visible;
                PasswordChek1.Visibility = Visibility.Visible;

                LabelUser.Visibility = Visibility.Collapsed;
                TemplatePath.Visibility = Visibility.Collapsed;
                BorderTemplatePAth.Visibility = Visibility.Collapsed;
                TextCSV.Visibility = Visibility.Collapsed;
                Template_Browse.Visibility = Visibility.Collapsed;

                LabelUser1.Visibility = Visibility.Visible;
                LabelHostName.Visibility = Visibility.Visible;
                BorderHosName.Visibility = Visibility.Visible;
                HostName.Visibility = Visibility.Visible;

                LabelDataBaseName.Visibility = Visibility.Visible;
                BorderDataBaseName.Visibility = Visibility.Visible;
                DataBaseName.Visibility = Visibility.Visible;

                LabelUserName.Visibility = Visibility.Visible;
                BorderUserName.Visibility = Visibility.Visible;
                UserName.Visibility = Visibility.Visible;

                LabelPassword.Visibility = Visibility.Visible;
                BorderPassword1.Visibility = Visibility.Visible;
                Password1.Visibility = Visibility.Visible;

               
                PasswordChek.IsChecked = false;
                WindRad.IsChecked = true;

                InsertXML.Visibility = Visibility.Visible;
                GeneratePBI.Visibility = Visibility.Visible;
                GenerateDoc.Visibility = Visibility.Visible;

                InsertXML.Margin = new Thickness(-760, -100, 171.333, 0);
                GeneratePBI.Margin = new Thickness(-20, -100, 463.667, 0);
                GenerateDoc.Margin = new Thickness(360, -100, 340, 0);

                LabelSource.Visibility = Visibility.Visible;
                BorderSource.Visibility = Visibility.Visible;
                WindRad.Visibility = Visibility.Visible;
                AuthRad.Visibility = Visibility.Visible;
                LabelDatabaseServer.Visibility = Visibility.Visible;
                BorderServer.Visibility = Visibility.Visible;

                DBConnection.Visibility = Visibility.Visible;
                CSVpath.Visibility = Visibility.Visible;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }


            
        }
        public void createsqlDatabase()
        {
            string connectionString = @"Data Source = " + Source123.Text.Replace("\\\\", "\\") + "; Integrated Security=true";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string strconnection = "Data Source = " + Source123.Text.ToString() + "; Integrated Security=true";

            string table = "IF NOT EXISTS(SELECT name FROM master.dbo.sysdatabases WHERE Name='SSRS Metadata') CREATE DATABASE[SSRS Metadata]";
            InsertQuery1(table, strconnection);
            run_cmd();


        }
        public void InsertQuery1(string qry, string connection)
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

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            TextCSV.IsEnabled = true;
            Source123.IsEnabled = true;
            TextPython.IsEnabled = true;
            TextCSV.Text = "";
            Source123.Text = "";
            TextPython.Text = "";

            //    GenerateMetadata.IsChecked = true;
            //   ProcessStart.Visibility = Visibility.Collapsed;
            InsertXML.Visibility = Visibility.Visible;
            GeneratePBI.Visibility = Visibility.Visible;
            GenerateDoc.Visibility = Visibility.Visible;
            ProcessImage.Visibility = Visibility.Collapsed;
            OutputImage.Visibility = Visibility.Collapsed;
            DocImage.Visibility = Visibility.Collapsed;

        }

        private void username_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

    }
}




