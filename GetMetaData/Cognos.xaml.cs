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
using System;
using System.IO;
using System.Text;
using System;
using System.Xml;
using Microsoft.Scripting.Hosting;
using IronPython.Hosting;
using System.Xml.Linq;

namespace GetMetaData
{
    /// <summary>
    /// Interaction logic for Cognos.xaml
    /// </summary>
    public partial class Cognos : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;
        DataSet ds = new DataSet();
        DataSet dsLocal = new DataSet();
        string countrows = "";
        string contentStoreDB = "";

        private static string PythonPath1;
        public Cognos()
        {
            InitializeComponent();
            Labelusername.Visibility = Visibility.Collapsed;
            Borderusername.Visibility = Visibility.Collapsed;
            Labelpasswd.Visibility = Visibility.Collapsed;
            Borderpasswd.Visibility = Visibility.Collapsed;
            PasswordChek.Visibility = Visibility.Collapsed;
            button1.Visibility = Visibility.Collapsed;
            PDF.Visibility = Visibility.Collapsed;
            ReqButton.Visibility = Visibility.Collapsed;
            //ProcessStart.Visibility = Visibility.Collapsed;
            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            MyNotifyIcon.Icon = new System.Drawing.Icon(
                            @"Final.ico");
            MyNotifyIcon.MouseDoubleClick +=
                new System.Windows.Forms.MouseEventHandler(MyNotifyIcon_MouseDoubleClick);
            //LabelPythonPath.Visibility = Visibility.Collapsed;
            //BorderPythonPath.Visibility = Visibility.Collapsed;
            //Browse_Copy.Visibility = Visibility.Collapsed;
            WindowMainName.Height = 650;
            Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
            ComboBoxZone.Items.Add("SQL Source");
            WindRad.IsChecked = true;
            WindRad1.IsChecked = true;
            ProcessStart.Visibility = Visibility.Collapsed;
            Output.Visibility = Visibility.Collapsed;
            GenerateMetadata.Visibility = Visibility.Collapsed;
            // GenerateMetadata.IsChecked = true;
            ImageToolTip.Text = "Fill the mandatory fields in the below sequence : ";
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("1. Choose Connector -  Choose the relevant connector on which the Cognos Reports are configured ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("2. Server Details - Server where the Content Store Database is configured (Windows or SQL Authentication) ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("3. Get Relevant Database - Clicking on this will list the databases where the XML object is available in the environment provided");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("4. Target SQL Server - Server where the generated metadata will be inserted (Windows or SQL Authentication) ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("5. Python Path - Directory where the python is installed in the system ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("6. Generate Metadata - Start the process for Metadata generation ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("7. Generate Output/Requirement Doc - generate output or requirement document based on the metadata inserted in Step 6 ");
            ProcessImage.Visibility = Visibility.Collapsed;
            OutputImage.Visibility = Visibility.Collapsed;

            
            
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

        private void SignOutButton_Click(object sender, RoutedEventArgs e)
        {

        }


        private void WindAuth_Checked(object sender, RoutedEventArgs e)
        {
            Labelusername.Visibility = Visibility.Collapsed;
            Borderusername.Visibility = Visibility.Collapsed;
            Labelpasswd.Visibility = Visibility.Collapsed;
            Borderpasswd.Visibility = Visibility.Collapsed;
            PasswordChek.Visibility = Visibility.Collapsed;
            if (WindRad.IsChecked == true && WindRad1.IsChecked == true)
            {
                WindowMainName.Height = 650;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;

            }
            else if (WindRad.IsChecked == true && WindRad1.IsChecked == false)
            {
                WindowMainName.Height = 703.334;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
            else if (WindRad1.IsChecked == true && WindRad.IsChecked == false)
            {
                WindowMainName.Height = 703.334;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
        }

        private void SQL_Checked(object sender, RoutedEventArgs e)
        {

            Labelusername.Visibility = Visibility.Visible;
            Borderusername.Visibility = Visibility.Visible;
            Labelpasswd.Visibility = Visibility.Visible;
            Borderpasswd.Visibility = Visibility.Visible;
            PasswordChek.Visibility = Visibility.Visible;
            if (AuthRad.IsChecked == true && AuthRad1.IsChecked == true)
            {
                WindowMainName.Height = 926;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
                
            }
            else if(AuthRad1.IsChecked==true && AuthRad.IsChecked==false)
            {
                WindowMainName.Height = 703.334;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
            else if (AuthRad.IsChecked == true && AuthRad1.IsChecked == false)
            {
                WindowMainName.Height = 703.334;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Generating Report. Please Wait ....");
            // string file = @"Metadata Output.pbix";
            string fileName = "BI4BI - Cognos.pbix";
            string path = Path.Combine(Environment.CurrentDirectory, @"Report\", fileName);
            Process.Start(path);
        }
        private async void run_cmd()
        {
            //MessageBox.Show("Report Generation in process. Average Wait Time less than 1 minute.");

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
                        sw.WriteLine("py Cognos_Python.py");
                    }
                }
                //string output = process.StandardOutput.ReadToEnd();
                
                process.WaitForExit();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            /*
            try
            {
                string workingDirectory = Directory.GetCurrentDirectory() + @"\PythonFile";
                string fileName = workingDirectory+@"\Cognos_Python.py";

                Process p = new Process();
                p.StartInfo = new ProcessStartInfo(PythonPathText.Text.ToString()+ @"\python.exe", fileName)
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                p.Start();

                string output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();

                Console.WriteLine(output);

                Console.ReadLine();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }*/




        }
        public async void createsqltableUsage()
        {


            try
            {
                string connectionString = @"Data Source = " + Source123.Text.Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=Cognos Metadata";
                SqlConnection sqlconnection = new SqlConnection(connectionString);
                sqlconnection.Open();
                string strconnection = "Data Source = " + Source123.Text.ToString() + "; Integrated Security=true; Initial Catalog=Cognos Metadata";
                string table = "";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='CognosDataItems') ";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='CognosFiltersExpression')   ";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='CognosReportVariables') ";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='CognosReportModificationTime') ";
                InsertQuery1(table, strconnection);

            }
            catch
            {
               // MessageBox.Show("Please check the SQL server Instance and try again");
            }


        }
        



        private void ReqButton_Click(object sender, RoutedEventArgs e)
        {
            int result = 0;

            string connectionstring = "Data Source=" + Server.Text.ToString() + "; Integrated Security=true; Initial Catalog=Cognos Metadata"; ; //your connectionstring    

            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("select COUNT(*) from dbo.CognosDataItems", conn);
                result = (int)cmd.ExecuteScalar();
                conn.Close();
            }
            if (Server.Text.ToString().Equals("") || result == 0)
            {
                MessageBox.Show("Either the Metadata is not extracted or the SQL Server details is blank");
            }
            else
            {
                Document_Generator_Cognos objWelcome = new Document_Generator_Cognos();
                objWelcome.SQLTB.Text = Server.Text;
                objWelcome.Show(); //Sending value from one form to another form.
                Close();
            }
        }

        private void CallDatabaseList(object sender, RoutedEventArgs e)
        {

            
            if (WindRad.IsChecked == true)
            {
                if (String.IsNullOrEmpty(ComboBoxZone.Text.ToString()) || String.IsNullOrEmpty(Source123.Text.ToString()))
                {
                    MessageBox.Show("Choose the Connector and Enter the server details to proceed.");
                }
                else
                {
                    Animation.Visibility = Visibility.Visible;
                    LabelReport.Visibility = Visibility.Collapsed;
                    Stack1.Visibility = Visibility.Collapsed;
                    LabelDatabaseServer.Visibility = Visibility.Collapsed;
                    BorderServer.Visibility = Visibility.Collapsed;
                    InsertintoLocal.Visibility = Visibility.Collapsed;
                    MessageBox.Show("Locating Object. Average wait time less than 1 minute.");
                    /* ds.Clear();
                     ds = GetConnectionStringForCombo(Source123.Text.ToString());
                     ComboBoxDB.ItemsSource = ds.Tables[0].DefaultView;
                     ComboBoxDB.DisplayMemberPath = ds.Tables[0].Columns["Database Name"].ToString();
                     ComboBoxDB.SelectedValuePath = ds.Tables[0].Columns["Database Name"].ToString();
                    */
                    SqlConnection SQLConnection = new SqlConnection();
                    SQLConnection.ConnectionString = "Data Source =" + Source123.Text.ToString() + "; Initial Catalog =master; " + "Integrated Security=true;";

                    string QueryDB =  "SELECT TOP 1 name [Database Name] FROM sys.databases WHERE CASE WHEN state_desc = 'ONLINE' THEN OBJECT_ID(QUOTENAME(name) +'.dbo.CMOBJPROPS7', 'U') END IS NOT NULL";
                    //Execute Queries and save results into variables
                    SqlCommand CmdCnt = SQLConnection.CreateCommand();
                    CmdCnt.CommandText = QueryDB;
                    SQLConnection.Open();
                    contentStoreDB =(string) CmdCnt.ExecuteScalar();
                    SQLConnection.Close();
                    MessageBox.Show("Object located successfully. Proceed on the next steps");
                }
            }
            else
            {
                if (PasswordChek.IsChecked == true)
                {
                    if (String.IsNullOrEmpty(ComboBoxZone.Text.ToString()) || String.IsNullOrEmpty(Source123.Text.ToString())
                        || String.IsNullOrEmpty(username.Text.ToString()) || String.IsNullOrEmpty(PasswordShow.Text.ToString()))
                    {
                        MessageBox.Show("Choose the Connector and Enter the server details to proceed.");
                    }
                    else
                    {
                        Animation.Visibility = Visibility.Visible;
                        LabelReport.Visibility = Visibility.Collapsed;
                        Stack1.Visibility = Visibility.Collapsed;
                        LabelDatabaseServer.Visibility = Visibility.Collapsed;
                        BorderServer.Visibility = Visibility.Collapsed;
                        MessageBox.Show("Locating Database. Average wait time less than 1 minute.");
                        /* ds.Clear();
                         ds = GetConnectionStringForComboSQLAuth(Source123.Text.ToString(), username.Text.ToString(), PasswordShow.Text.ToString());
                         ComboBoxDB.ItemsSource = ds.Tables[0].DefaultView;
                         ComboBoxDB.DisplayMemberPath = ds.Tables[0].Columns["Database Name"].ToString();
                         ComboBoxDB.SelectedValuePath = ds.Tables[0].Columns["Database Name"].ToString();
                        */
                        SqlConnection SQLConnection = new SqlConnection();
                        SQLConnection.ConnectionString = @"Data Source = " + Source123.Text.ToString() + "; ;User ID=" + username.Text.ToString() + ";Password=" + PasswordShow.Text.ToString() + ";";

                        string QueryDB = "SELECT TOP 1 name [Database Name] FROM sys.databases WHERE CASE WHEN state_desc = 'ONLINE' THEN OBJECT_ID(QUOTENAME(name) +'.dbo.CMOBJPROPS7', 'U') END IS NOT NULL";
                        //Execute Queries and save results into variables
                        SqlCommand CmdCnt = SQLConnection.CreateCommand();
                        CmdCnt.CommandText = QueryDB;
                        SQLConnection.Open();
                        contentStoreDB = (string)CmdCnt.ExecuteScalar();
                        SQLConnection.Close();
                        MessageBox.Show("Object located successfully. Proceed on the next steps");
                    }
                }
                else
                {
                    if (String.IsNullOrEmpty(ComboBoxZone.Text.ToString()) || String.IsNullOrEmpty(Source123.Text.ToString())
                        || String.IsNullOrEmpty(username.Text.ToString()) || String.IsNullOrEmpty(Password.Password.ToString()))
                    {
                        MessageBox.Show("Choose the Connector and Enter the server details to proceed.");
                    }
                    else
                    {
                        Animation.Visibility = Visibility.Visible;
                        LabelReport.Visibility = Visibility.Collapsed;
                        Stack1.Visibility = Visibility.Collapsed;
                        LabelDatabaseServer.Visibility = Visibility.Collapsed;
                        BorderServer.Visibility = Visibility.Collapsed;
                        MessageBox.Show("Retrieving relevant Database. Average wait time less than 1 minute.");
                        SqlConnection SQLConnection = new SqlConnection();
                        SQLConnection.ConnectionString = @"Data Source = " + Source123.Text.ToString() + "; ;User ID=" + username.Text.ToString() + ";Password=" + PasswordShow.Text.ToString() + ";";

                        string QueryDB = "SELECT TOP 1 name [Database Name] FROM sys.databases WHERE CASE WHEN state_desc = 'ONLINE' THEN OBJECT_ID(QUOTENAME(name) +'.dbo.CMOBJPROPS7', 'U') END IS NOT NULL";
                        //Execute Queries and save results into variables
                        SqlCommand CmdCnt = SQLConnection.CreateCommand();
                        CmdCnt.CommandText = QueryDB;
                        SQLConnection.Open();
                        contentStoreDB = (string)CmdCnt.ExecuteScalar();
                        SQLConnection.Close();
                        MessageBox.Show("Object located successfully. Proceed on the next steps");
                    }
                }

            }

            
            Animation.Visibility = Visibility.Collapsed;
            LabelReport.Visibility = Visibility.Visible;
            Stack1.Visibility = Visibility.Visible;
            LabelDatabaseServer.Visibility = Visibility.Visible;
            BorderServer.Visibility = Visibility.Visible;
            InsertintoLocal.Visibility = Visibility.Visible;
            
        }

        public DataSet GetConnectionStringForCombo(string server)
        {
            string connectionString = @"Data Source = " + server.Replace("\\\\", "\\") + "; Integrated Security=true;";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string query = "SELECT name [Database Name] FROM sys.databases WHERE CASE WHEN state_desc = 'ONLINE' THEN OBJECT_ID(QUOTENAME(name) + '.dbo.CMOBJPROPS7', 'U') END IS NOT NULL";

            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();

            // create data adapter
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            // this will query your database and return the result to your datatable
            da.Fill(ds);
            conn.Close();
            da.Dispose();
            return ds;

        }
        public DataSet GetConnectionStringForComboSQLAuth(string server, string user, string passwd)
        {
            string connectionString = @"Data Source = " + server.Replace("\\\\", "\\") + "; ;User ID=" + user + ";Password=" + passwd + ";";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string query = "SELECT name [Database Name] FROM sys.databases WHERE CASE WHEN state_desc = 'ONLINE' THEN OBJECT_ID(QUOTENAME(name) + '.dbo.CMOBJPROPS7', 'U') END IS NOT NULL";

            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();

            // create data adapter
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            // this will query your database and return the result to your datatable
            da.Fill(ds);
            conn.Close();
            da.Dispose();
            return ds;

        }
        public DataSet GetDataforLocalHostUSerNamePasswd(string server, string user, string passwd, string db)
        {
            string connectionString = @"Data Source = " + server.Replace("\\\\", "\\") + ";User ID=" + user + ";Password=" + passwd + ";Initial Catalog=" + db.ToString();
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string query = "SELECT SPEC FROM [" + db.ToString() + "].dbo.CMOBJPROPS7";

            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();

            // create data adapter
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            // this will query your database and return the result to your datatable
            da.Fill(dsLocal);
            conn.Close();
            da.Dispose();
            return dsLocal;

        }
        public DataSet GetDataforLocalHostWindowsAuth(string server,string db)
        {
            
            string connectionString = @"Data Source = " + server.Replace("\\\\", "\\") + ";Integrated Security=true ; Initial Catalog=" + db.ToString();
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
           
            string query = "SELECT SPEC FROM [" + db.ToString() + "].dbo.CMOBJPROPS7";

            SqlConnection conn = new SqlConnection(connectionString);
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();

            // create data adapter
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            // this will query your database and return the result to your datatable
            da.Fill(dsLocal);
            conn.Close();
            da.Dispose();
            return dsLocal;

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

        private void InsertintoLocal_Click(object sender, RoutedEventArgs e)
        {
            if (WindRad.IsChecked == true)
            {
                if ( String.IsNullOrEmpty(Server.Text.ToString()))
                {
                    MessageBox.Show("Choose the Connector and Enter the server details to proceed.");
                }
                else
                {
                    try
                    {
                        SqlConnection SQLConnection = new SqlConnection();
                        SQLConnection.ConnectionString = "Data Source =" + Source123.Text.ToString() + "; Initial Catalog =Cognos Metadata; " + "Integrated Security=true;";
                        
                        Animation.Visibility = Visibility.Visible;
                        LabelReport.Visibility = Visibility.Collapsed;
                        Stack1.Visibility = Visibility.Collapsed;
                        LabelDatabaseServer.Visibility = Visibility.Collapsed;
                        BorderServer.Visibility = Visibility.Collapsed;
                        button1.Visibility = Visibility.Collapsed;
                        PDF.Visibility = Visibility.Collapsed;
                        ReqButton.Visibility = Visibility.Collapsed;
                        InsertintoLocal.Visibility = Visibility.Collapsed;
                        LabelPythonPath.Visibility = Visibility.Collapsed;
                        BorderPythonPath.Visibility = Visibility.Collapsed;
                        Browse_Copy.Visibility = Visibility.Collapsed;
                        ProcessStart.Visibility = Visibility.Collapsed;
                        GenerateMetadata.Visibility = Visibility.Collapsed;
                        Output.Visibility = Visibility.Collapsed;
                        ProcessImage.Visibility = Visibility.Collapsed;
                        OutputImage.Visibility = Visibility.Collapsed;
                        MessageBox.Show("Inserting into Target SQL Server.Wait time depends on the Number of XML's being inserted.");
                        dsLocal.Clear();
                        dsLocal = GetDataforLocalHostWindowsAuth(Source123.Text.ToString(), contentStoreDB.ToString());
                        dsLocal.Tables[0].Columns["SPEC"].ColumnName = "XMLData";
                        createsqlDatabase();
                        createsqltable(dsLocal, "XMLData_Python");
                        string QueryXML = "select count(*) from dbo.XMLData_Python";
                        //Execute Queries and save results into variables
                        SqlCommand CmdCntXML = SQLConnection.CreateCommand();
                        CmdCntXML.CommandText = QueryXML;
                        SQLConnection.Open();
                        int XMLCnt = (Int32)CmdCntXML.ExecuteScalar();
                        SQLConnection.Close();
                            if (XMLCnt > 0)
                        {
                            
                            Animation.Visibility = Visibility.Collapsed;
                            LabelReport.Visibility = Visibility.Visible;
                            Stack1.Visibility = Visibility.Visible;
                            LabelDatabaseServer.Visibility = Visibility.Visible;
                            BorderServer.Visibility = Visibility.Visible;
                            button1.Visibility = Visibility.Collapsed;
                            PDF.Visibility = Visibility.Collapsed;
                            ReqButton.Visibility = Visibility.Collapsed;
                            InsertintoLocal.Visibility = Visibility.Visible;
                            LabelPythonPath.Visibility = Visibility.Visible;
                            BorderPythonPath.Visibility = Visibility.Visible;
                            Browse_Copy.Visibility = Visibility.Visible;
                            GenerateMetadata.IsChecked = true;
                            ProcessStart.Visibility = Visibility.Visible;
                            GenerateMetadata.Visibility = Visibility.Visible;
                            Output.Visibility = Visibility.Visible;
                            ProcessImage.Visibility = Visibility.Visible;
                            OutputImage.Visibility = Visibility.Visible;
                            Output.IsEnabled = false;
                            MessageBox.Show("Successfully moved the XML's from source to target server. Hover over Tips icon for more info.");
                            //MetadataToolTip.Text = "The number of XML's available in the server " + Source123.Text.ToString() + " and Database = Cognos Metadata\r\n";
                            MetadataToolTip.Text = "Number of Distinct XML's available in the server " + Source123.Text.ToString() + " = " + XMLCnt.ToString() + "\r\n";
                            MetadataToolTip.AppendText("Make sure the database has enough capacity to process the XML's ");
                            MetadataToolTip.AppendText(Environment.NewLine);
                            MetadataToolTip.AppendText(Environment.NewLine);
                            MetadataToolTip.AppendText("Tip : In case you are using your localhost server , the maximum limit of Daatbase is 10 GB");
                        }
                        else
                        {
                            Animation.Visibility = Visibility.Collapsed;
                            LabelReport.Visibility = Visibility.Visible;
                            Stack1.Visibility = Visibility.Visible;
                            LabelDatabaseServer.Visibility = Visibility.Visible;
                            BorderServer.Visibility = Visibility.Visible;
                            button1.Visibility = Visibility.Collapsed;
                            PDF.Visibility = Visibility.Collapsed;
                            ReqButton.Visibility = Visibility.Collapsed;
                            InsertintoLocal.Visibility = Visibility.Visible;
                            LabelPythonPath.Visibility = Visibility.Visible;
                            BorderPythonPath.Visibility = Visibility.Visible;
                            Browse_Copy.Visibility = Visibility.Visible;
                            ProcessStart.Visibility = Visibility.Collapsed;
                            GenerateMetadata.Visibility = Visibility.Collapsed;
                            Output.Visibility = Visibility.Collapsed;
                            ProcessImage.Visibility = Visibility.Collapsed;
                            OutputImage.Visibility = Visibility.Collapsed;
                            

                            MessageBox.Show("Some issue in inserting the XML Data.Please check the source or contact the adminstrator in case of any support needed");
                        }
                    }
                    catch (Exception ex)
                    {
                        Animation.Visibility = Visibility.Collapsed;
                        LabelReport.Visibility = Visibility.Visible;
                        Stack1.Visibility = Visibility.Visible;
                        LabelDatabaseServer.Visibility = Visibility.Visible;
                        BorderServer.Visibility = Visibility.Visible;
                        button1.Visibility = Visibility.Collapsed;
                        PDF.Visibility = Visibility.Collapsed;
                        ReqButton.Visibility = Visibility.Collapsed;
                        InsertintoLocal.Visibility = Visibility.Visible;
                        LabelPythonPath.Visibility = Visibility.Visible;
                        BorderPythonPath.Visibility = Visibility.Visible;
                        Browse_Copy.Visibility = Visibility.Visible;
                        ProcessStart.Visibility = Visibility.Collapsed;
                        GenerateMetadata.Visibility = Visibility.Collapsed;
                        Output.Visibility = Visibility.Collapsed;
                        ProcessImage.Visibility = Visibility.Collapsed;
                        OutputImage.Visibility = Visibility.Collapsed;

                        MessageBox.Show(ex.Message.ToString());
                        //MessageBox.Show("Some issue in inserting the XML Data.Please check the source or contact the adminstrator in case of any support needed");
                    }
            }
            }

            else
            {
                string password = "";
                if (PasswordChek.IsChecked == true)
                {
                    password = PasswordShow.Text.ToString();
                }
                else
                {
                    password = Password.Password;
                }
                if ( String.IsNullOrEmpty(Server.Text.ToString())
                         || String.IsNullOrEmpty(username.Text.ToString()) || String.IsNullOrEmpty(PasswordShow.Text.ToString()))
                {
                    MessageBox.Show("Choose the Connector and Enter the server details to proceed.");
                }
                else
                {
                    Animation.Visibility = Visibility.Visible;
                    LabelReport.Visibility = Visibility.Collapsed;
                    Stack1.Visibility = Visibility.Collapsed;
                    LabelDatabaseServer.Visibility = Visibility.Collapsed;
                    BorderServer.Visibility = Visibility.Collapsed;
                    MessageBox.Show("Retrieving relevant Database. Average wait time less than 1 minute.");
                    dsLocal.Clear();
                    dsLocal = GetDataforLocalHostUSerNamePasswd(Source123.Text.ToString(), username.Text.ToString(), password.ToString(), contentStoreDB.ToString());
                    createsqlDatabase();
                    createsqltable(dsLocal, "CMOBJPROPS7");
                    Animation.Visibility = Visibility.Collapsed;
                    LabelReport.Visibility = Visibility.Visible;
                    Stack1.Visibility = Visibility.Visible;
                    LabelDatabaseServer.Visibility = Visibility.Visible;
                    BorderServer.Visibility = Visibility.Visible;
                    button1.Visibility = Visibility.Visible;
                    PDF.Visibility = Visibility.Visible;
                    ReqButton.Visibility = Visibility.Visible;
                    InsertintoLocal.Visibility = Visibility.Visible;
                    LabelPythonPath.Visibility = Visibility.Visible;
                    BorderPythonPath.Visibility = Visibility.Visible;
                    Browse_Copy.Visibility = Visibility.Visible;
                }
            }
            
            
        }
        public void createsqltable(DataSet dt, string tablename)
        {
            createsqlDatabase();
            string connectionString = @"Data Source = " + Server.Text.Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=Cognos Metadata";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string strconnection = "Data Source = " + Server.Text.ToString() + "; Integrated Security=true; Initial Catalog=Cognos Metadata";

            string table = "\n IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='XMLData_Python') BEGIN DROP TABLE XMLData_Python END";
            table += "\n CREATE TABLE [dbo].[XMLData_Python](";
            //table += "\n 	[index] [int] IDENTITY(1,1) NOT NULL,";
            table += "\n 	[XMLData] [xml] NULL,";
            table += "\n 	[index] [int] IDENTITY(1,1) NOT NULL";
            table += "\n ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]";


            //string table = "TRUNCATE TABLE [Cognos Metadata].[dbo].[Metadata]";
            InsertQuery1(table, strconnection.ToString());
            CopyData1(strconnection, dt, tablename);
        }
        public void createsqlDatabase()
        {
            string connectionString = @"Data Source = " + Source123.Text.Replace("\\\\", "\\") + "; Integrated Security=true";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string strconnection = "Data Source = " + Source123.Text.ToString() + "; Integrated Security=true";

            string table = "IF NOT EXISTS(SELECT name FROM master.dbo.sysdatabases WHERE Name='Cognos Metadata') CREATE DATABASE [Cognos Metadata]";
            InsertQuery1(table, strconnection);
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
        public static void CopyData1(string connStr, DataSet dt, string tablename)
        {
            using (SqlBulkCopy bulkCopy =
            new SqlBulkCopy(connStr, SqlBulkCopyOptions.TableLock))
            {
                bulkCopy.DestinationTableName = tablename;
                bulkCopy.WriteToServer(dt.Tables[0]);
            }
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.ShowDialog();
            PythonPathText.Text = dialog.SelectedPath;
            PythonPath1 = PythonPathText.Text;
        }

        private void WindRad1_Checked(object sender, RoutedEventArgs e)
        {
            Labelusername1.Visibility = Visibility.Collapsed;
            Borderusername1.Visibility = Visibility.Collapsed;
            Labelpasswd1.Visibility = Visibility.Collapsed;
            Borderpasswd1.Visibility = Visibility.Collapsed;
            PasswordChek1.Visibility = Visibility.Collapsed;

            if (WindRad.IsChecked == true && WindRad1.IsChecked == true)
            {
                WindowMainName.Height = 650;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;

            }
            else if (WindRad.IsChecked == true && WindRad1.IsChecked == false)
            {
                WindowMainName.Height = 703.334;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
            else if (WindRad1.IsChecked == true && WindRad.IsChecked == false)
            {
                WindowMainName.Height = 703.334;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
        }

        private void AuthRad1_Checked(object sender, RoutedEventArgs e)
        {
            Labelusername1.Visibility = Visibility.Visible;
            Borderusername1.Visibility = Visibility.Visible;
            Labelpasswd1.Visibility = Visibility.Visible;
            Borderpasswd1.Visibility = Visibility.Visible;
            PasswordChek1.Visibility = Visibility.Visible;

            if (AuthRad.IsChecked == true && AuthRad1.IsChecked == true)
            {
                WindowMainName.Height = 926;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;

            }
            else if (AuthRad1.IsChecked == true && AuthRad.IsChecked == false)
            {
                WindowMainName.Height = 703.334;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
            else if (AuthRad.IsChecked == true && AuthRad1.IsChecked == false)
            {
                WindowMainName.Height = 703.334;
                Scroller.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            }
        }

        private void PasswordChek1_Checked(object sender, RoutedEventArgs e)
        {
            PasswordShow1.Text = Password.Password;
            BorderPasswordShow1.Visibility = Visibility.Visible;
            Borderpasswd1.Visibility = Visibility.Collapsed;
        }

        private void PasswordChek1_Unchecked(object sender, RoutedEventArgs e)
        {
            Password1.Password = PasswordShow.Text;
            BorderPasswordShow1.Visibility = Visibility.Collapsed;
            Borderpasswd1.Visibility = Visibility.Visible;
        }

        private void GenerateMetadata_Checked(object sender, RoutedEventArgs e)
        {
            ProcessStart.Visibility = Visibility.Visible;
            button1.Visibility = Visibility.Collapsed;
            PDF.Visibility = Visibility.Collapsed;
            ReqButton.Visibility = Visibility.Collapsed;
        }

        private void Output_Checked(object sender, RoutedEventArgs e)
        {
            ProcessStart.Visibility = Visibility.Collapsed;
            button1.Visibility = Visibility.Visible;
            PDF.Visibility = Visibility.Visible;
            ReqButton.Visibility = Visibility.Visible;
        }

        private void ProcessStart_Click(object sender, RoutedEventArgs e)
        {
            string path = Directory.GetCurrentDirectory() + @"\PythonFile\Cognos_Python.py";
            //MessageBox.Show(path.ToString());
            string password = "";
            SqlConnection SQLConnection = new SqlConnection();
            SQLConnection.ConnectionString = "Data Source =" + Source123.Text.ToString() + "; Initial Catalog =Cognos Metadata; " + "Integrated Security=true;";

            if (PasswordChek.IsChecked == true)
            {
                password = PasswordShow.Text.ToString();
            }
            else
            {
                password = Password.Password;
            }
            //string path = System.IO.Path.Combine(@"C:\Users\UT481LN\source\repos\GetMetaData\GetMetaData\GetMetaData\PythonFile",fileName);
            // string path = System.IO.Path.Combine(Environment.CurrentDirectory, @"PythonFile\", fileName);
            if (String.IsNullOrEmpty(ComboBoxZone.Text.ToString()) || String.IsNullOrEmpty(Source123.Text.ToString()))
            {
                MessageBox.Show("Please enter the mandatory fields and try again");
            }
            else
            {

                Animation.Visibility = Visibility.Visible;
                LabelReport.Visibility = Visibility.Collapsed;
                Stack1.Visibility = Visibility.Collapsed;
                LabelDatabaseServer.Visibility = Visibility.Collapsed;
                BorderServer.Visibility = Visibility.Collapsed;
                button1.Visibility = Visibility.Collapsed;
                PDF.Visibility = Visibility.Collapsed;
                ReqButton.Visibility = Visibility.Collapsed;
                InsertintoLocal.Visibility = Visibility.Collapsed;
                LabelPythonPath.Visibility = Visibility.Collapsed;
                BorderPythonPath.Visibility = Visibility.Collapsed;
                Browse_Copy.Visibility = Visibility.Collapsed;

                string QueryXML = "select count(*) from dbo.XMLData_Python";
                //Execute Queries and save results into variables
                SqlCommand CmdCntXML = SQLConnection.CreateCommand();
                CmdCntXML.CommandText = QueryXML;
                SQLConnection.Open();
                int XMLCnt = (Int32)CmdCntXML.ExecuteScalar();
                SQLConnection.Close();

                MessageBox.Show("Metadata Generation Proess started. Wait time depends on the number of XMLs being processed.\r\n" +
                "You can go to the desktop by pressing Windows+D and the process will run in the backgrounds.\r\n"
                + " XML's Processing in the current process = " + XMLCnt.ToString() + "\r\n"
                + " Tip : For 25,000 XML's the wait time is 180 minutes");

                string xmlpath = Directory.GetCurrentDirectory() + @"\Sample XML\Sample.xml";
                //Decalre a new XMLDocument object
                XmlDocument doc = new XmlDocument();

                //xml declaration is recommended, but not mandatory
                XmlDeclaration xmlDeclaration = doc.CreateXmlDeclaration("1.0", "UTF-8", null);

                //create the root element
                XmlElement root = doc.DocumentElement;
                doc.InsertBefore(xmlDeclaration, root);

                //string.Empty makes cleaner code
                XmlElement element1 = doc.CreateElement(string.Empty, "Mainbody", string.Empty);
                doc.AppendChild(element1);

                XmlElement element2 = doc.CreateElement(string.Empty, "level1", string.Empty);


                XmlElement element3 = doc.CreateElement(string.Empty, "level2", string.Empty);

                XmlText text1 = doc.CreateTextNode("Demo Text");

                element1.AppendChild(element2);
                element2.AppendChild(element3);
                element3.AppendChild(text1);


                XmlElement element4 = doc.CreateElement(string.Empty, "level2", string.Empty);
                XmlText text2 = doc.CreateTextNode("other text");
                element4.AppendChild(text2);
                element2.AppendChild(element4);

                doc.Save(xmlpath);


                // XmlTextWriter writer1 = new XmlTextWriter(xmlpath, null);


                string script = "\nimport xml.etree.cElementTree as et";
                script += "\nimport pandas as pd";
                script += "\nimport re";
                script += "\nfrom pandas import DataFrame";
                script += "\nfrom sqlalchemy import create_engine";
                script += "\nimport urllib";
                script += "\nfrom lxml import etree";
                script += "\nfrom bs4 import BeautifulSoup";
                script += "\nimport re";
                script += "\nfrom datetime import datetime";
                script += "\nprint(datetime.today())";
                script += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Server.Text.ToString() + ";DATABASE=Cognos Metadata;Trusted_Connection=yes;\")";
                script += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))";
                script += "\nconnection2=engine.connect()";
                script += "\nresoverall = connection2.execute(\"SELECT * FROM XMLData_Python order by [Index]\")";
                script += "\ndf = DataFrame(resoverall.fetchall())";
                script += "\ndf.columns = resoverall.keys()";
                script += "\nfor index, row in df.iterrows():";
                script += "\n    bs = BeautifulSoup(row['XMLData'], 'xml')";
                script += "\n    pretty_xml = bs.prettify()";
                script += "\n    filedata = pretty_xml";
                script += "\n    filedata = filedata.replace('xmlns', 'random')";
                script += "\n    filedata = filedata.replace('<root>', '')";
                script += "\n    filedata = filedata.replace('</root>', '')";
                script += "\n    filedata = re.sub('<staticValue>.*?</staticValue>','<staticValue> Dummy Text </staticValue>',filedata, flags=re.DOTALL)";
                script += "\n    filedata = filedata.encode('ascii', errors='ignore').decode(\"utf - 8\")";
                script += "\n    filedata = re.sub(r'[\x80-\xff]+', \"\", filedata)";
                script += "\n    #filedata = filedata.replace('<layouts/>','')";
                script += "\n    with open('" + xmlpath.ToString().Replace("\\", "\\\\") + "', 'w') as file:";
                script += "\n        file.write(filedata)";
                script += "\n    tree=et.parse(\"" + xmlpath.ToString().Replace("\\", "\\\\") + "\")";
                script += "\n    treemod =  etree.parse(\"" + xmlpath.ToString().Replace("\\", "\\\\") + "\")";
                script += "\n    root=tree.getroot()";
                script += "\n    rootmod=treemod.getroot()";
                script += "\n    #print(root)";
                script += "\n    ReportforDF=[]";
                script += "\n    Report=[]";
                script += "\n    QueryforDF=[]";
                script += "\n    QueryName=[]";
                script += "\n    QueryText=[]";
                script += "\n    Datasource=[]";
                script += "\n    Query=[]";
                script += "\n    DataItem=[]";
                script += "\n    DataItemAgg=[]";
                script += "\n    Expression=[]";
                script += "\n    FilterExpr=[]";
                script += "\n    FilterExprUse=[]";
                script += "\n    ReportVariables=[]";
                script += "\n    ReportVariableType=[]";
                script += "\n    ReportVariableExpr=[]";
                script += "\n    ModTime=[]";
                script += "\n    for y in root.iter('reportName'):";
                script += "\n        rootName=et.Element('root')";
                script += "\n        rootName=(y)";
                script += "\n        #print(rootName)";
                script += "\n        for em in rootName.iter('reportName'):";
                script += "\n            Report.append(em.text)";
                script += "\n    list = rootmod.xpath(\"//XMLAttribute[contains(@name,'RS_modelModificationTime')]\")";
                script += "\n    if not list:";
                script += "\n        list=\"1900 - 01 - 01\"";
                script += "\n    else:";
                script += "\n        ModTime.append (list[0].attrib['value'])";
                script += "\n    for y in root.iter('reportVariables'):";
                script += "\n        rootvar=et.Element('root')";
                script += "\n        rootvar=(y)";
                script += "\n        #print(rootvar)";
                script += "\n        for em in rootvar.iter('reportVariable'):";
                script += "\n            rootvar1=et.Element('root')";
                script += "\n            rootvar1=em ";
                script += "\n            for emrepexp in em.iter('reportExpression'):";
                script += "\n                ReportVariables.append(em.attrib['name'])";
                script += "\n                ReportVariableType.append(em.get('type'))";
                script += "\n                ReportVariableExpr.append(emrepexp.text)           ";
                script += "\n    for y in root.iter('reportName'):";
                script += "\n        rootName=et.Element('root')";
                script += "\n        rootName=(y)";
                script += "\n        #print(rootName)";
                script += "\n        for em in rootName.iter('reportName'):";
                script += "\n            ReportforDF.append(em.text)        ";
                script += "\n    for x in root.iter('queries'):";
                script += "\n        root1=et.Element('root')";
                script += "\n        root1=x";
                script += "\n        for supply in root1.iter('query'):";
                script += "\n            root2=et.Element('root')";
                script += "\n            root2=(supply) ";
                script += "\n            for source in root2.iter('source'):";
                script += "\n                for queryname in source.iter('sqlQuery'):";
                script += "\n                    rootq=et.Element('root')";
                script += "\n                    rootq=(queryname)";
                script += "\n                    for sqltext in source.iter('sqlText'):";
                script += "\n                        rootqtext=et.Element('root')";
                script += "\n                        rootq=(sqltext)";
                script += "\n            for tech in root2.iter('dataItem'):";
                script += "\n                root3 = et.Element('root')";
                script += "\n                root3=(tech)";
                script += "\n                for expr in root3.iter('expression'):";
                script += "\n                    root4 = et.Element('root')";
                script += "\n                    root4=(expr)";
                script += "\n                    for expr in root4.iter('expression'):";
                script += "\n                        Expression.append(expr.text)";
                script += "\n                        QueryforDF.append(supply.attrib['name'])";
                script += "\n                        DataItem.append(tech.attrib['name'])";
                script += "\n                        DataItemAgg.append(tech.get('aggregate'))";
                script += "\n                        QueryName.append(queryname.attrib['name'])";
                script += "\n                        Datasource.append(queryname.attrib['dataSource'])";
                script += "\n                        QueryText.append(sqltext.text)          ";
                script += "\n            for detailfs in root2.iter('detailFilters'):";
                script += "\n                rootdetailfs=et.Element('root')";
                script += "\n                rootdetailfs=detailfs";
                script += "\n                for detailf in detailfs.iter('detailFilter'):";
                script += "\n                    rootdetailf=et.Element('root')";
                script += "\n                    rootdetailf=detailf ";
                script += "\n                    for emfilt in detailf.iter('filterExpression'):";
                script += "\n                        Query.append(supply.attrib['name'])";
                script += "\n                        FilterExpr.append(emfilt.text)";
                script += "\n                        FilterExprUse.append(detailf.get('use'))";
                script += "\n    a = {'Report': ReportforDF,'Query':QueryforDF,'DataItem':DataItem,'Data Item Aggregate':DataItemAgg,'Expression':Expression,'Query Name':QueryName,'Datasource':Datasource,'Query Text':QueryText}";
                script += "\n    b = {'Report': Report,'Query':Query,'Filter Expression':FilterExpr,'Filter Expression Use':FilterExprUse}";
                script += "\n    c=  {'Report': Report,'Report Variables':ReportVariables,'Expression':ReportVariableExpr,'Report Variable Type':ReportVariableType}";
                script += "\n    d=  {'Report': Report,'Report Modification Time':ModTime}";
                script += "\n    df = pd.DataFrame.from_dict(a, orient='index')";
                script += "\n    df=df.transpose()  ";
                script += "\n    df['Report'] = df['Report'].ffill()";
                script += "\n    df1 = pd.DataFrame.from_dict(b, orient='index')";
                script += "\n    df1=df1.transpose()  ";
                script += "\n    df1['Report'] = df1['Report'].ffill()";
                script += "\n    df2 = pd.DataFrame.from_dict(c, orient='index')";
                script += "\n    df2=df2.transpose()  ";
                script += "\n    df2['Report'] = df2['Report'].ffill()";
                script += "\n    df3 = pd.DataFrame.from_dict(d, orient='index')";
                script += "\n    df3=df3.transpose()  ";
                script += "\n    df3['Report'] = df3['Report'].ffill() ";
                script += "\n    quoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Server.Text.ToString() + ";DATABASE=Cognos Metadata;Trusted_Connection=yes;\")";
                script += "\n    engine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))";
                script += "\n    df.to_sql('CognosDataItems', schema='dbo',if_exists = 'append', con = engine)";
                script += "\n    df1.to_sql('CognosFiltersExpression', schema='dbo',if_exists = 'append', con = engine)";
                script += "\n    df2.to_sql('CognosReportVariables', schema='dbo',if_exists = 'append', con = engine)";
                script += "\n    df3.to_sql('CognosReportModificationTime', schema='dbo',if_exists = 'append', con = engine)    ";



                /* string script = "import sys";
                 script += "\nimport numpy";
                 script += "\ndef add_numbers(x,y):";
                 script += "\n   sum = x + y";
                 script += "\n   return sum";
                 script += "\n";
                 script += "\nnum1 = int(sys.argv[1])";
                 script += "\nnum2 = int(sys.argv[2])";
                 script += "\nprint(add_numbers(num1, num2))";
                */

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

            


                // ProcessStartInfo myProcessStartInfo = new ProcessStartInfo(python);

                // createsqlDatabase();
                try
                {
                    createsqltableUsage();
                    run_cmd();

                   

                    string QueryDI = "select count(*) from dbo.CognosDataitems";
                    //Execute Queries and save results into variables
                    SqlCommand CmdCnt = SQLConnection.CreateCommand();
                    CmdCnt.CommandText = QueryDI;
                    SQLConnection.Open();
                    int DataITemCnt = (Int32)CmdCnt.ExecuteScalar();
                    SQLConnection.Close();

                    string QueryFE = "select count(*) from dbo.CognosFiltersExpression";
                    //Execute Queries and save results into variables
                    SqlCommand CmdCntFE = SQLConnection.CreateCommand();
                    CmdCntFE.CommandText = QueryFE;
                    SQLConnection.Open();
                    int FECnt = (Int32)CmdCntFE.ExecuteScalar();
                    SQLConnection.Close();

                    string QueryRMT = "select count(*) from dbo.CognosReportModificationTime";
                    //Execute Queries and save results into variables
                    SqlCommand CmdCntRMT = SQLConnection.CreateCommand();
                    CmdCntRMT.CommandText = QueryRMT;
                    SQLConnection.Open();
                    int RMTCnt = (Int32)CmdCntRMT.ExecuteScalar();
                    SQLConnection.Close();


                    string QueryRV = "select count(*) from dbo.CognosReportVariables";
                    //Execute Queries and save results into variables
                    SqlCommand CmdCntRV = SQLConnection.CreateCommand();
                    CmdCntRV.CommandText = QueryRV;
                    SQLConnection.Open();
                    int RVCnt = (Int32)CmdCntRV.ExecuteScalar();
                    SQLConnection.Close();




                    //string fileName = "Tableau Metadata.pbix";
                    //string path1 = System.IO.Path.Combine(Environment.CurrentDirectory, @"Report\", fileName);
                    //Process.Start(path1);

                    if (DataITemCnt > 0 || FECnt > 0 || RMTCnt > 0 || RVCnt > 0)
                    {
                        Animation.Visibility = Visibility.Collapsed;
                        LabelReport.Visibility = Visibility.Visible;
                        Stack1.Visibility = Visibility.Visible;
                        LabelDatabaseServer.Visibility = Visibility.Visible;
                        BorderServer.Visibility = Visibility.Visible;
                        button1.Visibility = Visibility.Collapsed;
                        PDF.Visibility = Visibility.Collapsed;
                        ReqButton.Visibility = Visibility.Collapsed;
                        ProcessStart.Visibility = Visibility.Collapsed;
                        InsertintoLocal.Visibility = Visibility.Visible;
                        LabelPythonPath.Visibility = Visibility.Visible;
                        BorderPythonPath.Visibility = Visibility.Visible;
                        Browse_Copy.Visibility = Visibility.Visible;
                        Output.IsEnabled = true;
                        MetadataToolTip.Text = "Please Find the summary of items inserted into the server " + Source123.Text.ToString();
                        MetadataToolTip.AppendText(Environment.NewLine);
                        MetadataToolTip.AppendText("Number of Dataitems = " + DataITemCnt + "\r\n");
                        MetadataToolTip.AppendText("Number of Filter Expressions = " + FECnt + "\r\n");
                        MetadataToolTip.AppendText("Number of Report Variables = " + RVCnt + "\r\n");
                        MetadataToolTip.AppendText("Number of Reports with Report Modification Time = " + RMTCnt);

                        OutputToolTip.Text = "Generate Power BI Report - The generated metadata is presented in a read-able format in a Power BI Report";
                        OutputToolTip.AppendText(Environment.NewLine);
                        OutputToolTip.AppendText("Requirement Document Generator - Generate Requirement Document for easier hand-over which will help in migration");


                        MessageBox.Show("Metadata is generated successfully. Click on Reset to start the process again.");
                    }
                    else
                    {
                        Animation.Visibility = Visibility.Collapsed;
                        LabelReport.Visibility = Visibility.Visible;
                        Stack1.Visibility = Visibility.Visible;
                        LabelDatabaseServer.Visibility = Visibility.Visible;
                        BorderServer.Visibility = Visibility.Visible;
                        button1.Visibility = Visibility.Collapsed;
                        PDF.Visibility = Visibility.Collapsed;
                        ReqButton.Visibility = Visibility.Collapsed;
                        ProcessStart.Visibility = Visibility.Collapsed;
                        InsertintoLocal.Visibility = Visibility.Visible;
                        LabelPythonPath.Visibility = Visibility.Visible;
                        BorderPythonPath.Visibility = Visibility.Visible;
                        Browse_Copy.Visibility = Visibility.Visible;
                        Output.IsEnabled = false;

                        //MetadataToolTip.Text = "The number of XML's available in the server " + Source123.Text.ToString() + " and Database = Cognos Metadata";
                        MetadataToolTip.Text = "Number of Distinct XML's available in the server " + Source123.Text.ToString() + " = " + XMLCnt.ToString()+"\r\n";
                        MetadataToolTip.AppendText("Make sure the database has enough capacity to process the XML's ");
                        MetadataToolTip.AppendText(Environment.NewLine);
                        MetadataToolTip.AppendText(Environment.NewLine);
                        MetadataToolTip.AppendText("Tip : In case you are using your localhost server , the maximum limit of Daatbase is 10 GB");

                        MessageBox.Show(" HERE Some issue in fetching the rows from the XML Data.Please check the XML or contact the adminstrator in case of any support needed");

                    }
                }
                catch (Exception ex)
                {

                    Animation.Visibility = Visibility.Collapsed;
                    LabelReport.Visibility = Visibility.Visible;
                    Stack1.Visibility = Visibility.Visible;
                    LabelDatabaseServer.Visibility = Visibility.Visible;
                    BorderServer.Visibility = Visibility.Visible;
                    button1.Visibility = Visibility.Collapsed;
                    PDF.Visibility = Visibility.Collapsed;
                    ReqButton.Visibility = Visibility.Collapsed;
                    ProcessStart.Visibility = Visibility.Collapsed;
                    InsertintoLocal.Visibility = Visibility.Visible;
                    LabelPythonPath.Visibility = Visibility.Visible;
                    BorderPythonPath.Visibility = Visibility.Visible;
                    Browse_Copy.Visibility = Visibility.Visible;
                    Output.IsEnabled = false;

                   // MetadataToolTip.Text = "The number of XML's available in the server " + Source123.Text.ToString() + " and Database = Cognos Metadata\r\n";
                    MetadataToolTip.Text = "Number of Distinct XML's available in the server " + Source123.Text.ToString() + " = " + XMLCnt.ToString() + "\r\n";
                    MetadataToolTip.AppendText("Make sure the database has enough capacity to process the XML's ");
                    MetadataToolTip.AppendText(Environment.NewLine);
                    MetadataToolTip.AppendText(Environment.NewLine);
                    MetadataToolTip.AppendText("Tip : In case you are using your localhost server , the maximum limit of Daatbase is 10 GB");
                    MessageBox.Show(ex.Message.ToString());
                   // MessageBox.Show("Some issue in fetching the rows from the XML Data.Please check the XML or contact the adminstrator in case of any support needed");

                }

            }
        }
    }
}

