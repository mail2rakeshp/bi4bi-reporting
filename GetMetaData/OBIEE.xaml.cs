using System;
using System.Collections.Generic;
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
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows;
using System.Security.Permissions;
using IronPython;
using Microsoft.Scripting.Hosting;
using System.Diagnostics;
using System.IO;
using IronPython.Hosting;
using Microsoft.Scripting;
using System.Data.SqlClient;
using System.Threading;

namespace GetMetadata
{
    /// <summary>
    /// Interaction logic for OBIEE.xaml
    /// </summary>
    public partial class OBIEE : Window
    {
       private System.Windows.Forms.NotifyIcon MyNotifyIcon;

        private static string PythonPath1;
        private static string TemplatePathString;
        private static string DestinationPathString;
        string server;

        string pythout = "";
        int XMLCnt=0;
        BackgroundWorker backgroundWorker1 = new BackgroundWorker();
        public OBIEE()
        {
            InitializeComponent();
            ProcessStart.Visibility = Visibility.Visible;
            GeneratePBI.Visibility = Visibility.Collapsed;
            GenerateDoc.Visibility = Visibility.Collapsed;
            Labelusername.Visibility = Visibility.Collapsed;
            Borderusername.Visibility = Visibility.Collapsed;
            Labelpasswd.Visibility = Visibility.Collapsed;
            Borderpasswd.Visibility = Visibility.Collapsed;
            PasswordChek.Visibility = Visibility.Collapsed;
            ProcessImage.Visibility = Visibility.Collapsed;
            OutputImage.Visibility = Visibility.Collapsed;
            Output.IsEnabled = false;
            WindRad.IsChecked = true;
            GenerateMetadata.IsChecked = true;

            InitializeComponent();
            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            MyNotifyIcon.Icon = new System.Drawing.Icon(
                            @"Final.ico");
            MyNotifyIcon.MouseDoubleClick +=
                new System.Windows.Forms.MouseEventHandler(MyNotifyIcon_MouseDoubleClick);

            MyNotifyIcon.BalloonTipClicked += new EventHandler(MyNotifyIcon_BalloonTipClicked);
            MyNotifyIcon.BalloonTipShown+= new EventHandler(MyNotifyIcon_BalloonTipClicked);
            MyNotifyIcon.BalloonTipClosed+= new EventHandler(MyNotifyIcon_BalloonTipClicked);
            ImageToolTip.Text = "Fill the mandatory fields in the below sequence : ";
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("1. CSV Folder Path -  Directory where the Catalog and RPD CSV's are available");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("2. SQL Server Details - Server where the Metadata needs to be inserted (Windows or SQL Authentication) ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("3. Python Path - Directory where the python is installed in the system ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("4. Generate Metadata - Insert XML's and start the process for Metadata generation ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("5. Generate Output/Requirement Doc - generate output or requirement document based on the metadata inserted in Step 4 ");

            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            backgroundWorker1.ProgressChanged += backgroundWorker1_ProgressChanged;
            backgroundWorker1.RunWorkerCompleted += backgroundWorker1_RunWorkerCompleted;  //Tell the user how the process went
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
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

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {

            //ExecuteMethodAsync();
            createsqlDatabase();
            createsqltableUsage();
            run_cmd();
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
            if (WindRad.IsChecked == true)
            {
                try
                {

                    SqlConnection SQLConnection = new SqlConnection();
                    SQLConnection.ConnectionString = "Data Source =" + Source123.Text.ToString() + "; Initial Catalog =OBIEE Metadata; " + "Integrated Security=true;";

                    string QueryXML = "select count(*) from dbo.Catalog_OBIEE";
                    //Execute Queries and save results into variables
                    SqlCommand CmdCntXML = SQLConnection.CreateCommand();
                    CmdCntXML.CommandText = QueryXML;
                    SQLConnection.Open();
                    int XMLCnt = (Int32)CmdCntXML.ExecuteScalar();
                    SQLConnection.Close();

                    string QueryXML1 = "select count(*) from dbo.RPD_OBIEE";
                    //Execute Queries and save results into variables
                    SqlCommand CmdCntXML1 = SQLConnection.CreateCommand();
                    CmdCntXML1.CommandText = QueryXML1;
                    SQLConnection.Open();
                    int XMLCnt1 = (Int32)CmdCntXML1.ExecuteScalar();
                    SQLConnection.Close();


                    if (XMLCnt == 0 && XMLCnt1 == 0)
                    {
                        Animation.Visibility = Visibility.Collapsed;
                        TemplatePath.Visibility = Visibility.Visible;
                        BorderTemplatePAth.Visibility = Visibility.Visible;
                        Template_Browse.Visibility = Visibility.Visible;
                        LabelSource.Visibility = Visibility.Visible;
                        BorderSource.Visibility = Visibility.Visible;
                        WindRad.Visibility = Visibility.Visible;
                        AuthRad.Visibility = Visibility.Visible;
                        LabelDatabaseServer.Visibility = Visibility.Visible;
                        BorderServer.Visibility = Visibility.Visible;
                        Browse.Visibility = Visibility.Visible;
                        GenerateMetadata.Visibility = Visibility.Visible;
                        Output.Visibility = Visibility.Visible;
                        ProcessStart.Visibility = Visibility.Collapsed;
                        ProcessImage.Visibility = Visibility.Collapsed;
                        OutputImage.Visibility = Visibility.Collapsed;
                        this.ShowInTaskbar = true;
                        MyNotifyIcon.Visible = true;
                        MyNotifyIcon.BalloonTipTitle = "Notification";
                        MyNotifyIcon.BalloonTipText = "Issues found in the details provided. Please contact the adminstrator in case of any support needed ";
                        MyNotifyIcon.ShowBalloonTip(5000);
                        //Thread.Sleep(5000);
                        //MyNotifyIcon.Dispose();

                    }
                    else
                    {
                        //MessageBox.Show("XML's inserted Successfully. Hover over the Tip icons for more info");

                        //Thread.Sleep(5000);
                        //MyNotifyIcon.Dispose();

                        MetadataToolTip.Text = "The Metadata been inserted into the server " + Source123.Text.ToString() + " and Database = OBIEE Metadata";
                        MetadataToolTip.AppendText(Environment.NewLine);
                        MetadataToolTip.AppendText("Number of Rows in Catalog=" + XMLCnt.ToString() + "\r\n");
                        MetadataToolTip.AppendText("Number of Rows in RPD=" + XMLCnt1.ToString() + "\r\n");


                        OutputToolTip.Text = "Generate Power BI Report - The generated metadata is presented in a read-able format in a Power BI Report";
                        OutputToolTip.AppendText(Environment.NewLine);
                        OutputToolTip.AppendText("Requirement Document Generator - Generate Requirement Document for easier hand-over which will help in migration");

                        //ImageToolTip.Content= "Number of Distinct XML's processed=" + pythout.ToString();
                        Animation.Visibility = Visibility.Collapsed;
                        TemplatePath.Visibility = Visibility.Visible;
                        BorderTemplatePAth.Visibility = Visibility.Visible;
                        Template_Browse.Visibility = Visibility.Visible;
                        LabelSource.Visibility = Visibility.Visible;
                        BorderSource.Visibility = Visibility.Visible;
                        WindRad.Visibility = Visibility.Visible;
                        AuthRad.Visibility = Visibility.Visible;
                        LabelDatabaseServer.Visibility = Visibility.Visible;
                        BorderServer.Visibility = Visibility.Visible;
                        Browse.Visibility = Visibility.Visible;
                        GenerateMetadata.Visibility = Visibility.Visible;
                        Output.Visibility = Visibility.Visible;
                        ProcessStart.Visibility = Visibility.Visible;
                        ProcessImage.Visibility = Visibility.Visible;
                        OutputImage.Visibility = Visibility.Visible;



                        TextCSV.IsEnabled = false;
                        Source123.IsEnabled = false;
                        TextPython.IsEnabled = false;
                        Output.IsEnabled = true;
                    }


                }
                catch (Exception ex)
                {
                    Animation.Visibility = Visibility.Collapsed;
                    TemplatePath.Visibility = Visibility.Visible;
                    BorderTemplatePAth.Visibility = Visibility.Visible;
                    Template_Browse.Visibility = Visibility.Visible;
                    LabelSource.Visibility = Visibility.Visible;
                    BorderSource.Visibility = Visibility.Visible;
                    WindRad.Visibility = Visibility.Visible;
                    AuthRad.Visibility = Visibility.Visible;
                    LabelDatabaseServer.Visibility = Visibility.Visible;
                    BorderServer.Visibility = Visibility.Visible;
                    Browse.Visibility = Visibility.Visible;
                    GenerateMetadata.Visibility = Visibility.Visible;
                    Output.Visibility = Visibility.Visible;
                    ProcessStart.Visibility = Visibility.Collapsed;
                    ProcessImage.Visibility = Visibility.Collapsed;
                    OutputImage.Visibility = Visibility.Collapsed;
                    MessageBox.Show(ex.Message.ToString());
                    // MessageBox.Show("Issues found in the details provided. Please contact the adminstrator in case of any support needed");
                }
            }
            else if (AuthRad.IsChecked == true)
            {
                try
                {
                   

                    SqlConnection SQLConnection = new SqlConnection();
                    SQLConnection.ConnectionString = "Data Source =" + Source123.Text.ToString() + "; Initial Catalog =OBIEE Metadata; " + "Integrated Security=true;";

                    string QueryXML = "select count(*) from dbo.Catalog_OBIEE";
                    //Execute Queries and save results into variables
                    SqlCommand CmdCntXML = SQLConnection.CreateCommand();
                    CmdCntXML.CommandText = QueryXML;
                    SQLConnection.Open();
                    int XMLCnt = (Int32)CmdCntXML.ExecuteScalar();
                    SQLConnection.Close();

                    string QueryXML1 = "select count(*) from dbo.RPD_OBIEE";
                    //Execute Queries and save results into variables
                    SqlCommand CmdCntXML1 = SQLConnection.CreateCommand();
                    CmdCntXML1.CommandText = QueryXML1;
                    SQLConnection.Open();
                    int XMLCnt1 = (Int32)CmdCntXML1.ExecuteScalar();
                    SQLConnection.Close();

                    if (XMLCnt == 0 && XMLCnt1 == 0)
                    {
                        Animation.Visibility = Visibility.Collapsed;
                        TemplatePath.Visibility = Visibility.Visible;
                        BorderTemplatePAth.Visibility = Visibility.Visible;
                        Template_Browse.Visibility = Visibility.Visible;
                        LabelSource.Visibility = Visibility.Visible;
                        BorderSource.Visibility = Visibility.Visible;
                        WindRad.Visibility = Visibility.Visible;
                        AuthRad.Visibility = Visibility.Visible;
                        LabelDatabaseServer.Visibility = Visibility.Visible;
                        BorderServer.Visibility = Visibility.Visible;
                        Browse.Visibility = Visibility.Visible;
                        GenerateMetadata.Visibility = Visibility.Visible;
                        Output.Visibility = Visibility.Visible;
                        ProcessStart.Visibility = Visibility.Collapsed;
                        PasswordChek.Visibility = Visibility.Visible;
                        ProcessImage.Visibility = Visibility.Collapsed;
                        OutputImage.Visibility = Visibility.Collapsed;
                        Labelusername.Visibility = Visibility.Collapsed;
                        Borderusername.Visibility = Visibility.Collapsed;
                        Labelpasswd.Visibility = Visibility.Collapsed;
                        if (PasswordChek.IsChecked == true)
                        {
                            Borderpasswd.Visibility = Visibility.Visible;
                        }
                        else
                        {
                            BorderPasswordShow.Visibility = Visibility.Visible;
                        }
                        this.ShowInTaskbar = true;
                        MyNotifyIcon.Visible = true;
                        MyNotifyIcon.BalloonTipTitle = "Notification";
                        MyNotifyIcon.BalloonTipText = "Issues found in the details provided. Please contact the adminstrator in case of any support needed ";
                        MyNotifyIcon.ShowBalloonTip(5000);
                        //Thread.Sleep(5000);
                        //MyNotifyIcon.Dispose();

                        //MessageBox.Show("Here");
                        // MessageBox.Show("Issues found in the details provided. Please contact the adminstrator in case of any support needed ");
                    }
                    else
                    {
                        MetadataToolTip.Text = "The Metadata have been inserted into the server " + Source123.Text.ToString() + " and Database = OBIEE Metadata";
                        MetadataToolTip.AppendText(Environment.NewLine);
                        MetadataToolTip.AppendText("Number of Rows in Catalog=" + XMLCnt.ToString() + "\r\n");
                        MetadataToolTip.AppendText("Number of Rows in RPD=" + XMLCnt1.ToString() + "\r\n");


                        OutputToolTip.Text = "Generate Power BI Report - The generated metadata is presented in a read-able format in a Power BI Report";
                        OutputToolTip.AppendText(Environment.NewLine);
                        OutputToolTip.AppendText("Requirement Document Generator - Generate Requirement Document for easier hand-over which will help in migration");

                        //ImageToolTip.Content= "Number of Distinct XML's processed=" + pythout.ToString();
                        Animation.Visibility = Visibility.Collapsed;
                        TemplatePath.Visibility = Visibility.Visible;
                        BorderTemplatePAth.Visibility = Visibility.Visible;
                        Template_Browse.Visibility = Visibility.Visible;
                        LabelSource.Visibility = Visibility.Visible;
                        BorderSource.Visibility = Visibility.Visible;
                        WindRad.Visibility = Visibility.Visible;
                        AuthRad.Visibility = Visibility.Visible;
                        LabelDatabaseServer.Visibility = Visibility.Visible;
                        BorderServer.Visibility = Visibility.Visible;
                        Browse.Visibility = Visibility.Visible;
                        GenerateMetadata.Visibility = Visibility.Visible;
                        Output.Visibility = Visibility.Visible;
                        ProcessStart.Visibility = Visibility.Visible;
                        ProcessImage.Visibility = Visibility.Visible;
                        OutputImage.Visibility = Visibility.Visible;

                        PasswordChek.Visibility = Visibility.Visible;
                        Labelusername.Visibility = Visibility.Collapsed;
                        Borderusername.Visibility = Visibility.Collapsed;
                        Labelpasswd.Visibility = Visibility.Collapsed;

                        Output.IsEnabled = true;
                        if (PasswordChek.IsChecked == true)
                        {
                            Borderpasswd.Visibility = Visibility.Visible;
                        }
                        else
                        {
                            BorderPasswordShow.Visibility = Visibility.Visible;
                        }
                    }


                }
                catch (Exception ex)
                {
                    Animation.Visibility = Visibility.Collapsed;
                    TemplatePath.Visibility = Visibility.Visible;
                    BorderTemplatePAth.Visibility = Visibility.Visible;
                    Template_Browse.Visibility = Visibility.Visible;
                    LabelSource.Visibility = Visibility.Visible;
                    BorderSource.Visibility = Visibility.Visible;
                    WindRad.Visibility = Visibility.Visible;
                    AuthRad.Visibility = Visibility.Visible;
                    LabelDatabaseServer.Visibility = Visibility.Visible;
                    BorderServer.Visibility = Visibility.Visible;
                    Browse.Visibility = Visibility.Visible;
                    GenerateMetadata.Visibility = Visibility.Visible;
                    Output.Visibility = Visibility.Visible;
                    ProcessStart.Visibility = Visibility.Collapsed;
                    ProcessImage.Visibility = Visibility.Collapsed;
                    OutputImage.Visibility = Visibility.Collapsed;
                    PasswordChek.Visibility = Visibility.Visible;
                    Labelusername.Visibility = Visibility.Collapsed;
                    Borderusername.Visibility = Visibility.Collapsed;
                    Labelpasswd.Visibility = Visibility.Collapsed;
                    if (PasswordChek.IsChecked == true)
                    {
                        Borderpasswd.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        BorderPasswordShow.Visibility = Visibility.Visible;
                    }
                    MessageBox.Show(ex.Message.ToString());
                    //MessageBox.Show("Issues found in the details provided. Please contact the adminstrator in case of any support needed");
                }
            }
        }
            private void ProcessStart_Click(object sender, RoutedEventArgs e)
        {
            string path = Directory.GetCurrentDirectory() + @"\PythonFile\OBIEE_Python.py";
            //MessageBox.Show(path.ToString());
            string password = "";

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
            if (WindRad.IsChecked == true)
            {
                if (String.IsNullOrEmpty(TextCSV.Text.ToString()) || String.IsNullOrEmpty(Source123.Text.ToString()) || String.IsNullOrEmpty(TextPython.Text.ToString())
                    || (GenerateMetadata.IsChecked == false && Output.IsChecked == false))
                {
                    MessageBox.Show("Please enter the mandatory fields and try again");
                }
                else
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
                    GenerateMetadata.Visibility = Visibility.Collapsed;
                    Output.Visibility = Visibility.Collapsed;
                    GeneratePBI.Visibility = Visibility.Collapsed;
                    GenerateDoc.Visibility = Visibility.Collapsed;
                    ProcessStart.Visibility = Visibility.Collapsed;
                    ProcessImage.Visibility = Visibility.Collapsed;
                    OutputImage.Visibility = Visibility.Collapsed;

                    Animation.Visibility = Visibility.Visible;



                    string script = "import csv";
                    script += "\nimport pandas as pd";
                    script += "\nimport urllib";
                    script += "\nfrom sqlalchemy import create_engine";
                    script += "\ndata_list = []";
                    script += "\nimport os";
                    script += "\nimport glob";
                    script += "\nfiles1 = glob.glob(\"" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "\\\\*Catalog.csv\")";
                    script += "\nPath1 = \"" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "\\\\\"";
                    script += "\nFile1 = os.path.basename(files1[0]) ";
                    script += "\nFullName1 = Path1+File1";
                    script += "\nfiles2 = glob.glob(\"" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "\\\\*RPD.csv\")";
                    script += "\nPath2 = \"" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "\\\\\"";
                    script += "\nFile2 = os.path.basename(files2[0]) ";
                    script += "\nFullName2 = Path2+File2";
                    script += "\nwith open(FullName1, mode ='r') as file:";
                    script += "\n    csvFile = csv.reader(file)";
                    script += "\n    for lines in csvFile:";
                    script += "\n        data_list.append(lines) ";
                    script += "\ncatalog = []";
                    script += "\nfor each_data in data_list:";
                    script += "\n    catalog.append(\",\".join(map(str, each_data)).replace('\"', '').split('\t'))";
                    script += "\ncatalog_df = pd.DataFrame(catalog[1:], columns=catalog[0])";
                    script += "\nRPD_df = pd.read_csv(FullName2)";
                    script += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=OBIEE Metadata;Trusted_Connection=yes;\")";
                    script += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))";
                    script += "\ncatalog_df.to_sql('Catalog_OBIEE', schema='dbo', con = engine)";
                    script += "\nRPD_df.to_sql('RPD_OBIEE', schema='dbo', con = engine)";
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

                }


            }
            else if (AuthRad.IsChecked == true)
            {
                if (String.IsNullOrEmpty(TextCSV.Text.ToString()) || String.IsNullOrEmpty(Source123.Text.ToString()) || String.IsNullOrEmpty(TextPython.Text.ToString())
                    || GenerateMetadata.IsChecked == false || Output.IsChecked == false || String.IsNullOrEmpty(username.Text.ToString()) || String.IsNullOrEmpty(password.ToString()))
                {
                    MessageBox.Show("Please enter the mandatory fields and try again");
                }
                else
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
                    GenerateMetadata.Visibility = Visibility.Collapsed;
                    Output.Visibility = Visibility.Collapsed;
                    GeneratePBI.Visibility = Visibility.Collapsed;
                    GenerateDoc.Visibility = Visibility.Collapsed;
                    ProcessStart.Visibility = Visibility.Collapsed;
                    ProcessImage.Visibility = Visibility.Collapsed;
                    OutputImage.Visibility = Visibility.Collapsed;

                    Animation.Visibility = Visibility.Visible;



                    string script = "import csv";
                    script += "\nimport pandas as pd";
                    script += "\nimport urllib";
                    script += "\nfrom sqlalchemy import create_engine";
                    script += "\ndata_list = []";
                    script += "\nimport os";
                    script += "\nimport glob";
                    script += "\nfiles1 = glob.glob(\"" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "\\\\*Catalog.csv\")";
                    script += "\nPath1 = \"" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "\\\\\"";
                    script += "\nFile1 = os.path.basename(files1[0]) ";
                    script += "\nFullName1 = Path1+File1";
                    script += "\nfiles2 = glob.glob(\"" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "\\\\*RPD.csv\")";
                    script += "\nPath2 = \"" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "\\\\\"";
                    script += "\nFile2 = os.path.basename(files2[0]) ";
                    script += "\nFullName2 = Path2+File2";
                    script += "\nwith open(FullName1, mode ='r') as file:";
                    script += "\n    csvFile = csv.reader(file)";
                    script += "\n    for lines in csvFile:";
                    script += "\n        data_list.append(lines) ";
                    script += "\ncatalog = []";
                    script += "\nfor each_data in data_list:";
                    script += "\n    catalog.append(\",\".join(map(str, each_data)).replace('\"', '').split('\t'))";
                    script += "\ncatalog_df = pd.DataFrame(catalog[1:], columns=catalog[0])";
                    script += "\nRPD_df = pd.read_csv(FullName2)";
                    script += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=OBIEE Metadata;UID=" + username.Text.ToString() + ";PWD=" + password + ";\")";
                    script += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))";
                    script += "\ncatalog_df.to_sql('Catalog_OBIEE', schema='dbo', con = engine)";
                    script += "\nRPD_df.to_sql('RPD_OBIEE', schema='dbo', con = engine)";
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
                    

                }

            }

            server = Source123.Text.ToString();
            backgroundWorker1.RunWorkerAsync();


        }

        private void GeneratePBI_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Generating Report. Please Wait ....");
            // string file = @"Metadata Output.pbix";
            string fileName = "BI4BI - OBIEE.pbix";
            string path = System.IO.Path.Combine(Environment.CurrentDirectory, @"Report\", fileName);
            Process.Start(path);
        }

        private void GenerateDoc_Click(object sender, RoutedEventArgs e)
        {
            
        }
        private void GenerateMetadata_Checked(object sender, RoutedEventArgs e)
        {
            GeneratePBI.Visibility = Visibility.Collapsed;
            GenerateDoc.Visibility = Visibility.Collapsed;
            ProcessStart.Visibility = Visibility.Visible;
        }

        private void Output_Checked(object sender, RoutedEventArgs e)
        {
            ProcessStart.Visibility = Visibility.Collapsed;
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

        private void SignOutButton_Click(object sender, RoutedEventArgs e)
        {

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
                        sw.WriteLine(PythonPath1 + @"\Scripts\activate.bat");
                        // Activate your environment
                        // sw.WriteLine("conda activate py3.9.7");
                        // run your script. You can also pass in arguments
                        sw.WriteLine("py OBIEE_Python.py");
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
        private async void run_cmd_1()
        {
            //MessageBox.Show("Report Generation in process. Average Wait Time less than 1 minute.");

            string output = "";
            
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
                            sw.WriteLine("py OBIEE_Python.py");
                        }
                    }
                    //string output = process.StandardOutput.ReadToEnd();

                    process.WaitForExit();

                }
            catch (Exception ex)
            {
                //MessageBox.Show("Issues found in the details provided. Please contact the adminstrator in case of any support needed");


            }
            /*
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
                output = process.StandardOutput.ReadToEnd();

                process.WaitForExit();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }*/


        }

        public void createsqlDatabase()
        {
            string connectionString = @"Data Source = " + server.ToString().Replace("\\\\", "\\") + "; Integrated Security=true";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string strconnection = "Data Source = " + server.ToString() + "; Integrated Security=true";

            string table = "IF NOT EXISTS(SELECT name FROM master.dbo.sysdatabases WHERE Name='OBIEE Metadata') CREATE DATABASE[OBIEE Metadata]";
            InsertQuery1(table, strconnection);
        }
        public async void createsqltableUsage()
        {


            try
            {
                string connectionString = @"Data Source = " + server.ToString().Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=OBIEE Metadata";
                SqlConnection sqlconnection = new SqlConnection(connectionString);
                sqlconnection.Open();
                string strconnection = "Data Source = " + server.ToString() + "; Integrated Security=true; Initial Catalog=OBIEE Metadata";
                string table = "";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='Catalog_OBIEE') BEGIN DROP TABLE Catalog_OBIEE END";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='RPD_OBIEE') BEGIN DROP TABLE RPD_OBIEE END  ";
                InsertQuery1(table, strconnection);

            }
            catch
            {
                MessageBox.Show("Please check the SQL server Instance and try again");
            }


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

            GenerateMetadata.IsChecked = true;
            Output.IsChecked = false;
            ProcessStart.Visibility = Visibility.Visible;
            ProcessImage.Visibility = Visibility.Collapsed;
            OutputImage.Visibility = Visibility.Collapsed;


        }
    }
}




