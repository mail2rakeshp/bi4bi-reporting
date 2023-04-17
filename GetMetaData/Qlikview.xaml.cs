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

namespace GetMetaData
{
    /// <summary>
    /// Interaction logic for Qlikview.xaml
    /// </summary>
    public partial class Qlikview : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;
        private static string PythonPath1;
        public Qlikview()
        {
            InitializeComponent();
            Animation.Visibility = Visibility.Collapsed;


            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            MyNotifyIcon.Icon = new System.Drawing.Icon(
                            @"Final.ico");
            MyNotifyIcon.MouseDoubleClick +=
                new System.Windows.Forms.MouseEventHandler(MyNotifyIcon_MouseDoubleClick);
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

        private void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            Window1 window1 = new Window1();
            window1.ShowDialog();

        }
        private async void button1_Click(object sender, RoutedEventArgs e)
        {




            string path = Directory.GetCurrentDirectory() + @"\PythonFile\Qlikview_Python_File.py";
            //MessageBox.Show(path.ToString());
            string password = "";

            //string path = System.IO.Path.Combine(@"C:\Users\UT481LN\source\repos\GetMetaData\GetMetaData\GetMetaData\PythonFile",fileName);
            // string path = System.IO.Path.Combine(Environment.CurrentDirectory, @"PythonFile\", fileName);
            if (ResultText.Text.Equals("") || apiversion.Text.Equals("") || USerName.Text.Equals("") || SQLServer.Text.Equals("") || PythonPathText.Text.Equals(""))
            {
                MessageBox.Show("Please enter the mandatory fields and try again");
            }
            else
            {
                Animation.Visibility = Visibility.Visible;
                //ServerStack.Visibility = Visibility.Collapsed;
                LabelServer.Visibility = Visibility.Collapsed;
                Border1.Visibility = Visibility.Collapsed;
                Labelapiversion.Visibility = Visibility.Collapsed;
                Borderapiversion.Visibility = Visibility.Collapsed;
                LabelUserName.Visibility = Visibility.Collapsed;
                BorderUserName.Visibility = Visibility.Collapsed;
                //BorderPasswordShow.Visibility = Visibility.Collapsed;
                SQLServerL.Visibility = Visibility.Collapsed;
                LabelPythonPath.Visibility = Visibility.Collapsed;
                BorderPythonPath.Visibility = Visibility.Collapsed;
                Browse_Copy.Visibility = Visibility.Collapsed;
                button1.Visibility = Visibility.Collapsed;
                SignOutButton.Visibility = Visibility.Collapsed;
                MessageBox.Show("Report Generation in process. Average Wait Time less than 1 minute.");

                string url = "";
                int pos = ResultText.Text.LastIndexOf("/") ;

                if(pos>0)
                {
                    url = ResultText.Text.ToString();
                }
                else
                {
                    url = ResultText.Text.ToString()+"/";
                }

                string script = "import requests";
                script += "\nimport pandas as pd";
                script += "\nimport urllib";
                script += "\nfrom sqlalchemy import create_engine";
                script += "\nfrom pandas import json_normalize";
                script += "\nimport pyodbc ";
                script += "\nheaders = {";
                script += "\n   'Authorization': 'Bearer "+ USerName.Text.ToString()+ "',";
                script += "\n}";
                script += "\nparams = {";
                script += "\n   'resourceType': 'app',";
                script += "\n}";
                script += "\nresponse = requests.get('"+url+"api/v1/items', params=params, headers=headers)";
                script += "\nwb_df = json_normalize(response.json()['data'])";
                script += "\nwb_df=wb_df[[\"name\",\"resourceId\",\"createdAt\",\"updatedAt\",\"links.thumbnail.href\"]]";
                script += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + SQLServer.Text.ToString() + ";DATABASE=Power BI COE;Trusted_Connection=yes;\")";
                script += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))";
                script += "\nurl='" + url + "api/v1/apps/'+'" + apiversion.Text.ToString()+ "'+'/data/metadata'";
                script += "\nheaders = {";
                script += "\n'Authorization': 'Bearer "+ USerName.Text.ToString()+ "',";
                script += "\n          }";
                script += "\nresponse = requests.get(url, headers=headers)";
                script += "\ndashboard_info_json = response.json()";
                script += "\ndashboard_info_df=json_normalize(data=dashboard_info_json,record_path='fields')";
                script += "\ndashboard_info_df['app ID'] = '" + apiversion.Text.ToString() + "'";
                script += "\ndashboard_info_df = dashboard_info_df.astype({\"src_tables\": str})";
                script += "\ndashboard_info_df = dashboard_info_df.astype({\"tags\": str})";
                script += "\ndashboard_info_df.drop(columns=['hash'], inplace=True)";
                script += "\ntable_df=json_normalize(data=dashboard_info_json,record_path='tables')";
                script += "\ntable_df['app ID'] = '" + apiversion.Text.ToString() + "'";
                script += "\ndashboard_info_df.to_sql('QlikAppColumns', schema='dbo', con = engine)";
                script += "\nwb_df.to_sql('QlikApps', schema='dbo', con = engine)";
                script += "\ntable_df.to_sql('QlikAppTables', schema='dbo', con = engine)";
               

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

                createsqlDatabase();
                createsqltableUsage();
                run_cmd();

                string fileName = "Qlikview Metadata.pbix";
                string path1 = System.IO.Path.Combine(Environment.CurrentDirectory, @"Report\", fileName);
                Process.Start(path1);


                Animation.Visibility = Visibility.Collapsed;
                //ServerStack.Visibility = Visibility.Collapsed;
                LabelServer.Visibility = Visibility.Visible;
                Border1.Visibility = Visibility.Visible;
                Labelapiversion.Visibility = Visibility.Visible;
                Borderapiversion.Visibility = Visibility.Visible;
                LabelUserName.Visibility = Visibility.Visible;
                BorderUserName.Visibility = Visibility.Visible;
               // BorderPasswordShow.Visibility = Visibility.Visible;



                SQLServerL.Visibility = Visibility.Visible;
                LabelPythonPath.Visibility = Visibility.Visible;
                BorderPythonPath.Visibility = Visibility.Visible;
                Browse_Copy.Visibility = Visibility.Visible;
                button1.Visibility = Visibility.Visible;
                SignOutButton.Visibility = Visibility.Visible;

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
                        sw.WriteLine(PythonPath1 + @"\Scripts\activate.bat");
                        // Activate your environment
                        // sw.WriteLine("conda activate py3.9.7");
                        // run your script. You can also pass in arguments
                        sw.WriteLine("py Qlikview_Python_File.py");
                    }
                }
                // read multiple output lines
                /*while (!process.StandardOutput.EndOfStream)
                {
                    var line = process.StandardOutput.ReadLine();
                    //Console.WriteLine(line);
                    //Thread.Sleep(500);
                }*/
            }
            catch
            {
                MessageBox.Show("Please check the Qlik details and try again");
            }


        }
        public async void createsqltableUsage()
        {


            try
            {
                string connectionString = @"Data Source = " + SQLServer.Text.Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=Power BI COE";
                SqlConnection sqlconnection = new SqlConnection(connectionString);
                sqlconnection.Open();
                string strconnection = "Data Source = " + SQLServer.Text.ToString() + "; Integrated Security=true; Initial Catalog=Power BI COE";
                string table = "";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='QlikAppColumns') BEGIN DROP TABLE QlikAppColumns END";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='QlikApps') BEGIN DROP TABLE QlikApps END  ";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='QlikAppTables') BEGIN DROP TABLE QlikAppTables END";

                InsertQuery(table, strconnection);

            }
            catch
            {
                MessageBox.Show("Please check the SQL server Instance and try again");
            }


        }

        public void createsqlDatabase()
        {
            string connectionString = @"Data Source = " + SQLServer.Text.Replace("\\\\", "\\") + "; Integrated Security=true";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string strconnection = "Data Source = " + SQLServer.Text.ToString() + "; Integrated Security=true";

            string table = "IF NOT EXISTS(SELECT name FROM master.dbo.sysdatabases WHERE Name='Power BI COE') CREATE DATABASE [Power BI COE]";
            InsertQuery(table, strconnection);
        }

        public async void InsertQuery(string qry, string connection)
        {

            try
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
            catch (Exception ex)
            {
                MessageBox.Show("Please check the SQL server Instance and try again");
            }
        }


    }
}
