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
    /// Interaction logic for Tableau.xaml
    /// </summary>
    public partial class Tableau : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;
        private static string PythonPath1;
        string password = "";
        string apiver;
        string Tableausite;
        string siteurl;
        string sitename;
        string username;
        string server;
        BackgroundWorker backgroundWorker1 = new BackgroundWorker();
        public Tableau()
        {
            InitializeComponent();
            Animation.Visibility = Visibility.Collapsed;


            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            MyNotifyIcon.Icon = new System.Drawing.Icon(
                            @"Final.ico");
            MyNotifyIcon.MouseDoubleClick +=
                new System.Windows.Forms.MouseEventHandler(MyNotifyIcon_MouseDoubleClick);
            GenerateMetadata.IsChecked =true;
            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            backgroundWorker1.ProgressChanged += backgroundWorker1_ProgressChanged;
            backgroundWorker1.RunWorkerCompleted += backgroundWorker1_RunWorkerCompleted;  //Tell the user how the process went
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            PDF.IsEnabled = false;
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {

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

            SqlConnection SQLConnection = new SqlConnection();
            SQLConnection.ConnectionString = "Data Source =" + server.ToString() + ";Initial Catalog=Tableau Metadata; " + "Integrated Security=true;";
            string QueryDI = "select count(*) from dbo.TableauWorkbooks";
            //Execute Queries and save results into variables
            SqlCommand CmdCnt = SQLConnection.CreateCommand();
            CmdCnt.CommandText = QueryDI;
            SQLConnection.Open();
            //int DataITemCnt = (Int32)CmdCnt.ExecuteScalar();
            int DataITemCnt = Convert.ToInt32(CmdCnt.ExecuteScalar());
            
            SQLConnection.Close();

            string QueryFE = "select count(*) from dbo.TableauDatabaseServers";
            //Execute Queries and save results into variables
            SqlCommand CmdCntFE = SQLConnection.CreateCommand();
            CmdCntFE.CommandText = QueryFE;
            SQLConnection.Open();
            int FECnt = (Int32)CmdCntFE.ExecuteScalar();
            SQLConnection.Close();

            string QueryRMT = "select count(*) from dbo.TableauFileSources";
            //Execute Queries and save results into variables
            SqlCommand CmdCntRMT = SQLConnection.CreateCommand();
            CmdCntRMT.CommandText = QueryRMT;
            SQLConnection.Open();
            int RMTCnt = (Int32)CmdCntRMT.ExecuteScalar();
            SQLConnection.Close();


            string QueryRV = "select count(*) from dbo.TableauRefreshTime";
            //Execute Queries and save results into variables
            SqlCommand CmdCntRV = SQLConnection.CreateCommand();
            CmdCntRV.CommandText = QueryRV;
            SQLConnection.Open();
            int RVCnt = (Int32)CmdCntRV.ExecuteScalar();
            SQLConnection.Close();

            string QueryCalc = "select count(*) from dbo.TableauCalculations";
            //Execute Queries and save results into variables
            SqlCommand CmdCntCalc = SQLConnection.CreateCommand();
            CmdCntCalc.CommandText = QueryCalc;
            SQLConnection.Open();
            int CalcCnt = (Int32)CmdCntCalc.ExecuteScalar();
            SQLConnection.Close();
            if (DataITemCnt > 0 || FECnt > 0 || RMTCnt > 0 || RVCnt > 0 || CalcCnt > 0)
            {
                GenerateMetadata.IsChecked = false;
                Output.IsChecked = true;
                Generate_Metadata.Visibility = Visibility.Collapsed;
                button1.Visibility = Visibility.Visible;
                Req.Visibility = Visibility.Visible;
                PDF.Visibility = Visibility.Visible;
            }
            else
            {
                GenerateMetadata.IsChecked = true;
                Output.IsChecked = false;
                MessageBox.Show("Some issues in Metadata generation. Please contact the administrator for further assistance.");
                Generate_Metadata.Visibility = Visibility.Visible;
                button1.Visibility = Visibility.Collapsed;
                Req.Visibility = Visibility.Collapsed;
                PDF.Visibility = Visibility.Collapsed;
            }
            Animation.Visibility = Visibility.Collapsed;
            //ServerStack.Visibility = Visibility.Collapsed;
            LabelServer.Visibility = Visibility.Visible;
            Border1.Visibility = Visibility.Visible;
            Labelapiversion.Visibility = Visibility.Visible;
            Borderapiversion.Visibility = Visibility.Visible;
            LabelUserName.Visibility = Visibility.Visible;
            BorderUserName.Visibility = Visibility.Visible;
            LabelPassword.Visibility = Visibility.Visible;
            BorderPassword.Visibility = Visibility.Visible;
            BorderPasswordShow.Visibility = Visibility.Visible;
            PasswordChek.Visibility = Visibility.Visible;
            LabelSiteName.Visibility = Visibility.Visible;
            BorderSiteName.Visibility = Visibility.Visible;

            LabelSiteURL.Visibility = Visibility.Visible;
            BorderSiteURL.Visibility = Visibility.Visible;
            SQLServerL.Visibility = Visibility.Visible;
            LabelPythonPath.Visibility = Visibility.Visible;
            BorderPythonPath.Visibility = Visibility.Visible;
            Browse_Copy.Visibility = Visibility.Visible;
            SignOutButton.Visibility = Visibility.Visible;
            GenerateMetadata.Visibility = Visibility.Visible;
            Output.Visibility = Visibility.Visible;

        }
        private void GenerateMetadata_Click(object sender, RoutedEventArgs e)
        {


            
            if (PasswordChek.IsChecked == true)
            {
                password = PasswordShow.Text.ToString();
            }
            else
            {
                password = Password.Password;
            }

            Tableausite = ResultText.Text.ToString();
            apiver = apiversion.Text.ToString();
            username = USerName.Text.ToString();
            sitename = SiteName.Text.ToString();
            siteurl = SiteURL.Text.ToString();
            server = SQLServer.Text.ToString();
            PythonPath1 = PythonPathText.Text.ToString();
            if (Tableausite.ToString().Equals("") || apiver.ToString().Equals("") || username.ToString().Equals("") || password == "" || sitename.ToString().Equals("") || siteurl.ToString().Equals("") || server.ToString().Equals("") || PythonPath1.ToString().Equals(""))
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
                LabelPassword.Visibility = Visibility.Collapsed;
                BorderPassword.Visibility = Visibility.Collapsed;
                BorderPasswordShow.Visibility = Visibility.Collapsed;
                PasswordChek.Visibility = Visibility.Collapsed;
                LabelSiteName.Visibility = Visibility.Collapsed;
                BorderSiteName.Visibility = Visibility.Collapsed;

                LabelSiteURL.Visibility = Visibility.Collapsed;
                BorderSiteURL.Visibility = Visibility.Collapsed;
                SQLServerL.Visibility = Visibility.Collapsed;
                LabelPythonPath.Visibility = Visibility.Collapsed;
                BorderPythonPath.Visibility = Visibility.Collapsed;
                Browse_Copy.Visibility = Visibility.Collapsed;
                button1.Visibility = Visibility.Collapsed;
                SignOutButton.Visibility = Visibility.Collapsed;
                GenerateMetadata.Visibility = Visibility.Collapsed;
                Output.Visibility = Visibility.Collapsed;
                Generate_Metadata.Visibility = Visibility.Collapsed;
                PDF.Visibility = Visibility.Collapsed;
                SQLServerL.Visibility = Visibility.Collapsed;
                SQLServerLocal.Visibility = Visibility.Collapsed;
                SQLServer.Visibility = Visibility.Collapsed;
                MessageBox.Show("Metadata Generation in process");
                backgroundWorker1.RunWorkerAsync();
            }

            }
        private void ExecuteMethodAsync()
        {
            string path = Directory.GetCurrentDirectory() + @"\PythonFile\Tableau_Python.py";
            //MessageBox.Show(path.ToString());


                string script = "import pandas as pd                                                                                                                                                                                      ";
                script += "\nimport numpy as np                                                                                                                                                                                       ";
                script += "\nimport urllib                                                                                                                                                                                            ";
                script += "\nfrom sqlalchemy import create_engine                                                                                                                                                                     ";
                script += "\nimport pyodbc                                                                                                                                                                                            ";
                script += "\nfrom tableau_api_lib import TableauServerConnection                                                                                                                                                      ";
                script += "\nfrom tableau_api_lib.utils import flatten_dict_column, flatten_dict_list_column";
                script += "\ntableau_server_config = {";
                script += "\n         'tableau_prod': {";
                script += "\n                 'server': '" + Tableausite.ToString() + "',";
                script += "\n                 'api_version': '" + apiver.ToString() + "',";
                script += "\n                 'personal_access_token_name': '" + username.ToString() + "',";
                script += "\n                 'personal_access_token_secret': '" + password + "',"; ;
                script += "\n                 'site_name': '" + sitename.ToString() + "',";
                script += "\n                 'site_url':  '" + siteurl.ToString() + "'";
                script += "\n         }";
                script += "\n}"; 
                script += "\nconn = TableauServerConnection(tableau_server_config, ssl_verify = False)                                                                                                                                ";
                script += "\nconn.sign_in()                                                                                                                                                                                           ";
                script += "\nquery_workbooks =  \"\"\"                                                                                                                                                                               ";
                script += "\n{                                                                                                                                                                                                        ";
                script += "\nworkbooks {                                                                                                                                                                                              ";
                script += "\n    id                                                                                                                                                                                                   ";
                script += "\n    luid                                                                                                                                                                                                 ";
                script += "\n    name                                                                                                                                                                                                 ";
                script += "\n    description                                                                                                                                                                                          ";
                script += "\n    createdAt                                                                                                                                                                                            ";
                script += "\n    updatedAt                                                                                                                                                                                            ";
                script += "\n    site {                                                                                                                                                                                               ";
                script += "\n        luid                                                                                                                                                                                             ";
                script += "\n        name                                                                                                                                                                                             ";
                script += "\n      }                                                                                                                                                                                                  ";
                script += "\n    projectName                                                                                                                                                                                          ";
                script += "\n    projectVizportalUrlId                                                                                                                                                                                ";
                script += "\n    owner {                                                                                                                                                                                              ";
                script += "\n        id                                                                                                                                                                                               ";
                script += "\n        name                                                                                                                                                                                             ";
                script += "\n        }                                                                                                                                                                                                ";
                script += "\n    uri                                                                                                                                                                                                  ";
                script += "\n    upstreamDatasources {                                                                                                                                                                                ";
                script += "\n        id                                                                                                                                                                                               ";
                script += "\n        luid                                                                                                                                                                                             ";
                script += "\n        name                                                                                                                                                                                             ";
                script += "\n        }                                                                                                                                                                                                ";
                script += "\n    }                                                                                                                                                                                                    ";
                script += "\n}                                                                                                                                                                                                        ";
                script += "\n\"\"\"                                                                                                                                                                                                      ";
                script += "\nquery_servers = \"\"\"                                                                                                                                                                                    ";
                script += "\nquery Embedded                                                                                                                                                                                           ";
                script += "\n{                                                                                                                                                                                                        ";
                script += "\n  databaseServers                                                                                                                                                                                        ";
                script += "\n  {                                                                                                                                                                                                      ";
                script += "\n    name                                                                                                                                                                                                 ";
                script += "\n  	hostName                                                                                                                                                                                              ";
                script += "\n    connectionType                                                                                                                                                                                       ";
                script += "\n    isEmbedded                                                                                                                                                                                           ";
                script += "\n    downstreamWorkbooks                                                                                                                                                                                  ";
                script += "\n    {                                                                                                                                                                                                    ";
                script += "\n      id                                                                                                                                                                                                 ";
                script += "\n      name                                                                                                                                                                                               ";
                script += "\n      upstreamDatabases                                                                                                                                                                                  ";
                script += "\n      {                                                                                                                                                                                                  ";
                script += "\n        name                                                                                                                                                                                             ";
                script += "\n        tables                                                                                                                                                                                           ";
                script += "\n        {                                                                                                                                                                                                ";
                script += "\n          fullName                                                                                                                                                                                       ";
                script += "\n          columns                                                                                                                                                                                        ";
                script += "\n          {                                                                                                                                                                                              ";
                script += "\n            name                                                                                                                                                                                         ";
                script += "\n            remoteType                                                                                                                                                                                   ";
                script += "\n            isNullable                                                                                                                                                                                   ";
                script += "\n          }                                                                                                                                                                                              ";
                script += "\n        }                                                                                                                                                                                                ";
                script += "\n      }                                                                                                                                                                                                  ";
                script += "\n    }                                                                                                                                                                                                    ";
                script += "\n  }                                                                                                                                                                                                      ";
                script += "\n}                                                                                                                                                                                                        ";
                script += "\n\"\"\"                                                                                                                                                                                                     ";
                script += "\nquery_refreshtime = \"\"\"                                                                                                                                                                                  ";
                script += "\nquery RefreshTime                                                                                                                                                                                        ";
                script += "\n{                                                                                                                                                                                                        ";
                script += "\n  datasources                                                                                                                                                                                            ";
                script += "\n  {                                                                                                                                                                                                      ";
                script += "\n    name                                                                                                                                                                                                 ";
                script += "\n    extractLastRefreshTime                                                                                                                                                                               ";
                script += "\n    createdAt                                                                                                                                                                                            ";
                script += "\n    updatedAt                                                                                                                                                                                            ";
                script += "\n    __typename                                                                                                                                                                                           ";
                script += "\n    downstreamWorkbooks                                                                                                                                                                                  ";
                script += "\n    {                                                                                                                                                                                                    ";
                script += "\n      id                                                                                                                                                                                                 ";
                script += "\n      name                                                                                                                                                                                               ";
                script += "\n    }                                                                                                                                                                                                    ";
                script += "\n  }                                                                                                                                                                                                      ";
                script += "\n}                                                                                                                                                                                                        ";
                script += "\n\"\"\"                                                                                                                                                                                                      ";
                script += "\nquery_files_2 = \"\"\"                                                                                                                                                                                  ";
                script += "\nquery files                                                                                                                                                                                              ";
                script += "\n{                                                                                                                                                                                                        ";
                script += "\ndatabases                                                                                                                                                                                                ";
                script += "\n  {                                                                                                                                                                                                      ";
                script += "\n    name                                                                                                                                                                                                 ";
                script += "\n    __typename                                                                                                                                                                                           ";
                script += "\n    downstreamWorkbooks                                                                                                                                                                                  ";
                script += "\n    {                                                                                                                                                                                                    ";
                script += "\n      name                                                                                                                                                                                               ";
                script += "\n    }                                                                                                                                                                                                    ";
                script += "\n    tables                                                                                                                                                                                               ";
                script += "\n    {                                                                                                                                                                                                    ";
                script += "\n      fullName                                                                                                                                                                                           ";
                script += "\n      columns                                                                                                                                                                                            ";
                script += "\n      {                                                                                                                                                                                                  ";
                script += "\n        name                                                                                                                                                                                             ";
                script += "\n        remoteType                                                                                                                                                                                       ";
                script += "\n        isNullable                                                                                                                                                                                       ";
                script += "\n      }                                                                                                                                                                                                  ";
                script += "\n    }                                                                                                                                                                                                    ";
                script += "\n  }                                                                                                                                                                                                      ";
                script += "\n}                                                                                                                                                                                                        ";
                script += "\n\"\"\"                                                                                                                                                                                                     ";
                script += "\nquery_calculations = \"\"\"                                                                                                                                                                                 ";
                script += "\nquery query_Calculations_info                                                                                                                                                                            ";
                script += "\n{                                                                                                                                                                                                        ";
                script += "\ncalculatedFields                                                                                                                                                                                         ";
                script += "\n{                                                                                                                                                                                                        ";
                script += "\ncalculatedfield_id: id                                                                                                                                                                                   ";
                script += "\nname                                                                                                                                                                                                     ";
                script += "\nformula                                                                                                                                                                                                  ";
                script += "\nrole                                                                                                                                                                                                     ";
                script += "\nisHidden                                                                                                                                                                                                 ";
                script += "\ndownstreamDashboards                                                                                                                                                                                     ";
                script += "\n{                                                                                                                                                                                                        ";
                script += "\nworkbook                                                                                                                                                                                                 ";
                script += "\n{                                                                                                                                                                                                        ";
                script += "\nname                                                                                                                                                                                                     ";
                script += "\n}                                                                                                                                                                                                        ";
                script += "\n}                                                                                                                                                                                                        ";
                script += "\n}                                                                                                                                                                                                        ";
                script += "\n}                                                                                                                                                                                                        ";
                script += "\n\"\"\"                                                                                                                                                                                                      ";
                script += "\nwb_query_results = conn.metadata_graphql_query(query_workbooks)                                                                                                                                          ";
                script += "\nwb_query_results_json = wb_query_results.json()                                                                                                                                                          ";
                script += "\ninner = wb_query_results_json['data']['workbooks']                                                                                                                                                       ";
                script += "\nresult = []                                                                                                                                                                                              ";
                script += "\nfor each_workbook in inner:                                                                                                                                                                              ";
                script += "\n    result.append([each_workbook['id'], each_workbook['name'], each_workbook['createdAt'], each_workbook['updatedAt'],                                                                                   ";
                script += "\n                   each_workbook['projectVizportalUrlId'], each_workbook['projectName'],                                                                                                                 ";
                script += "\n                   each_workbook['owner']['id'], each_workbook['owner']['name'],                                                                                                                         ";
                script += "\n                   each_workbook['site']['luid'], each_workbook['site']['name']])                                                                                                                        ";
                script += "\nresult_df = pd.DataFrame(result, columns=['Workbook ID', 'Workbook Name', 'Workbook Created At', 'Workbook updated At',                                                                                  ";
                script += "\n                                          'Project Vizportal Url ID', 'Project Name', 'Owner ID', 'Owner Name', 'Site LUID', 'Site Name'])                                                               ";
                script += "\nwb_query_results = conn.metadata_graphql_query(query_servers)                                                                                                                                            ";
                script += "\nwb_query_results_json = wb_query_results.json()                                                                                                                                                          ";
                script += "\ninner_1 = wb_query_results_json['data']['databaseServers']                                                                                                                                               ";
                script += "\nresult_1 = []                                                                                                                                                                                            ";
                script += "\nfor each in inner_1:                                                                                                                                                                                     ";
                script += "\n    for each_downstreamWorkbooks in each['downstreamWorkbooks']:                                                                                                                                         ";
                script += "\n        for each_upstreamDatabases in each_downstreamWorkbooks['upstreamDatabases']:                                                                                                                     ";
                script += "\n            for each_tables in each_upstreamDatabases['tables']:                                                                                                                                         ";
                script += "\n                for each_columns in each_tables['columns']:                                                                                                                                              ";
                script += "\n                    result_1.append([each['name'], each['hostName'] ,each['connectionType'], each['isEmbedded'],                                                                                         ";
                script += "\n                                  each_downstreamWorkbooks['id'], each_downstreamWorkbooks['name'],                                                                                                      ";
                script += "\n                                  each_upstreamDatabases['name'], each_tables['fullName'],                                                                                                               ";
                script += "\n                                  each_columns['name'], each_columns['remoteType'], each_columns['isNullable']])                                                                                         ";
                script += "\nresult_1_df = pd.DataFrame(result_1, columns =['databaseServersName', 'Server Name', 'Connection Type', 'Is Embedded Source',                                                                            ";
                script += "\n                                           'Workbook ID', 'Workbook Name',                                                                                                                               ";
                script += "\n                                           'Database Name', 'Table Name',                                                                                                                                ";
                script += "\n                                           'Column Name', 'Datatype', 'Is Null'])                                                                                                                        ";
                script += "\nwb_query_results = conn.metadata_graphql_query(query_refreshtime)                                                                                                                                        ";
                script += "\nwb_query_results_json = wb_query_results.json()                                                                                                                                                          ";
                script += "\ninner_4 = wb_query_results_json['data']['datasources']                                                                                                                                                   ";
                script += "\nresult_4 = []                                                                                                                                                                                            ";
                script += "\nfor each in inner_4:                                                                                                                                                                                     ";
                script += "\n    if each['downstreamWorkbooks'] != []:                                                                                                                                                                ";
                script += "\n        result_4.append([each['name'], each['extractLastRefreshTime'], each['createdAt'], each['updatedAt'], each['__typename'],                                                                         ";
                script += "\n                       each['downstreamWorkbooks'][0]['id'], each['downstreamWorkbooks'][0]['name']])                                                                                                    ";
                script += "\n    else:                                                                                                                                                                                                ";
                script += "\n        result_4.append([each['name'], each['extractLastRefreshTime'], each['createdAt'], each['updatedAt'], each['__typename'],                                                                         ";
                script += "\n                       '', ''])                                                                                                                                                                          ";
                script += "\nresult_4_df = pd.DataFrame(result_4, columns =['Datasource Name', 'Last Refresh Time', 'createdAt', 'updatedAt',                                                                                         ";
                script += "\n                                          'Connection Type', 'Workbook ID', 'Workbook Name'])                                                                                                            ";
                script += "\nwb_query_results = conn.metadata_graphql_query(query_files_2)                                                                                                                                            ";
                script += "\nwb_query_results_json = wb_query_results.json()                                                                                                                                                          ";
                script += "\ninner_5 = wb_query_results_json['data']['databases']                                                                                                                                                     ";
                script += "\nresult_5 = []                                                                                                                                                                                            ";
                script += "\nfor each in inner_5:                                                                                                                                                                                     ";
                script += "\n    inter = ''                                                                                                                                                                                           ";
                script += "\n    if each['downstreamWorkbooks'] != []:                                                                                                                                                                ";
                script += "\n        for each_downstreamWorkbooks in each['downstreamWorkbooks']:                                                                                                                                     ";
                script += "\n            inter = each_downstreamWorkbooks['name']                                                                                                                                                     ";
                script += "\n    for each_table in each['tables']:                                                                                                                                                                    ";
                script += "\n        for each_column in each_table['columns']:                                                                                                                                                        ";
                script += "\n            result_5.append([each['name'], each['__typename'], inter, each_table['fullName'],                                                                                                            ";
                script += "\n                           each_column['name'], each_column['remoteType'], each_column['isNullable']])                                                                                                   ";
                script += "\nresult_5_df = pd.DataFrame(result_5, columns =['File Name', 'Type', 'Workbook Name', 'Table Name',                                                                                                       ";
                script += "\n                                          'Column Name', 'Data Type', 'Is Null'])                                                                                                                        ";
                script += "\nresult_5_df=result_5_df.loc[result_5_df[\"Type\"] != 'DatabaseServer']                                                                                                                                     ";
                script += "\nwb_query_results = conn.metadata_graphql_query(query_calculations)                                                                                                                                       ";
                script += "\nwb_query_results_json = wb_query_results.json()                                                                                                                                                          ";
                script += "\ninner_6 = wb_query_results_json['data']['calculatedFields']                                                                                                                                              ";
                script += "\nresult_6 = []                                                                                                                                                                                            ";
                script += "\nfor each in inner_6:                                                                                                                                                                                     ";
                script += "\n    if each['downstreamDashboards'] != []:                                                                                                                                                               ";
                script += "\n        for each_downstreamDashboards in each['downstreamDashboards']:                                                                                                                                   ";
                script += "\n            result_6.append([each['calculatedfield_id'], each['name'], each['formula'], each['role'], each['isHidden'],                                                                                  ";
                script += "\n                          each_downstreamDashboards['workbook']['name']])                                                                                                                                ";
                script += "\n    else:                                                                                                                                                                                                ";
                script += "\n        result_6.append([each['calculatedfield_id'], each['name'], each['formula'], each['role'], each['isHidden'],                                                                                      ";
                script += "\n                          ''])                                                                                                                                                                           ";
                script += "\nresult_6_df = pd.DataFrame(result_6, columns =['calculatedfield_id', 'Calculated Fields Name', 'Formula', 'Role', 'Is Hidden',                                                                           ";
                script += "\n                                            'Workbook Name'])                                                                                                                                            ";
                script += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + server.ToString() + ";DATABASE=Tableau Metadata;Trusted_Connection=yes; \")";
                script += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))";
                //script += "\nquoted = urllib.parse.quote_plus("DRIVER={ SQL Server Native Client 11.0}; SERVER = IN3087262W1\SQLEXPRESS; DATABASE = Tableau Metadata; Trusted_Connection = yes; ")                                              ";
                script += "\nif result_df.empty:                                                                                            ";
                script += "\n    result_df.to_sql('TableauWorkbooks', schema='dbo',if_exists = 'append', con = engine, index=False)         ";
                script += "\nelse:                                                                                                          ";
                script += "\n    result_df.to_sql('TableauWorkbooks', schema='dbo',if_exists = 'append', con = engine)                      ";
                script += "\nif result_1_df.empty:                                                                                          ";
                script += "\n    result_1_df.to_sql('TableauDatabaseServers', schema='dbo',if_exists = 'append', con = engine, index=False) ";
                script += "\nelse:                                                                                                          ";
                script += "\n    result_1_df.to_sql('TableauDatabaseServers', schema='dbo',if_exists = 'append', con = engine)              ";
                script += "\nif result_5_df.empty:                                                                                          ";
                script += "\n    result_5_df.to_sql('TableauFileSources', schema='dbo',if_exists = 'append', con = engine, index=False)     ";
                script += "\nelse:                                                                                                          ";
                script += "\n    result_5_df.to_sql('TableauFileSources', schema='dbo',if_exists = 'append', con = engine)                  ";
                script += "\nif result_4_df.empty:                                                                                          ";
                script += "\n    result_4_df.to_sql('TableauRefreshTime', schema='dbo',if_exists = 'append', con = engine, index=False)     ";
                script += "\nelse:                                                                                                          ";
                script += "\n    result_4_df.to_sql('TableauRefreshTime', schema='dbo',if_exists = 'append', con = engine)                  ";
                script += "\nif result_6_df.empty:                                                                                          ";
                script += "\n    result_6_df.to_sql('TableauCalculations', schema='dbo',if_exists = 'append', con = engine, index=False)    ";
                script += "\nelse:                                                                                                          ";
                script += "\n    result_6_df.to_sql('TableauCalculations', schema='dbo',if_exists = 'append', con = engine)                 "; 
                script += "\nconn_str = (\"DRIVER={SQL Server Native Client 11.0};SERVER=" + server.ToString() + ";DATABASE=Tableau Metadata;Trusted_Connection=yes; \")                                                          ";
                script += "\ncnxn = pyodbc.connect(conn_str)                                                                                                                                                                          ";
                script += "\ncursor = cnxn.cursor()                                                                                                                                                                                   ";
                script += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted), fast_executemany=True)";
                script += "\nresult_df = pd.read_sql('select * from TableauWorkbooks', cnxn)                                                                                                                                          ";
                script += "\nresult_df = result_df.drop_duplicates()                                                                                                                                                                  ";
                script += "\nresult_df.to_sql('TableauWorkbooks', schema='dbo', if_exists = 'replace', con = engine, index=False)                                                                                                     ";
                script += "\nresult_1_df = pd.read_sql('select * from TableauDatabaseServers', cnxn)                                                                                                                                  ";
                script += "\nresult_1_df = result_1_df.drop_duplicates()                                                                                                                                                              ";
                script += "\nresult_1_df.to_sql('TableauDatabaseServers', schema='dbo', if_exists = 'replace', con = engine, index=False)                                                                                             ";
                script += "\nresult_5_df = pd.read_sql('select * from TableauFileSources', cnxn)                                                                                                                                      ";
                script += "\nresult_5_df = result_5_df.drop_duplicates()                                                                                                                                                              ";
                script += "\nresult_5_df.to_sql('TableauFileSources', schema='dbo', if_exists = 'replace', con = engine, index=False)                                                                                                 ";
                script += "\nresult_4_df = pd.read_sql('select * from TableauRefreshTime', cnxn)                                                                                                                                      ";
                script += "\nresult_4_df = result_4_df.drop_duplicates()                                                                                                                                                              ";
                script += "\nresult_4_df.to_sql('TableauRefreshTime', schema='dbo', if_exists = 'replace', con = engine, index=False)                                                                                                 ";
                script += "\nresult_6_df = pd.read_sql('select * from TableauCalculations', cnxn)                                                                                                                                     ";
                script += "\nresult_6_df = result_6_df.drop_duplicates()                                                                                                                                                              ";
                script += "\nresult_6_df.to_sql('TableauCalculations', schema='dbo', if_exists = 'replace', con = engine, index=False)                                                                                                ";
                script += "\nserver_wb = result_1_df['Workbook Name'].unique().tolist()                                                                                                                                               ";
                script += "\nfile_source_wb = result_5_df['Workbook Name'].unique().tolist()                                                                                                                                          ";
                script += "\nserver_wb_result = []                                                                                                                                                                                    ";
                script += "\nfor each_server_wb in server_wb:                                                                                                                                                                         ";
                script += "\n    database_name = []                                                                                                                                                                                   ";
                script += "\n    table_name = []                                                                                                                                                                                      ";
                script += "\n    column_name = []                                                                                                                                                                                     ";
                script += "\n    temp_data = result_1_df[result_1_df['Workbook Name'] == each_server_wb]                                                                                                                              ";
                script += "\n    for _, each_row in temp_data.iterrows():                                                                                                                                                             ";
                script += "\n        database_name.append(each_row['Database Name'])                                                                                                                                                  ";
                script += "\n        table_name.append(each_row['Table Name'])                                                                                                                                                        ";
                script += "\n        column_name.append(each_row['Column Name'])                                                                                                                                                      ";
                script += "\n    database_name = [i for i in database_name if i is not None]                                                                                                                                          ";
                script += "\n    table_name = [i for i in table_name if i is not None]                                                                                                                                                ";
                script += "\n    column_name = [i for i in column_name if i is not None]                                                                                                                                              ";
                script += "\n    server_wb_result.append([each_server_wb, database_name, table_name, column_name])                                                                                                                    ";
                script += "\nserver_wb_percentage_table = []                                                                                                                                                                          ";
                script += "\nfor i in range(0, len(server_wb_result)):                                                                                                                                                                ";
                script += "\n    for j in range(i+1, len(server_wb_result)):                                                                                                                                                          ";
                script += "\n        database_name_per = 0                                                                                                                                                                            ";
                script += "\n        table_name_per = 0                                                                                                                                                                               ";
                script += "\n        column_name_per = 0                                                                                                                                                                              ";
                script += "\n        if server_wb_result[i][1] == [] and server_wb_result[j][1] == []:                                                                                                                                ";
                script += "\n            database_name_per = np.nan                                                                                                                                                                   ";
                script += "\n        elif server_wb_result[i][1] != [] or server_wb_result[j][1] != []:                                                                                                                               ";
                script += "\n            database_name_per = len(set(server_wb_result[i][1]).intersection(set(server_wb_result[j][1]))) / float(len(set(server_wb_result[i][1] + server_wb_result[j][1]))) * 100                      ";
                script += "\n        if server_wb_result[i][2] == [] and server_wb_result[j][2] == []:                                                                                                                                ";
                script += "\n            table_name_per = np.nan                                                                                                                                                                      ";
                script += "\n        elif server_wb_result[i][2] != [] or server_wb_result[j][2] != []:                                                                                                                               ";
                script += "\n            table_name_per = len(set(server_wb_result[i][2]).intersection(set(server_wb_result[j][2]))) / float(len(set(server_wb_result[i][2] + server_wb_result[j][2]))) * 100                         ";
                script += "\n        if server_wb_result[i][3] == [] and server_wb_result[j][3] == []:                                                                                                                                ";
                script += "\n            column_name_per = np.nan                                                                                                                                                                     ";
                script += "\n        elif server_wb_result[i][3] != [] or server_wb_result[j][3] != []:                                                                                                                               ";
                script += "\n            column_name_per = len(set(server_wb_result[i][3]).intersection(set(server_wb_result[j][3]))) / float(len(set(server_wb_result[i][3] + server_wb_result[j][3]))) * 100                        ";
                script += "\n        server_wb_percentage_table.append([server_wb_result[i][0], server_wb_result[j][0], database_name_per, table_name_per, column_name_per])                                                          ";
                script += "\nserver_wb_percentage_table_df = pd.DataFrame(server_wb_percentage_table, columns=['Report A', 'Report B', 'Database/ File Name', 'Table Name', 'Column Name'])                                           ";
                script += "\nfile_source_wb_result = []                                                                                                                                                                               ";
                script += "\nfor each_file_source_wb in file_source_wb:                                                                                                                                                               ";
                script += "\n    file_name = []                                                                                                                                                                                       ";
                script += "\n    table_name = []                                                                                                                                                                                      ";
                script += "\n    column_name = []                                                                                                                                                                                     ";
                script += "\n    if each_file_source_wb != '':                                                                                                                                                                        ";
                script += "\n        temp_data = result_5_df[result_5_df['Workbook Name'] == each_file_source_wb]                                                                                                                     ";
                script += "\n        for _, each_row in temp_data.iterrows():                                                                                                                                                         ";
                script += "\n            file_name.append(each_row['File Name'])                                                                                                                                                      ";
                script += "\n            table_name.append(each_row['Table Name'])                                                                                                                                                    ";
                script += "\n            column_name.append(each_row['Column Name'])                                                                                                                                                  ";
                script += "\n        file_name = [i for i in file_name if i is not None]                                                                                                                                              ";
                script += "\n        table_name = [i for i in table_name if i is not None]                                                                                                                                            ";
                script += "\n        column_name = [i for i in column_name if i is not None]                                                                                                                                          ";
                script += "\n        file_source_wb_result.append([each_file_source_wb, file_name, table_name, column_name])                                                                                                          ";
                script += "\nfile_source_wb_percentage_table = []                                                                                                                                                                     ";
                script += "\nfor i in range(0, len(file_source_wb_result)):                                                                                                                                                           ";
                script += "\n    for j in range(i+1, len(file_source_wb_result)):                                                                                                                                                     ";
                script += "\n        file_name_per = 0                                                                                                                                                                                ";
                script += "\n        table_name_per = 0                                                                                                                                                                               ";
                script += "\n        column_name_per = 0                                                                                                                                                                              ";
                script += "\n        if file_source_wb_result[i][1] == [] and file_source_wb_result[j][1] == []:                                                                                                                      ";
                script += "\n            file_name_per = np.nan                                                                                                                                                                       ";
                script += "\n        elif file_source_wb_result[i][1] != [] or file_source_wb_result[j][1] != []:                                                                                                                     ";
                script += "\n            file_name_per = len(set(file_source_wb_result[i][1]).intersection(set(file_source_wb_result[j][1]))) / float(len(set(file_source_wb_result[i][1] + file_source_wb_result[j][1]))) * 100      ";
                script += "\n        if file_source_wb_result[i][2] == [] and file_source_wb_result[j][2] == []:                                                                                                                      ";
                script += "\n            table_name_per = np.nan                                                                                                                                                                      ";
                script += "\n        elif file_source_wb_result[i][2] != [] or file_source_wb_result[j][2] != []:                                                                                                                     ";
                script += "\n            table_name_per = len(set(file_source_wb_result[i][2]).intersection(set(file_source_wb_result[j][2]))) / float(len(set(file_source_wb_result[i][2] + file_source_wb_result[j][2]))) * 100     ";
                script += "\n        if file_source_wb_result[i][3] == [] and file_source_wb_result[j][3] == []:                                                                                                                      ";
                script += "\n            column_name_per = np.nan                                                                                                                                                                     ";
                script += "\n        elif file_source_wb_result[i][3] != [] or file_source_wb_result[j][3] != []:                                                                                                                     ";
                script += "\n            column_name_per = len(set(file_source_wb_result[i][3]).intersection(set(file_source_wb_result[j][3]))) / float(len(set(file_source_wb_result[i][3] + file_source_wb_result[j][3]))) * 100    ";
                script += "\n        file_source_wb_percentage_table.append([file_source_wb_result[i][0], file_source_wb_result[j][0], file_name_per, table_name_per, column_name_per])                                               ";
                script += "\nfile_source_wb_percentage_table_df = pd.DataFrame(file_source_wb_percentage_table, columns=['Report A', 'Report B', 'Database/ File Name', 'Table Name', 'Column Name'])                                 ";
                script += "\ncross_percentage_table = []                                                                                                                                                                              ";
                script += "\nfor i in range(0, len(server_wb_result)):                                                                                                                                                                ";
                script += "\n    for j in range(0, len(file_source_wb_result)):                                                                                                                                                       ";
                script += "\n        column_name_per = 0                                                                                                                                                                              ";
                script += "\n        if server_wb_result[i][3] == [] and file_source_wb_result[j][3] == []:                                                                                                                           ";
                script += "\n            column_name_per = np.nan                                                                                                                                                                     ";
                script += "\n        elif server_wb_result[i][3] != [] or file_source_wb_result[j][3] != []:                                                                                                                          ";
                script += "\n            column_name_per = len(set(server_wb_result[i][3]).intersection(set(file_source_wb_result[j][3]))) / float(len(set(server_wb_result[i][3] + file_source_wb_result[j][3]))) * 100              ";
                script += "\n        cross_percentage_table.append([server_wb_result[i][0], file_source_wb_result[j][0], np.nan, np.nan, column_name_per])                                                                            ";
                script += "\ncross_percentage_table_df = pd.DataFrame(cross_percentage_table, columns=['Report A', 'Report B', 'Database/ File Name', 'Table Name', 'Column Name'])                                                   ";
                script += "\npercentage_table_df = pd.concat([server_wb_percentage_table_df, file_source_wb_percentage_table_df, cross_percentage_table_df])                                                                          ";
                script += "\npercentage_table_df.reset_index(inplace = True, drop = True)                                                                                                                                             ";
                script += "\nif percentage_table_df.empty:                                                                                                    ";
                script += "\n    percentage_table_df.to_sql('Tableau_report_percentage_match', schema='dbo',if_exists = 'replace', con = engine, index=False) ";
                script += "\nelse:                                                                                                                            ";
                script += "\n    percentage_table_df.to_sql('Tableau_report_percentage_match', schema='dbo',if_exists = 'replace', con = engine)              ";
                


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

            //File.SetAttributes(path, FileAttributes.Normal);

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
             //createsqltableUsage();
                run_cmd();

                
            
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



        private async void button1_Click(object sender, RoutedEventArgs e)
        {
            
            


            

                string fileName = "BI4BI - Tableau.pbix";
                string path1 = System.IO.Path.Combine(Environment.CurrentDirectory, @"Report\", fileName);
                Process.Start(path1);


                

            
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
                        sw.WriteLine("python " + '"' + workingDirectory + @"\Tableau_Python.py" + '"');
                      //  sw.WriteLine("python Tableau_Python.py");
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

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            PasswordShow.Text = Password.Password;
            BorderPasswordShow.Visibility = Visibility.Visible;
            BorderPassword.Visibility = Visibility.Collapsed;
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            Password.Password = PasswordShow.Text;
            BorderPasswordShow.Visibility = Visibility.Collapsed;
            BorderPassword.Visibility = Visibility.Visible;
        }
        public async void createsqltableUsage()
        {


            try
            {
                string connectionString = @"Data Source = " + server.ToString().Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=Tableau Metadata";
                SqlConnection sqlconnection = new SqlConnection(connectionString);
                sqlconnection.Open();
                string strconnection = "Data Source = " + server.ToString() + "; Integrated Security=true; Initial Catalog=Tableau Metadata";
                string table = "";
                table += " TRUNCATE TABLE TableauWorkbooks ";
                table += " TRUNCATE TABLE TableauFileSources ";
                table += " TRUNCATE TABLE TableauDatabaseServers ";
                table += " TRUNCATE TABLE TableauRefreshTime ";
                table += " TRUNCATE TABLE TableauCalculations ";
                InsertQuery1(table, strconnection);
               
            }
            catch
            {
                MessageBox.Show("Please check the SQL server Instance and try again");
            }


        }

        public void createsqlDatabase()
        {
            string connectionString = @"Data Source = " + server.ToString().Replace("\\\\", "\\") + "; Integrated Security=true";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string strconnection = "Data Source = " + server.ToString() + "; Integrated Security=true";

            string table = "IF NOT EXISTS(SELECT name FROM master.dbo.sysdatabases WHERE Name='Tableau Metadata') CREATE DATABASE [Tableau Metadata]";

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
            catch(Exception ex)
            {
                MessageBox.Show("Please check the SQL server Instance and try again");
            }
        }

        private void GenerateMetadata_Checked(object sender, RoutedEventArgs e)
        {
            button1.Visibility = Visibility.Collapsed;
            Req.Visibility = Visibility.Collapsed;
            PDF.Visibility = Visibility.Collapsed;
            Generate_Metadata.Visibility = Visibility.Visible;
        }

        private void Output_Checked(object sender, RoutedEventArgs e)
        {
            button1.Visibility = Visibility.Visible;
            Req.Visibility = Visibility.Visible;
            PDF.Visibility = Visibility.Visible;
            Generate_Metadata.Visibility = Visibility.Collapsed;

        }

        //private void Req_Click(object sender, RoutedEventArgs e)
        //{

        //}
        
        



        private void Req_Click(object sender, RoutedEventArgs e)

        {

            int result = 0;



            string connectionstring = "Data Source=" + server.ToString() + "; Integrated Security=true; Initial Catalog=Tableau Metadata"; ; //your connectionstring    



            if (server.Equals(""))

            {

                MessageBox.Show("Click Load Data to populate the data base-> Then click the document generator");

            }

            else

            {

                using (SqlConnection conn = new SqlConnection(connectionstring))

                {

                    conn.Open();

                    SqlCommand cmd = new SqlCommand("select COUNT(*) from dbo.Tableau_report_percentage_match", conn);

                    result = (int)cmd.ExecuteScalar();

                    conn.Close();

                }

                if (server.ToString().Equals("") || result == 0)

                {

                    MessageBox.Show("Either the Metadata is not extracted or the SQL Server details is blank");

                }

                else

                {

                    Document_Generator_Tableau objWelcome = new Document_Generator_Tableau();

                    objWelcome.SQLTB_DGMS.Text = server;

                    objWelcome.Show(); //Sending value from one form to another form.

                    Close();

                }

            }

        }







        private void PDF_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ResultText_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void apiversion_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void USerName_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void SiteName_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void SiteURL_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void SQLServer_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void PythonPathText_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
    
}