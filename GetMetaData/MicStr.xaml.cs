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
    public partial class MicStr : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;

        private static string PythonPath1;
        private static string TemplatePathString;
        private static string DestinationPathString;
        string pythout = "";
        int XMLCnt = 0;
        public MicStr()
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

        /*private void Maximize_Click(object sender, RoutedEventArgs e)
        {
            if (this.WindowState == WindowState.Maximized)
            {
                this.WindowState = WindowState.Normal;
            }
            else
            {
                this.WindowState = WindowState.Maximized;

            }
        }*/

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

        private void ProcessStart_Click(object sender, RoutedEventArgs e)
        {

        }



        private void GeneratePBI_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Generating Report. Please Wait ....");
            // string file = @"Metadata Output.pbix";
            string fileName = "BI4BI - MicroStrategy.pbix";
            string path = System.IO.Path.Combine(Environment.CurrentDirectory, @"Report\", fileName);
            Process.Start(path);
        }

        private void GenerateDoc_Click(object sender, RoutedEventArgs e)
        {
            int result = 0;

            string connectionstring = "Data Source=" + Source123.Text.ToString() + "; Integrated Security=true; Initial Catalog=Microstrategy Metadata"; ; //your connectionstring    

            if (Source123.Text.Equals(""))
            {
                MessageBox.Show("Click Load Data to populate the data base-> Then click the document generator");
            }
            else
            {
                using (SqlConnection conn = new SqlConnection(connectionstring))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("select COUNT(*) from dbo.mstr_object_component_list", conn);
                    result = (int)cmd.ExecuteScalar();
                    conn.Close();
                }
                if (Source123.Text.ToString().Equals("") || result == 0)
                {
                    MessageBox.Show("Either the Metadata is not extracted or the SQL Server details is blank");
                }
                else
                {
                    Document_Generator_Ms objWelcome = new Document_Generator_Ms();
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

        private void SignOutButton_Click(object sender, RoutedEventArgs e)
        {

            this.Close();
            Window1 window1 = new Window1();
            window1.ShowDialog();

        }

        private void InsertXML_Click(object sender, RoutedEventArgs e)
        {

            string path = Directory.GetCurrentDirectory() + @"\PythonFile\MicroStrategy_Process_Python.py";
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
                    )
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
                  //  GenerateMetadata.Visibility = Visibility.Collapsed;
                    GeneratePBI.Visibility = Visibility.Collapsed;
                    GenerateDoc.Visibility = Visibility.Collapsed;
                    InsertXML.Visibility = Visibility.Collapsed;
                    //  ProcessStart.Visibility = Visibility.Collapsed;
                    ProcessImage.Visibility = Visibility.Collapsed;
                    OutputImage.Visibility = Visibility.Collapsed;
                    DocImage.Visibility = Visibility.Collapsed;
                    Animation.Visibility = Visibility.Visible;



                      MessageBox.Show("Loading data into " + Source123.Text.ToString());
                      string script = "import pandas as pd";
                      script += "\nimport pyodbc";
                      script += "\nimport numpy as np";
                      script += "\nfrom sqlalchemy import create_engine";
                      script += "\nimport urllib";
                      script += "\nfrom pathlib import Path";
                      script += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=Microstrategy Metadata;Trusted_Connection=yes;\")";
                      script += "\nconn_str =(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=Microstrategy Metadata;Trusted_Connection=yes;\")";
                      script += "\ncnxn = pyodbc.connect(conn_str)";
                      script += "\ncursor = cnxn.cursor()";
                      script += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted), fast_executemany=True)";
                      script += "\ndirectory = r'" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "'";
                      script += "\nfiles = Path(directory).glob('*.csv')";

                      script += "\nfor file in files:";
                      script += "\n    df = pd.read_csv(file)";

                      script += "\n    sql_file_name = file.__str__().split('\\\\')[-1].removesuffix('.csv')";
                      script += "\n    df.to_sql(sql_file_name, schema = 'dbo', if_exists = 'replace', con = engine)";

                      
                    //Code to create flattened data structure.
                      script += "\nbase_data = pd.read_sql(\"select project_id, object_name, object_id, object_type_desc, component_object_id, component_object_name, component_object_type_desc from mstr_object_component_list\", cnxn)";
                      script += "\nbase_data = base_data[base_data['project_id'] == 7038510635751051264] # Restricting the data to single project for demo";
                      script += "\nl0_data = pd.read_sql(\"select project_id, object_name as l0_object_name, object_id as l0_object_id, object_type_desc as l0_object_type_desc,component_object_id as l1_object_id, component_object_name as l1_object_name, component_object_type_desc as l1_object_type_desc from mstr_object_component_list where object_type_desc in ('Grid Report', 'Document', 'Grid and Graph Report', 'SQL Report', 'Graph Report', 'Dossier', 'Managed Grid Report', 'Incremental Refresh Report', 'Transaction Services Report')\", cnxn)";
                      script += "\nl0_data = l0_data[l0_data['project_id'] == 7038510635751051264] # Restricting the data to single project for demo";
                      script += "\nbase_columns = ['project_id', 'object_name', 'object_id', 'object_type_desc', 'component_object_id', 'component_object_name', 'component_object_type_desc']";
                      script += "\ntable_list = []";
                      script += "\ntem_data = base_data[base_data['object_id'].isin(l0_data['l1_object_id'])][base_columns].copy()";
                      script += "\nlevel = 1";
                      script += "\nwhile tem_data.shape[0] != 0:";
                      script += "\n    updated_columns = ['project_id',";
                      script += "\n                       'l' + str(level) + '_object_name', ";
                      script += "\n                       'l' + str(level) + '_object_id',";
                      script += "\n                       'l' + str(level) + '_object_type_desc', ";
                      script += "\n                       'l' + str(level+1) + '_object_id',";
                      script += "\n                       'l' + str(level+1) + '_object_name',";
                      script += "\n                       'l' + str(level+1) + '_object_type_desc']";
                      script += "\n    tem_data.columns.values[:] = updated_columns";
                      script += "\n    table_list.append(tem_data)";
                      script += "\n    condition_l1_data = tem_data[tem_data['l' + str(level+1) + '_object_type_desc'] != 'Logical Table']";
                      script += "\n    condition_l1_data = condition_l1_data[condition_l1_data['l' + str(level+1) + '_object_type_desc'] != 'Database Table']";
                      script += "\n    condition_l1_data = condition_l1_data[condition_l1_data['l' + str(level+1) + '_object_type_desc'] != 'Fact']";
                      script += "\n    tem_data = base_data[base_data['object_id'].isin(condition_l1_data['l' + str(level+1) + '_object_id'])][base_columns]";
                      script += "\n    level = level+1";
                      script += "\ntable_list.insert(0, l0_data)";
                      script += "\nmerge = table_list[0]";
                      script += "\ncount = 0";
                      script += "\nfor i in range(0, len(table_list)-1):";
                      script += "\n    merge = pd.merge(merge, table_list[i+1], ";
                      script += "\n                          on = ['l' + str(i+1) + '_object_id', 'l' + str(i+1) + '_object_type_desc'],";
                      script += "\n                          how = 'left')";
                      script += "\n    merge = merge.drop(['l' + str(i+1) + '_object_name_y'], axis=1)";
                      script += "\n    merge.rename(columns = {'l' + str(i+1) + '_object_name_x':'l' + str(i+1) + '_object_name'}, inplace = True)";
                      script += "\n    if count == 0:";
                      script += "\n        merge = merge.drop(['project_id_y'], axis=1)";
                      script += "\n        merge.rename(columns = {'project_id_x':'project_id'}, inplace = True)";
                      script += "\nmerge = merge.dropna(axis=1, how='all')";
                      script += "\nsub_merge = merge.iloc[:,4:]";
                      script += "\nresult = []";
                      script += "\nfor i, row in sub_merge.iterrows():";
                      script += "\n    inner_result = []";
                      script += "\n    for x in range(0, len(sub_merge.columns), 3):";
                      script += "\n        if (row[x+2] is not np.nan) and (row[x+1] is not np.nan):";
                      script += "\n            inner_result.append([row[x+2], row[x+1]])";
                      script += "\n    result.append(inner_result)";
                      script += "\nmerge['all_columns'] = result";
                      script += "\nmerge_to_sql = merge.copy()";
                      script += "\nmerge_to_sql['all_columns'] = merge_to_sql['all_columns'].astype(str)";
                      script += "\nmerge_to_sql.to_sql('flattened_mstr_object_component_list', schema='dbo',if_exists = 'replace', con = engine)";
                      script += "\nsubset = merge[['project_id', 'l0_object_name', 'l0_object_id', 'all_columns']]";
                      script += "\ncomp_list = []";
                      script += "\nfor each_report in subset['l0_object_id'].unique():";
                      script += "\n    inner = []";
                      script += "\n    data_subset = subset[subset['l0_object_id'] == each_report]";
                      script += "\n    for i, row in data_subset.iterrows():";
                      script += "\n        inner = inner + row['all_columns']";
                      script += "\n    comp_list.append([data_subset['project_id'].unique()[0], ";
                      script += "\n                      data_subset['l0_object_name'].unique()[0], ";
                      script += "\n                      data_subset['l0_object_id'].unique()[0], ";
                      script += "\n                      inner])";
                      script += "\nunique_comp_list = []";
                      script += "\nfor each_element in comp_list:";
                      script += "\n    inner_unique = []";
                      script += "\n    for each_inner in each_element[3]:";
                      script += "\n        if each_inner not in inner_unique:";
                      script += "\n           inner_unique.append(each_inner)";
                      script += "\n    unique_comp_list.append([each_element[0], each_element[1], each_element[2], inner_unique])";
                      script += "\nupdated_unique_comp_list = []";
                      script += "\nfor each_row in unique_comp_list:";
                      script += "\n    rem_list = []";
                      script += "\n    for each_ele in each_row[3]:";
                      script += "\n        if each_ele[0] != 'Attribute Form Category':";
                      script += "\n            rem_list.append(each_ele)";
                      script += "\n    updated_unique_comp_list.append([each_row[0], each_row[1], each_row[2], rem_list])";
                      script += "\nindividual_unique_comp_list = []";
                      script += "\nfor each_row in unique_comp_list:";

                      script += "\n    dict_info = {}";
                      script += "\n    attribute = []";
                      script += "\n    fact = []";
                      script += "\n    column = []";
                      script += "\n    metric = []";

                      script += "\n    logical_table = []";
                      script += "\n    filters = []";
                      script += "\n    prompt = []";
                      script += "\n    for each_ele in each_row[3]:";
                      script += "\n        if each_ele[0] == 'Attribute':";
                      script += "\n            attribute.append(each_ele[1])";
                      script += "\n        elif each_ele[0] == 'Fact':";
                      script += "\n            fact.append(each_ele[1])";
                      script += "\n        elif each_ele[0] == 'Column':";
                      script += "\n            column.append(each_ele[1])";
                      script += "\n        elif each_ele[0] == 'Metric':";
                      script += "\n            metric.append(each_ele[1])";
                      script += "\n        elif each_ele[0] == 'Logical Table':";
                      script += "\n            logical_table.append(each_ele[1])";
                      script += "\n        elif each_ele[0] == 'Filter':";
                      script += "\n            filters.append(each_ele[1])";
                      script += "\n        elif each_ele[0] == 'Prompt':";
                      script += "\n            prompt.append(each_ele[1])";
                      script += "\n        else:";
                      script += "\n            pass";
                      script += "\n    dict_info['Attribute'] = attribute";
                      script += "\n    dict_info['Fact'] = fact";
                      script += "\n    dict_info['Column'] = column";
                      script += "\n    dict_info['Metric'] = metric";
                      script += "\n    dict_info['Logical Table'] = logical_table";
                      script += "\n    dict_info['Filter'] = filters";
                      script += "\n    dict_info['Prompt'] = prompt";
                      script += "\n    individual_unique_comp_list.append([each_row[0], each_row[1], each_row[2], each_row[3], dict_info])";
                      script += "\npercentage_table = []";
                      script += "\nfor i in range(0, len(individual_unique_comp_list)):";
                      script += "\n    attribute_per = 0.0";
                      script += "\n    fact_per = 0.0";
                      script += "\n    column_per = 0.0";
                      script += "\n    metric_per = 0.0";
                      script += "\n    logical_table_per = 0.0";
                      script += "\n    filters_per = 0.0";
                      script += "\n    prompt_per = 0.0";
                      script += "\n    report_a_attribute_count = 0";
                      script += "\n    report_b_attribute_count = 0";
                      script += "\n    report_a_fact_count = 0";
                      script += "\n    report_b_fact_count = 0";
                      script += "\n    report_a_column_count = 0";
                      script += "\n    report_b_column_count = 0";
                      script += "\n    report_a_metric_count = 0";
                      script += "\n    report_b_metric_count = 0";
                      script += "\n    report_a_logical_table_count = 0";
                      script += "\n    report_b_logical_table_count = 0";
                      script += "\n    report_a_filters_count = 0";
                      script += "\n    report_b_filters_count = 0";
                      script += "\n    report_a_prompt_count = 0";
                      script += "\n    report_b_prompt_count = 0";
                      script += "\n    for j in range(i+1, len(individual_unique_comp_list)):";
                      script += "\n        if individual_unique_comp_list[i][0] == individual_unique_comp_list[j][0]:";
                      script += "\n            report_a_attribute = individual_unique_comp_list[i][4]['Attribute']";
                      script += "\n            report_b_attribute = individual_unique_comp_list[j][4]['Attribute']";
                      script += "\n            report_a_fact = individual_unique_comp_list[i][4]['Fact']";
                      script += "\n            report_b_fact = individual_unique_comp_list[j][4]['Fact']";
                      script += "\n            report_a_column = individual_unique_comp_list[i][4]['Column']";
                      script += "\n            report_b_column = individual_unique_comp_list[j][4]['Column']";
                      script += "\n            report_a_metric = individual_unique_comp_list[i][4]['Metric']";
                      script += "\n            report_b_metric = individual_unique_comp_list[j][4]['Metric']";
                      script += "\n            report_a_logical_table = individual_unique_comp_list[i][4]['Logical Table']";
                      script += "\n            report_b_logical_table = individual_unique_comp_list[j][4]['Logical Table']";
                      script += "\n            report_a_filters = individual_unique_comp_list[i][4]['Filter']";
                      script += "\n            report_b_filters = individual_unique_comp_list[j][4]['Filter']";
                      script += "\n            report_a_prompt = individual_unique_comp_list[i][4]['Prompt']";
                      script += "\n            report_b_prompt = individual_unique_comp_list[j][4]['Prompt']";
                      script += "\n            report_a_attribute_count = len(set(report_a_attribute))";
                      script += "\n            report_b_attribute_count = len(set(report_b_attribute))";
                      script += "\n            report_a_fact_count = len(set(report_a_fact))";
                      script += "\n            report_b_fact_count = len(set(report_b_fact))";
                      script += "\n            report_a_column_count = len(set(report_a_column))";
                      script += "\n            report_b_column_count = len(set(report_b_column))";
                      script += "\n            report_a_metric_count = len(set(report_a_metric))";
                      script += "\n            report_b_metric_count = len(set(report_b_metric))";
                      script += "\n            report_a_logical_table_count = len(set(report_a_logical_table))";
                      script += "\n            report_b_logical_table_count = len(set(report_b_logical_table))";
                      script += "\n            report_a_filters_count = len(set(report_a_filters))";
                      script += "\n            report_b_filters_count = len(set(report_b_filters))";
                      script += "\n            report_a_prompt_count = len(set(report_a_prompt))";
                      script += "\n            report_b_prompt_count = len(set(report_b_prompt))";
                      script += "\n            if report_a_attribute == [] and report_b_attribute == []:";
                      script += "\n                attribute_per = np.nan";
                      script += "\n            elif report_a_attribute != [] or report_b_attribute != []:";
                      script += "\n                attribute_per = len(set(report_a_attribute).intersection(set(report_b_attribute))) / float(len(set(report_a_attribute + report_b_attribute))) * 100";
                      script += "\n            if report_a_fact == [] and report_b_fact == []:";
                      script += "\n                fact_per = np.nan";
                      script += "\n            elif report_a_fact != [] or report_b_fact != []:";
                      script += "\n                fact_per = len(set(report_a_fact).intersection(set(report_b_fact))) / float(len(set(report_a_fact + report_b_fact))) * 100";
                      script += "\n            if report_a_column == [] and report_b_column == []:";
                      script += "\n                column_per = np.nan";
                      script += "\n            elif report_a_column != [] or report_b_column != []:";
                      script += "\n                column_per = len(set(report_a_column).intersection(set(report_b_column))) / float(len(set(report_a_column + report_b_column))) * 100";
                      script += "\n            if report_a_metric == [] and report_b_metric == []:";
                      script += "\n                metric_per = np.nan";
                      script += "\n            elif report_a_metric != [] or report_b_metric != []:";
                      script += "\n                metric_per = len(set(report_a_metric).intersection(set(report_b_metric))) / float(len(set(report_a_metric + report_b_metric))) * 100";
                      script += "\n            if report_a_logical_table == [] and report_b_logical_table == []:";
                      script += "\n                logical_table_per = np.nan";
                      script += "\n            elif report_a_logical_table != [] or report_b_logical_table != []:";
                      script += "\n                logical_table_per = len(set(report_a_logical_table).intersection(set(report_b_logical_table))) / float(len(set(report_a_logical_table + report_b_logical_table))) * 100";
                      script += "\n            if report_a_filters == [] and report_b_filters == []:";
                      script += "\n                filters_per = np.nan";
                      script += "\n            elif report_a_filters != [] or report_b_filters != []:";
                      script += "\n                filters_per = len(set(report_a_filters).intersection(set(report_b_filters))) / float(len(set(report_a_filters + report_b_filters))) * 100";
                      script += "\n            if report_a_prompt == [] and report_b_prompt == []:";
                      script += "\n                prompt_per = np.nan";
                      script += "\n            elif report_a_prompt != [] or report_b_prompt != []:";
                      script += "\n                prompt_per = len(set(report_a_prompt).intersection(set(report_b_prompt))) / float(len(set(report_a_prompt + report_b_prompt))) * 100";
                      script += "\n            percentage_table.append([individual_unique_comp_list[i][1], individual_unique_comp_list[j][1],";
                      script += "\n                                     attribute_per, report_a_attribute_count, report_b_attribute_count,";
                      script += "\n                                     fact_per, report_a_fact_count, report_b_fact_count,";
                      script += "\n                                     column_per, report_a_column_count, report_b_column_count,";
                      script += "\n                                     metric_per, report_a_metric_count, report_b_metric_count,";
                      script += "\n                                     logical_table_per, report_a_logical_table_count, report_b_logical_table_count,";
                      script += "\n                                     filters_per, report_a_filters_count, report_b_filters_count,";
                      script += "\n                                     prompt_per, report_a_prompt_count, report_b_prompt_count])";
                      script += "\nreport_match = pd.DataFrame(percentage_table, columns=['Report A', 'Report B', ";
                      script += "\n                                                       'Attribute', 'Report A Attribute', 'Report B Attribute',";
                      script += "\n                                                       'Fact', 'Report A Fact', 'Report B Fact',";
                      script += "\n                                                       'Column', 'Report A Column', 'Report B Column',";
                      script += "\n                                                       'Metric', 'Report A Metric', 'Report B Metric',";
                      script += "\n                                                       'Logical Table', 'Report A Logical Table', 'Report B Logical Table',";
                      script += "\n                                                       'Filter', 'Report A Filter', 'Report B Filter',";
                      script += "\n                                                       'Prompt', 'Report A Prompt', 'Report B Prompt'])";
                      script += "\nreport_match.to_sql('report_match_percentage', schema='dbo',if_exists = 'replace', con = engine)";
                    
                    // 

                    /*MessageBox.Show("Loading data into " + Source123.Text.ToString());
                    string script = "import pandas as pd";
                    script += "\nimport pyodbc";
                    script += "\nfrom sqlalchemy import create_engine";
                    script += "\nimport urllib";
                    script += "\nfrom pathlib import Path";
                    script += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=Microstrategy Metadata;Trusted_Connection=yes;\")";
                    script += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))";
                    script += "\ndirectory = r'" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "'";
                    script += "\nfiles = Path(directory).glob('*.csv')";

                    script += "\nfor file in files:";
                    script += "\n    df = pd.read_csv(file)";

                    script += "\n    sql_file_name = file.__str__().split('\\\\')[-1].removesuffix('.csv')";
                    script += "\n    df.to_sql(sql_file_name, schema = 'dbo', if_exists = 'replace', con = engine)";
                    // */
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
                    try
                    {


                        createsqlDatabase();



                    }
                    catch (Exception ex)
                    {

                    }


                }


            }
            else if (AuthRad.IsChecked == true)
            {
                if (String.IsNullOrEmpty(TextCSV.Text.ToString()) || String.IsNullOrEmpty(Source123.Text.ToString()) || String.IsNullOrEmpty(TextPython.Text.ToString())
                 || String.IsNullOrEmpty(username.Text.ToString()) || String.IsNullOrEmpty(password.ToString()))
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
               //     GenerateMetadata.Visibility = Visibility.Collapsed;                  
                    GeneratePBI.Visibility = Visibility.Collapsed;
                    GenerateDoc.Visibility = Visibility.Collapsed;
                    InsertXML.Visibility = Visibility.Collapsed;
                    //   ProcessStart.Visibility = Visibility.Collapsed;
                    ProcessImage.Visibility = Visibility.Collapsed;
                    OutputImage.Visibility = Visibility.Collapsed;
                    DocImage.Visibility = Visibility.Collapsed;
                    Animation.Visibility = Visibility.Visible;



                    MessageBox.Show("Loading data into " + Source123.Text.ToString());
                    string script = "import pandas as pd";
                    script += "\nimport pyodbc";
                    script += "\nimport numpy as np";
                    script += "\nfrom sqlalchemy import create_engine";
                    script += "\nimport urllib";
                    script += "\nfrom pathlib import Path";
                    script += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=Microstrategy Metadata;Trusted_Connection=yes;\")";
                    script += "\nconn_str =(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=Microstrategy Metadata;Trusted_Connection=yes;\")";
                    script += "\ncnxn = pyodbc.connect(conn_str)";
                    script += "\ncursor = cnxn.cursor()";
                    script += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted), fast_executemany=True)";
                    script += "\ndirectory = r'" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "'";
                    script += "\nfiles = Path(directory).glob('*.csv')";

                    script += "\nfor file in files:";
                    script += "\n    df = pd.read_csv(file)";

                    script += "\n    sql_file_name = file.__str__().split('\\\\')[-1].removesuffix('.csv')";
                    script += "\n    df.to_sql(sql_file_name, schema = 'dbo', if_exists = 'replace', con = engine)";

                    
                    //Code to create flattened data structure.
                    script += "\nbase_data = pd.read_sql(\"select project_id, object_name, object_id, object_type_desc, component_object_id, component_object_name, component_object_type_desc from mstr_object_component_list\", cnxn)";
                    script += "\nbase_data = base_data[base_data['project_id'] == 7038510635751051264] # Restricting the data to single project for demo";
                    script += "\nl0_data = pd.read_sql(\"select project_id, object_name as l0_object_name, object_id as l0_object_id, object_type_desc as l0_object_type_desc,component_object_id as l1_object_id, component_object_name as l1_object_name, component_object_type_desc as l1_object_type_desc from mstr_object_component_list where object_type_desc in ('Grid Report', 'Document', 'Grid and Graph Report', 'SQL Report', 'Graph Report', 'Dossier', 'Managed Grid Report', 'Incremental Refresh Report', 'Transaction Services Report')\", cnxn)";
                    script += "\nl0_data = l0_data[l0_data['project_id'] == 7038510635751051264] # Restricting the data to single project for demo";
                    script += "\nbase_columns = ['project_id', 'object_name', 'object_id', 'object_type_desc', 'component_object_id', 'component_object_name', 'component_object_type_desc']";
                    script += "\ntable_list = []";
                    script += "\ntem_data = base_data[base_data['object_id'].isin(l0_data['l1_object_id'])][base_columns].copy()";
                    script += "\nlevel = 1";
                    script += "\nwhile tem_data.shape[0] != 0:";
                    script += "\n    updated_columns = ['project_id',";
                    script += "\n                       'l' + str(level) + '_object_name', ";
                    script += "\n                       'l' + str(level) + '_object_id',";
                    script += "\n                       'l' + str(level) + '_object_type_desc', ";
                    script += "\n                       'l' + str(level+1) + '_object_id',";
                    script += "\n                       'l' + str(level+1) + '_object_name',";
                    script += "\n                       'l' + str(level+1) + '_object_type_desc']";
                    script += "\n    tem_data.columns.values[:] = updated_columns";
                    script += "\n    table_list.append(tem_data)";
                    script += "\n    condition_l1_data = tem_data[tem_data['l' + str(level+1) + '_object_type_desc'] != 'Logical Table']";
                    script += "\n    condition_l1_data = condition_l1_data[condition_l1_data['l' + str(level+1) + '_object_type_desc'] != 'Database Table']";
                    script += "\n    condition_l1_data = condition_l1_data[condition_l1_data['l' + str(level+1) + '_object_type_desc'] != 'Fact']";
                    script += "\n    tem_data = base_data[base_data['object_id'].isin(condition_l1_data['l' + str(level+1) + '_object_id'])][base_columns]";
                    script += "\n    level = level+1";
                    script += "\ntable_list.insert(0, l0_data)";
                    script += "\nmerge = table_list[0]";
                    script += "\ncount = 0";
                    script += "\nfor i in range(0, len(table_list)-1):";
                    script += "\n    merge = pd.merge(merge, table_list[i+1], ";
                    script += "\n                          on = ['l' + str(i+1) + '_object_id', 'l' + str(i+1) + '_object_type_desc'],";
                    script += "\n                          how = 'left')";
                    script += "\n    merge = merge.drop(['l' + str(i+1) + '_object_name_y'], axis=1)";
                    script += "\n    merge.rename(columns = {'l' + str(i+1) + '_object_name_x':'l' + str(i+1) + '_object_name'}, inplace = True)";
                    script += "\n    if count == 0:";
                    script += "\n        merge = merge.drop(['project_id_y'], axis=1)";
                    script += "\n        merge.rename(columns = {'project_id_x':'project_id'}, inplace = True)";
                    script += "\nmerge = merge.dropna(axis=1, how='all')";
                    script += "\nsub_merge = merge.iloc[:,4:]";
                    script += "\nresult = []";
                    script += "\nfor i, row in sub_merge.iterrows():";
                    script += "\n    inner_result = []";
                    script += "\n    for x in range(0, len(sub_merge.columns), 3):";
                    script += "\n        if (row[x+2] is not np.nan) and (row[x+1] is not np.nan):";
                    script += "\n            inner_result.append([row[x+2], row[x+1]])";
                    script += "\n    result.append(inner_result)";
                    script += "\nmerge['all_columns'] = result";
                    script += "\nmerge_to_sql = merge.copy()";
                    script += "\nmerge_to_sql['all_columns'] = merge_to_sql['all_columns'].astype(str)";
                    script += "\nmerge_to_sql.to_sql('flattened_mstr_object_component_list', schema='dbo',if_exists = 'replace', con = engine)";
                    script += "\nsubset = merge[['project_id', 'l0_object_name', 'l0_object_id', 'all_columns']]";
                    script += "\ncomp_list = []";
                    script += "\nfor each_report in subset['l0_object_id'].unique():";
                    script += "\n    inner = []";
                    script += "\n    data_subset = subset[subset['l0_object_id'] == each_report]";
                    script += "\n    for i, row in data_subset.iterrows():";
                    script += "\n        inner = inner + row['all_columns']";
                    script += "\n    comp_list.append([data_subset['project_id'].unique()[0], ";
                    script += "\n                      data_subset['l0_object_name'].unique()[0], ";
                    script += "\n                      data_subset['l0_object_id'].unique()[0], ";
                    script += "\n                      inner])";
                    script += "\nunique_comp_list = []";
                    script += "\nfor each_element in comp_list:";
                    script += "\n    inner_unique = []";
                    script += "\n    for each_inner in each_element[3]:";
                    script += "\n        if each_inner not in inner_unique:";
                    script += "\n           inner_unique.append(each_inner)";
                    script += "\n    unique_comp_list.append([each_element[0], each_element[1], each_element[2], inner_unique])";
                    script += "\nupdated_unique_comp_list = []";
                    script += "\nfor each_row in unique_comp_list:";
                    script += "\n    rem_list = []";
                    script += "\n    for each_ele in each_row[3]:";
                    script += "\n        if each_ele[0] != 'Attribute Form Category':";
                    script += "\n            rem_list.append(each_ele)";
                    script += "\n    updated_unique_comp_list.append([each_row[0], each_row[1], each_row[2], rem_list])";
                    script += "\nindividual_unique_comp_list = []";
                    script += "\nfor each_row in unique_comp_list:";

                    script += "\n    dict_info = {}";
                    script += "\n    attribute = []";
                    script += "\n    fact = []";
                    script += "\n    column = []";
                    script += "\n    metric = []";

                    script += "\n    logical_table = []";
                    script += "\n    filters = []";
                    script += "\n    prompt = []";
                    script += "\n    for each_ele in each_row[3]:";
                    script += "\n        if each_ele[0] == 'Attribute':";
                    script += "\n            attribute.append(each_ele[1])";
                    script += "\n        elif each_ele[0] == 'Fact':";
                    script += "\n            fact.append(each_ele[1])";
                    script += "\n        elif each_ele[0] == 'Column':";
                    script += "\n            column.append(each_ele[1])";
                    script += "\n        elif each_ele[0] == 'Metric':";
                    script += "\n            metric.append(each_ele[1])";
                    script += "\n        elif each_ele[0] == 'Logical Table':";
                    script += "\n            logical_table.append(each_ele[1])";
                    script += "\n        elif each_ele[0] == 'Filter':";
                    script += "\n            filters.append(each_ele[1])";
                    script += "\n        elif each_ele[0] == 'Prompt':";
                    script += "\n            prompt.append(each_ele[1])";
                    script += "\n        else:";
                    script += "\n            pass";
                    script += "\n    dict_info['Attribute'] = attribute";
                    script += "\n    dict_info['Fact'] = fact";
                    script += "\n    dict_info['Column'] = column";
                    script += "\n    dict_info['Metric'] = metric";
                    script += "\n    dict_info['Logical Table'] = logical_table";
                    script += "\n    dict_info['Filter'] = filters";
                    script += "\n    dict_info['Prompt'] = prompt";
                    script += "\n    individual_unique_comp_list.append([each_row[0], each_row[1], each_row[2], each_row[3], dict_info])";
                    script += "\npercentage_table = []";
                    script += "\nfor i in range(0, len(individual_unique_comp_list)):";
                    script += "\n    attribute_per = 0.0";
                    script += "\n    fact_per = 0.0";
                    script += "\n    column_per = 0.0";
                    script += "\n    metric_per = 0.0";
                    script += "\n    logical_table_per = 0.0";
                    script += "\n    filters_per = 0.0";
                    script += "\n    prompt_per = 0.0";
                    script += "\n    report_a_attribute_count = 0";
                    script += "\n    report_b_attribute_count = 0";
                    script += "\n    report_a_fact_count = 0";
                    script += "\n    report_b_fact_count = 0";
                    script += "\n    report_a_column_count = 0";
                    script += "\n    report_b_column_count = 0";
                    script += "\n    report_a_metric_count = 0";
                    script += "\n    report_b_metric_count = 0";
                    script += "\n    report_a_logical_table_count = 0";
                    script += "\n    report_b_logical_table_count = 0";
                    script += "\n    report_a_filters_count = 0";
                    script += "\n    report_b_filters_count = 0";
                    script += "\n    report_a_prompt_count = 0";
                    script += "\n    report_b_prompt_count = 0";
                    script += "\n    for j in range(i+1, len(individual_unique_comp_list)):";
                    script += "\n        if individual_unique_comp_list[i][0] == individual_unique_comp_list[j][0]:";
                    script += "\n            report_a_attribute = individual_unique_comp_list[i][4]['Attribute']";
                    script += "\n            report_b_attribute = individual_unique_comp_list[j][4]['Attribute']";
                    script += "\n            report_a_fact = individual_unique_comp_list[i][4]['Fact']";
                    script += "\n            report_b_fact = individual_unique_comp_list[j][4]['Fact']";
                    script += "\n            report_a_column = individual_unique_comp_list[i][4]['Column']";
                    script += "\n            report_b_column = individual_unique_comp_list[j][4]['Column']";
                    script += "\n            report_a_metric = individual_unique_comp_list[i][4]['Metric']";
                    script += "\n            report_b_metric = individual_unique_comp_list[j][4]['Metric']";
                    script += "\n            report_a_logical_table = individual_unique_comp_list[i][4]['Logical Table']";
                    script += "\n            report_b_logical_table = individual_unique_comp_list[j][4]['Logical Table']";
                    script += "\n            report_a_filters = individual_unique_comp_list[i][4]['Filter']";
                    script += "\n            report_b_filters = individual_unique_comp_list[j][4]['Filter']";
                    script += "\n            report_a_prompt = individual_unique_comp_list[i][4]['Prompt']";
                    script += "\n            report_b_prompt = individual_unique_comp_list[j][4]['Prompt']";
                    script += "\n            report_a_attribute_count = len(set(report_a_attribute))";
                    script += "\n            report_b_attribute_count = len(set(report_b_attribute))";
                    script += "\n            report_a_fact_count = len(set(report_a_fact))";
                    script += "\n            report_b_fact_count = len(set(report_b_fact))";
                    script += "\n            report_a_column_count = len(set(report_a_column))";
                    script += "\n            report_b_column_count = len(set(report_b_column))";
                    script += "\n            report_a_metric_count = len(set(report_a_metric))";
                    script += "\n            report_b_metric_count = len(set(report_b_metric))";
                    script += "\n            report_a_logical_table_count = len(set(report_a_logical_table))";
                    script += "\n            report_b_logical_table_count = len(set(report_b_logical_table))";
                    script += "\n            report_a_filters_count = len(set(report_a_filters))";
                    script += "\n            report_b_filters_count = len(set(report_b_filters))";
                    script += "\n            report_a_prompt_count = len(set(report_a_prompt))";
                    script += "\n            report_b_prompt_count = len(set(report_b_prompt))";
                    script += "\n            if report_a_attribute == [] and report_b_attribute == []:";
                    script += "\n                attribute_per = np.nan";
                    script += "\n            elif report_a_attribute != [] or report_b_attribute != []:";
                    script += "\n                attribute_per = len(set(report_a_attribute).intersection(set(report_b_attribute))) / float(len(set(report_a_attribute + report_b_attribute))) * 100";
                    script += "\n            if report_a_fact == [] and report_b_fact == []:";
                    script += "\n                fact_per = np.nan";
                    script += "\n            elif report_a_fact != [] or report_b_fact != []:";
                    script += "\n                fact_per = len(set(report_a_fact).intersection(set(report_b_fact))) / float(len(set(report_a_fact + report_b_fact))) * 100";
                    script += "\n            if report_a_column == [] and report_b_column == []:";
                    script += "\n                column_per = np.nan";
                    script += "\n            elif report_a_column != [] or report_b_column != []:";
                    script += "\n                column_per = len(set(report_a_column).intersection(set(report_b_column))) / float(len(set(report_a_column + report_b_column))) * 100";
                    script += "\n            if report_a_metric == [] and report_b_metric == []:";
                    script += "\n                metric_per = np.nan";
                    script += "\n            elif report_a_metric != [] or report_b_metric != []:";
                    script += "\n                metric_per = len(set(report_a_metric).intersection(set(report_b_metric))) / float(len(set(report_a_metric + report_b_metric))) * 100";
                    script += "\n            if report_a_logical_table == [] and report_b_logical_table == []:";
                    script += "\n                logical_table_per = np.nan";
                    script += "\n            elif report_a_logical_table != [] or report_b_logical_table != []:";
                    script += "\n                logical_table_per = len(set(report_a_logical_table).intersection(set(report_b_logical_table))) / float(len(set(report_a_logical_table + report_b_logical_table))) * 100";
                    script += "\n            if report_a_filters == [] and report_b_filters == []:";
                    script += "\n                filters_per = np.nan";
                    script += "\n            elif report_a_filters != [] or report_b_filters != []:";
                    script += "\n                filters_per = len(set(report_a_filters).intersection(set(report_b_filters))) / float(len(set(report_a_filters + report_b_filters))) * 100";
                    script += "\n            if report_a_prompt == [] and report_b_prompt == []:";
                    script += "\n                prompt_per = np.nan";
                    script += "\n            elif report_a_prompt != [] or report_b_prompt != []:";
                    script += "\n                prompt_per = len(set(report_a_prompt).intersection(set(report_b_prompt))) / float(len(set(report_a_prompt + report_b_prompt))) * 100";
                    script += "\n            percentage_table.append([individual_unique_comp_list[i][1], individual_unique_comp_list[j][1],";
                    script += "\n                                     attribute_per, report_a_attribute_count, report_b_attribute_count,";
                    script += "\n                                     fact_per, report_a_fact_count, report_b_fact_count,";
                    script += "\n                                     column_per, report_a_column_count, report_b_column_count,";
                    script += "\n                                     metric_per, report_a_metric_count, report_b_metric_count,";
                    script += "\n                                     logical_table_per, report_a_logical_table_count, report_b_logical_table_count,";
                    script += "\n                                     filters_per, report_a_filters_count, report_b_filters_count,";
                    script += "\n                                     prompt_per, report_a_prompt_count, report_b_prompt_count])";
                    script += "\nreport_match = pd.DataFrame(percentage_table, columns=['Report A', 'Report B', ";
                    script += "\n                                                       'Attribute', 'Report A Attribute', 'Report B Attribute',";
                    script += "\n                                                       'Fact', 'Report A Fact', 'Report B Fact',";
                    script += "\n                                                       'Column', 'Report A Column', 'Report B Column',";
                    script += "\n                                                       'Metric', 'Report A Metric', 'Report B Metric',";
                    script += "\n                                                       'Logical Table', 'Report A Logical Table', 'Report B Logical Table',";
                    script += "\n                                                       'Filter', 'Report A Filter', 'Report B Filter',";
                    script += "\n                                                       'Prompt', 'Report A Prompt', 'Report B Prompt'])";
                    script += "\nreport_match.to_sql('report_match_percentage', schema='dbo',if_exists = 'replace', con = engine)";
                    


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
                    try
                    {
                        createsqlDatabase();
                        int XMLCnt = 10;

                        if (XMLCnt == 0)
                        {
                            Animation.Visibility = Visibility.Collapsed;
                            TemplatePath.Visibility = Visibility.Visible;
                            BorderTemplatePAth.Visibility = Visibility.Visible;
                            Template_Browse.Visibility = Visibility.Visible;
                            LabelSource.Visibility = Visibility.Visible;
                            BorderSource.Visibility = Visibility.Visible;
                            WindRad.Visibility = Visibility.Visible;
                            AuthRad.Visibility = Visibility.Collapsed;
                            LabelDatabaseServer.Visibility = Visibility.Visible;
                            BorderServer.Visibility = Visibility.Visible;
                            Browse.Visibility = Visibility.Visible;
                        //    GenerateMetadata.Visibility = Visibility.Visible;                           
                            //       ProcessStart.Visibility = Visibility.Collapsed;
                            InsertXML.Visibility = Visibility.Visible;
                            PasswordChek.Visibility = Visibility.Visible;
                            ProcessImage.Visibility = Visibility.Collapsed;
                            OutputImage.Visibility = Visibility.Collapsed;
                            DocImage.Visibility = Visibility.Collapsed;
                            Labelusername.Visibility = Visibility.Collapsed;
                            Borderusername.Visibility = Visibility.Collapsed;
                            Labelpasswd.Visibility = Visibility.Collapsed;
                            Borderpasswd.Visibility = Visibility.Collapsed;
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
                            // MessageBox.Show("Issues found in the details provided. Please contact the adminstrator in case of any support needed "); */
                        }
                        else
                        {
                            this.ShowInTaskbar = true;
                            MyNotifyIcon.Visible = true;
                            MyNotifyIcon.BalloonTipTitle = "Notification";
                            MyNotifyIcon.BalloonTipText = "XML's inserted Successfully. Hover over the Tip icons for more info";
                            MyNotifyIcon.ShowBalloonTip(5000);
                            //Thread.Sleep(5000);
                            MyNotifyIcon.Dispose();
                            //MessageBox.Show("XML's inserted Successfully. Hover over the Tip icons for more info");

                            MetadataToolTip.Text = "The XML's have been inserted into the server " + Source123.Text.ToString() + " and Database = Microstrategy Metadata";
                            MetadataToolTip.AppendText(Environment.NewLine);
                            MetadataToolTip.AppendText("Number of Distinct XML's processed=" + XMLCnt.ToString() + "\r\n");
                            MetadataToolTip.AppendText("Make sure the database has enough capacity to process the XML's ");
                            MetadataToolTip.AppendText(Environment.NewLine);
                            MetadataToolTip.AppendText(Environment.NewLine);
                            MetadataToolTip.AppendText("Tip : In case you are using your localhost server , the maximum limit of Daatbase is 10 GB");


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
                            WindRad.Visibility = Visibility.Collapsed;
                            AuthRad.Visibility = Visibility.Visible;
                            LabelDatabaseServer.Visibility = Visibility.Visible;
                            BorderServer.Visibility = Visibility.Visible;
                            Browse.Visibility = Visibility.Visible;
                       //     GenerateMetadata.Visibility = Visibility.Visible;                           
                            InsertXML.Visibility = Visibility.Visible;
                            GeneratePBI.Visibility = Visibility.Visible;
                            GenerateDoc.Visibility = Visibility.Visible;
                            //   ProcessStart.Visibility = Visibility.Visible;
                            ProcessImage.Visibility = Visibility.Collapsed;
                            OutputImage.Visibility = Visibility.Collapsed;
                            DocImage.Visibility = Visibility.Collapsed;
                           PasswordChek.Visibility = Visibility.Visible;
                            Labelusername.Visibility = Visibility.Visible;
                            Borderusername.Visibility = Visibility.Visible;
                            Labelpasswd.Visibility = Visibility.Visible;
                            Borderpasswd.Visibility = Visibility.Visible;



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
                      //  GenerateMetadata.Visibility = Visibility.Visible;
                        InsertXML.Visibility = Visibility.Visible;
                        GeneratePBI.Visibility = Visibility.Visible;
                        GenerateDoc.Visibility = Visibility.Visible;
                        // ProcessStart.Visibility = Visibility.Collapsed;
                        ProcessImage.Visibility = Visibility.Collapsed;
                        OutputImage.Visibility = Visibility.Collapsed;
                        DocImage.Visibility = Visibility.Collapsed;
                        PasswordChek.Visibility = Visibility.Visible;
                        Labelusername.Visibility = Visibility.Visible;
                        Borderusername.Visibility = Visibility.Visible;
                        Labelpasswd.Visibility = Visibility.Visible;
                        Borderpasswd.Visibility = Visibility.Visible;
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
            MessageBox.Show("Data loaded to the " + Source123.Text.ToString());



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
                        // Activate your environment
                        // sw.WriteLine("conda activate py3.9.7");
                        // run your script. You can also pass in arguments
                        sw.WriteLine("python MicroStrategy_Process_Python.py");
                    }
                }
                //string output = process.StandardOutput.ReadToEnd();

                process.WaitForExit();
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
              //  GenerateMetadata.Visibility = Visibility.Visible;               
                InsertXML.Visibility = Visibility.Visible;
                GeneratePBI.Visibility = Visibility.Visible;
                GenerateDoc.Visibility = Visibility.Visible;
                //   ProcessStart.Visibility = Visibility.Visible;
                ProcessImage.Visibility = Visibility.Collapsed;
                OutputImage.Visibility = Visibility.Collapsed;
                DocImage.Visibility = Visibility.Collapsed;
             //   PasswordChek.Visibility = Visibility.Visible;
                Labelusername.Visibility = Visibility.Collapsed;
                Borderusername.Visibility = Visibility.Collapsed;
                Labelpasswd.Visibility = Visibility.Collapsed;
                Borderpasswd.Visibility = Visibility.Collapsed;

                if (PasswordChek.IsChecked == true)
                {
                    Borderpasswd.Visibility = Visibility.Visible;
                }
                else
                {
                    BorderPasswordShow.Visibility = Visibility.Visible;
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }



        }

        private async void run_cmd_1()
        {

        }

        public void createsqlDatabase()
        {
            string connectionString = @"Data Source = " + Source123.Text.Replace("\\\\", "\\") + "; Integrated Security=true";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string strconnection = "Data Source = " + Source123.Text.ToString() + "; Integrated Security=true";

            string table = "IF NOT EXISTS(SELECT name FROM master.dbo.sysdatabases WHERE Name='Microstrategy Metadata') CREATE DATABASE[Microstrategy Metadata]";
            InsertQuery1(table, strconnection);
            run_cmd();


        }
        public async void createsqltableUsage()
        {


        }
        public async void createsqltableUsage1()
        {




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




