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
    /// Interaction logic for Cognos_CSV.xaml
    /// </summary>
    public partial class Cognos_CSV : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;

        private static string PythonPath1;
        private static string TemplatePathString;
        private static string DestinationPathString;
        string pythout = "";
        int XMLCnt=0;
        public Cognos_CSV()
        {
            InitializeComponent();
            ProcessStart.Visibility = Visibility.Collapsed;
            GeneratePBI.Visibility = Visibility.Collapsed;
            GenerateDoc.Visibility = Visibility.Collapsed;
            Labelusername.Visibility = Visibility.Collapsed;
            Borderusername.Visibility = Visibility.Collapsed;
            Labelpasswd.Visibility = Visibility.Collapsed;
            Borderpasswd.Visibility = Visibility.Collapsed;
            PasswordChek.Visibility = Visibility.Collapsed;
            InsertXML.Visibility = Visibility.Collapsed;
            ProcessImage.Visibility = Visibility.Collapsed;
            OutputImage.Visibility = Visibility.Collapsed;
            Output.IsEnabled = false;
            WindRad.IsChecked = true;

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
            ImageToolTip.AppendText("1. CSV Folder Path -  Directory where the XML extracts in CSV format needs to be loaded");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("2. SQL Server Details - Server where the XML's and relevant Metadata needs to be inserted (Windows or SQL Authentication) ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("3. Python Path - Directory where the python is installed in the system ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("4. Generate Metadata - Insert XML's and start the process for Metadata generation ");
            ImageToolTip.AppendText(Environment.NewLine);
            ImageToolTip.AppendText("5. Generate Output/Requirement Doc - generate output or requirement document based on the metadata inserted in Step 4 ");
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

        private void ProcessStart_Click(object sender, RoutedEventArgs e)
        {
            string path = Directory.GetCurrentDirectory() + @"\PythonFile\Cognos_Process_Python.py";
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
                GenerateMetadata.Visibility = Visibility.Collapsed;
                Output.Visibility = Visibility.Collapsed;
                GeneratePBI.Visibility = Visibility.Collapsed;
                GenerateDoc.Visibility = Visibility.Collapsed;
                InsertXML.Visibility = Visibility.Collapsed;
                ProcessStart.Visibility = Visibility.Collapsed;
                ProcessImage.Visibility = Visibility.Collapsed;
                OutputImage.Visibility = Visibility.Collapsed;

                Animation.Visibility = Visibility.Visible;

                SqlConnection SQLConnection = new SqlConnection();
                SQLConnection.ConnectionString = "Data Source =" + Source123.Text.ToString() + "; Initial Catalog =Cognos Metadata; " + "Integrated Security=true;";

                string QueryXML = "select count(*) from dbo.XMLData_Python";
                //Execute Queries and save results into variables
                SqlCommand CmdCntXML = SQLConnection.CreateCommand();
                CmdCntXML.CommandText = QueryXML;
                SQLConnection.Open();
                int XMLCnt = (Int32)CmdCntXML.ExecuteScalar();
                SQLConnection.Close();

                MessageBox.Show("Metadata Generation Proess started. Wait time depends on the number of XMLs being processed.\r\n"+ 
                    "You can go to the desktop by pressing Windows+D and the process will run in the backgrounds.\r\n" 
                    + " XML's Processing in the current process = " + XMLCnt.ToString()+ "\r\n"
                    +" Tip : For 25,000 XML's the wait time is 180 minutes");

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
                script += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER="+Source123.Text.ToString()+";DATABASE=Cognos Metadata;Trusted_Connection=yes;\")";
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
                script += "\n    quoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=Cognos Metadata;Trusted_Connection=yes;\")";
                script += "\n    engine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))";
                script += "\n    df.to_sql('CognosDataItems', schema='dbo',if_exists = 'append', con = engine)";
                script += "\n    df1.to_sql('CognosFiltersExpression', schema='dbo',if_exists = 'append', con = engine)";
                script += "\n    df2.to_sql('CognosReportVariables', schema='dbo',if_exists = 'append', con = engine)";
                script += "\n    df3.to_sql('CognosReportModificationTime', schema='dbo',if_exists = 'append', con = engine)    ";
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

                createsqltableUsage1();
                run_cmd_1();

                

                string QueryDI = "select count(*) from dbo.CognosDataItems";
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
                SqlCommand CmdCntRMT= SQLConnection.CreateCommand();
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



                if (DataITemCnt > 0 || FECnt>0 || RMTCnt>0 || RVCnt>0)
                {
                    
                    TemplatePath.Visibility = Visibility.Visible;
                    BorderTemplatePAth.Visibility = Visibility.Visible;
                    Template_Browse.Visibility = Visibility.Visible;
                    LabelSource.Visibility = Visibility.Visible;
                    BorderSource.Visibility = Visibility.Visible;
                    WindRad.Visibility = Visibility.Visible;
                    AuthRad.Visibility = Visibility.Visible;
                    Labelusername.Visibility = Visibility.Collapsed;
                    Borderusername.Visibility = Visibility.Collapsed;
                    Labelpasswd.Visibility = Visibility.Collapsed;
                    Borderpasswd.Visibility = Visibility.Collapsed;
                    BorderPasswordShow.Visibility = Visibility.Collapsed;
                    PasswordChek.Visibility = Visibility.Collapsed;
                    LabelDatabaseServer.Visibility = Visibility.Visible;
                    BorderServer.Visibility = Visibility.Visible;
                    Browse.Visibility = Visibility.Visible;
                    GenerateMetadata.Visibility = Visibility.Visible;
                    Output.Visibility = Visibility.Visible;
                    GeneratePBI.Visibility = Visibility.Collapsed;
                    GenerateDoc.Visibility = Visibility.Collapsed;
                    InsertXML.Visibility = Visibility.Collapsed;
                    ProcessStart.Visibility = Visibility.Collapsed;
                    ProcessImage.Visibility = Visibility.Visible;
                    OutputImage.Visibility = Visibility.Visible;
                    Animation.Visibility = Visibility.Collapsed;
                    Output.IsEnabled = true;
                    MetadataToolTip.Text = "Please Find the summary of items inserted into the server " + Source123.Text.ToString();
                    MetadataToolTip.AppendText(Environment.NewLine);
                    MetadataToolTip.AppendText ("Number of Dataitems = "+ DataITemCnt+"\r\n");
                    OutputToolTip.AppendText(Environment.NewLine);
                    MetadataToolTip.AppendText("Number of Filter Expressions = "+FECnt + "\r\n");
                    OutputToolTip.AppendText(Environment.NewLine);
                    MetadataToolTip.AppendText("Number of Report Variables = "+RVCnt + "\r\n");
                    OutputToolTip.AppendText(Environment.NewLine);
                    MetadataToolTip.AppendText("Number of Reports with Report Modification Time = "+RMTCnt);
                    //MessageBox.Show("Metadata is generated successfully. Click on Reset to start the process again.");
                    this.ShowInTaskbar = true;
                    MyNotifyIcon.Visible = true;
                    MyNotifyIcon.BalloonTipTitle = "Notification";
                    MyNotifyIcon.BalloonTipText = "Metadata is generated successfully. Click on Reset to start the process again.";
                    MyNotifyIcon.ShowBalloonTip(5000);
                    //Thread.Sleep(5000);
                    //MyNotifyIcon.Dispose();
                    
                }
                else
                {
                    
                    TemplatePath.Visibility = Visibility.Visible;
                    BorderTemplatePAth.Visibility = Visibility.Visible;
                    Template_Browse.Visibility = Visibility.Visible;
                    LabelSource.Visibility = Visibility.Visible;
                    BorderSource.Visibility = Visibility.Visible;
                    WindRad.Visibility = Visibility.Visible;
                    AuthRad.Visibility = Visibility.Visible;
                    Labelusername.Visibility = Visibility.Collapsed;
                    Borderusername.Visibility = Visibility.Collapsed;
                    Labelpasswd.Visibility = Visibility.Collapsed;
                    Borderpasswd.Visibility = Visibility.Collapsed;
                    BorderPasswordShow.Visibility = Visibility.Collapsed;
                    PasswordChek.Visibility = Visibility.Collapsed;
                    LabelDatabaseServer.Visibility = Visibility.Visible;
                    BorderServer.Visibility = Visibility.Visible;
                    Browse.Visibility = Visibility.Visible;
                    GenerateMetadata.Visibility = Visibility.Visible;
                    Output.Visibility = Visibility.Visible;
                    GeneratePBI.Visibility = Visibility.Collapsed;
                    GenerateDoc.Visibility = Visibility.Collapsed;
                    InsertXML.Visibility = Visibility.Collapsed;
                    ProcessStart.Visibility = Visibility.Collapsed;
                    ProcessImage.Visibility = Visibility.Visible;
                    OutputImage.Visibility = Visibility.Visible;
                    Animation.Visibility = Visibility.Collapsed;
                    //Output.IsEnabled = true;
                    MetadataToolTip.Text = "The XML's have been inserted into the server " + Source123.Text.ToString() + " and Database = Cognos Metadata";
                    MetadataToolTip.AppendText(Environment.NewLine);
                    MetadataToolTip.AppendText("Number of Distinct XML's processed=" + XMLCnt.ToString());
                    MetadataToolTip.AppendText("Make sure the database has enough capacity to process the XML's ");
                    MetadataToolTip.AppendText(Environment.NewLine);
                    MetadataToolTip.AppendText(Environment.NewLine);
                    MetadataToolTip.AppendText("Tip : In case you are using your localhost server , the maximum limit of Daatbase is 10 GB");

                    this.ShowInTaskbar = true;
                    MyNotifyIcon.Visible = true;
                    MyNotifyIcon.BalloonTipTitle = "Notification";
                    MyNotifyIcon.BalloonTipText = "Some issue in fetching the rows from the XML.Please check the XML or contact the adminstrator in case of any support needed";
                    MyNotifyIcon.ShowBalloonTip(5000);
                    //Thread.Sleep(5000);
                    //MyNotifyIcon.Dispose();
                   
                }


                
            }
            else if (AuthRad.IsChecked==true)
            {
                
            }


        }

        private void GeneratePBI_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Generating Report. Please Wait ....");
            // string file = @"Metadata Output.pbix";
            string fileName = "BI4BI - Cognos.pbix";
            string path = System.IO.Path.Combine(Environment.CurrentDirectory, @"Report\", fileName);
            Process.Start(path);
        }

        private void GenerateDoc_Click(object sender, RoutedEventArgs e)
        {
            int result = 0;

            string connectionstring = "Data Source=" + Source123.Text.ToString() + "; Integrated Security=true; Initial Catalog=Cognos Metadata"; ; //your connectionstring    

            using (SqlConnection conn = new SqlConnection(connectionstring))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("select COUNT(*) from dbo.CognosDataItems", conn);
                result = (int)cmd.ExecuteScalar();
                conn.Close();
            }
            if (Source123.Text.ToString().Equals("") || result == 0)
            {
                MessageBox.Show("Either the Metadata is not extracted or the SQL Server details is blank");
            }
            else
            {
                Document_Generator_Cognos objWelcome = new Document_Generator_Cognos();
                objWelcome.SQLTB.Text = Source123.Text;
                objWelcome.Show(); //Sending value from one form to another form.
                Close();
            }
        }
        private void GenerateMetadata_Checked(object sender, RoutedEventArgs e)
        {
            if (XMLCnt==0)
            {
                InsertXML.Visibility = Visibility.Visible;
                GeneratePBI.Visibility = Visibility.Collapsed;
                GenerateDoc.Visibility = Visibility.Collapsed;
                ProcessStart.Visibility = Visibility.Collapsed;
            }
            else
            {
                InsertXML.Visibility = Visibility.Collapsed;
                GeneratePBI.Visibility = Visibility.Collapsed;
                GenerateDoc.Visibility = Visibility.Collapsed;
                ProcessStart.Visibility = Visibility.Visible;
            }
        }

        private void Output_Checked(object sender, RoutedEventArgs e)
        {
            InsertXML.Visibility = Visibility.Collapsed;
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

        private void InsertXML_Click(object sender, RoutedEventArgs e)
        {
           
            string path = Directory.GetCurrentDirectory() + @"\PythonFile\Cognos_CSV_Python.py";
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
                    InsertXML.Visibility = Visibility.Collapsed;
                    ProcessStart.Visibility = Visibility.Collapsed;
                    ProcessImage.Visibility = Visibility.Collapsed;
                    OutputImage.Visibility = Visibility.Collapsed;

                    Animation.Visibility = Visibility.Visible;


                    MessageBox.Show("Inserting XMLs into the server "+Source123.Text.ToString());
                    string script = "import pandas as pd";
                    script+= "\nimport xml.etree.cElementTree as et";
                    script+= "\nimport pandas as pd";
                    script+= "\nfrom sqlalchemy import create_engine";
                    script+= "\nimport urllib";
                    script+= "\nimport glob";
                    script+= "\nfrom datetime import datetime";
                    script+= "\nfiles = glob.glob(\"" + TextCSV.Text.ToString().Replace("\\","\\\\")+ "\\\\*.csv\")";
                    script+= "\ndf = pd.DataFrame()";
                    script+= "\nfor f in files:";
                    script+= "\n    csv = pd.read_csv(f)";
                    script+= "\n    df = df.append(csv)";
                    script+= "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=Cognos Metadata;Trusted_Connection=yes;\")";
                    script+= "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))";
                    script+= "\ndf1=df[df.SPEC.str.startswith('<report')]";
                    script+= "\ndf2 =df1[df1.SPEC.str.endswith('</report>')]";
                    script+= "\ndf3 = df2['SPEC'].unique()";
                    script+= "\ndf3=pd.DataFrame(df3)";
                    script+= "\ndf3.columns=['XMLData']";
                    script+= "\ndf3.to_sql('XMLData_Python', schema='dbo',if_exists = 'append', con = engine, index=False)";
                    script += "\nprint(df3.shape[0])";
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
                        createsqltableUsage();
                        run_cmd();
                        SqlConnection SQLConnection = new SqlConnection();
                        SQLConnection.ConnectionString = "Data Source =" + Source123.Text.ToString() + "; Initial Catalog =Cognos Metadata; " + "Integrated Security=true;";

                        string QueryXML = "select count(*) from dbo.XMLData_Python";
                        //Execute Queries and save results into variables
                        SqlCommand CmdCntXML = SQLConnection.CreateCommand();
                        CmdCntXML.CommandText = QueryXML;
                        SQLConnection.Open();
                        int XMLCnt = (Int32)CmdCntXML.ExecuteScalar();
                        SQLConnection.Close();

                        if (XMLCnt==0)
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
                            InsertXML.Visibility = Visibility.Visible;
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
                        else {
                            //MessageBox.Show("XML's inserted Successfully. Hover over the Tip icons for more info");

                            this.ShowInTaskbar = true;
                            MyNotifyIcon.Visible = true;
                            MyNotifyIcon.BalloonTipTitle = "Notification";
                            MyNotifyIcon.BalloonTipText = "XML's inserted Successfully. Hover over the Tip icons for more info";
                            MyNotifyIcon.ShowBalloonTip(5000);
                            //Thread.Sleep(5000);
                            //MyNotifyIcon.Dispose();

                            MetadataToolTip.Text = "The XML's have been inserted into the server "+Source123.Text.ToString()+" and Database = Cognos Metadata";
                            MetadataToolTip.AppendText(Environment.NewLine);
                            MetadataToolTip.AppendText("Number of Distinct XML's processed=" + XMLCnt.ToString()+"\r\n");
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
                            WindRad.Visibility = Visibility.Visible;
                            AuthRad.Visibility = Visibility.Visible;
                            LabelDatabaseServer.Visibility = Visibility.Visible;
                            BorderServer.Visibility = Visibility.Visible;
                            Browse.Visibility = Visibility.Visible;
                            GenerateMetadata.Visibility = Visibility.Visible;
                            Output.Visibility = Visibility.Visible;
                            InsertXML.Visibility = Visibility.Collapsed;
                            ProcessStart.Visibility = Visibility.Visible;
                            ProcessImage.Visibility = Visibility.Visible;
                            OutputImage.Visibility = Visibility.Visible;

                            

                            TextCSV.IsEnabled = false;
                            Source123.IsEnabled = false;
                            TextPython.IsEnabled = false;
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
                        InsertXML.Visibility = Visibility.Visible;
                        ProcessStart.Visibility = Visibility.Collapsed;
                        ProcessImage.Visibility = Visibility.Collapsed;
                        OutputImage.Visibility = Visibility.Collapsed;
                        MessageBox.Show(ex.Message.ToString());
                        // MessageBox.Show("Issues found in the details provided. Please contact the adminstrator in case of any support needed");
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
                    InsertXML.Visibility = Visibility.Collapsed;
                    ProcessStart.Visibility = Visibility.Collapsed;
                    ProcessImage.Visibility = Visibility.Collapsed;
                    OutputImage.Visibility = Visibility.Collapsed;

                    Animation.Visibility = Visibility.Visible;


                    MessageBox.Show("Inserting XMLs into the server " + Source123.Text.ToString());
                    string script = "import pandas as pd";
                    script += "\nimport xml.etree.cElementTree as et";
                    script += "\nimport pandas as pd";
                    script += "\nfrom sqlalchemy import create_engine";
                    script += "\nimport urllib";
                    script += "\nimport glob";
                    script += "\nfrom datetime import datetime";
                    script += "\nfiles = glob.glob(\"" + TextCSV.Text.ToString().Replace("\\", "\\\\") + "\\\\*.csv\")";
                    script += "\ndf = pd.DataFrame()";
                    script += "\nfor f in files:";
                    script += "\n    csv = pd.read_csv(f)";
                    script += "\n    df = df.append(csv)";
                    script += "\nquoted = urllib.parse.quote_plus(\"DRIVER={SQL Server Native Client 11.0};SERVER=" + Source123.Text.ToString() + ";DATABASE=Cognos Metadata;UID=" + username.Text.ToString() + ";PWD=" + password + ";\")";
                    script += "\nengine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(quoted))";
                    script += "\ndf1=df[df.SPEC.str.startswith('<report')]";
                    script += "\ndf2 =df1[df1.SPEC.str.endswith('</report>')]";
                    script += "\ndf3 = df2['SPEC'].unique()";
                    script += "\ndf3=pd.DataFrame(df3)";
                    script += "\ndf3.columns=['XMLData']";
                    script += "\ndf3.to_sql('XMLData_Python', schema='dbo',if_exists = 'append', con = engine, index=False)";
                    script += "\nprint(df3.shape[0])";
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
                        createsqltableUsage();
                        run_cmd();

                        SqlConnection SQLConnection = new SqlConnection();
                        SQLConnection.ConnectionString = "Data Source =" + Source123.Text.ToString() + "; Initial Catalog =Cognos Metadata; " + "Integrated Security=true;";

                        string QueryXML = "select count(*) from dbo.XMLData_Python";
                        //Execute Queries and save results into variables
                        SqlCommand CmdCntXML = SQLConnection.CreateCommand();
                        CmdCntXML.CommandText = QueryXML;
                        SQLConnection.Open();
                        int XMLCnt = (Int32)CmdCntXML.ExecuteScalar();
                        SQLConnection.Close();

                        if (XMLCnt==0)
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
                            InsertXML.Visibility = Visibility.Visible;
                            PasswordChek.Visibility = Visibility.Visible;
                            ProcessImage.Visibility = Visibility.Collapsed;
                            OutputImage.Visibility = Visibility.Collapsed;
                            Labelusername.Visibility = Visibility.Collapsed;
                            Borderusername.Visibility = Visibility.Collapsed;
                            Labelpasswd.Visibility = Visibility.Collapsed;
                            if (PasswordChek.IsChecked ==true)
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
                            this.ShowInTaskbar = true;
                            MyNotifyIcon.Visible = true;
                            MyNotifyIcon.BalloonTipTitle = "Notification";
                            MyNotifyIcon.BalloonTipText = "XML's inserted Successfully. Hover over the Tip icons for more info";
                            MyNotifyIcon.ShowBalloonTip(5000);
                            //Thread.Sleep(5000);
                            MyNotifyIcon.Dispose();
                            //MessageBox.Show("XML's inserted Successfully. Hover over the Tip icons for more info");
                            
                            MetadataToolTip.Text = "The XML's have been inserted into the server " + Source123.Text.ToString() + " and Database = Cognos Metadata";
                            MetadataToolTip.AppendText(Environment.NewLine);
                            MetadataToolTip.AppendText("Number of Distinct XML's processed=" + XMLCnt.ToString()+"\r\n");
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
                            WindRad.Visibility = Visibility.Visible;
                            AuthRad.Visibility = Visibility.Visible;
                            LabelDatabaseServer.Visibility = Visibility.Visible;
                            BorderServer.Visibility = Visibility.Visible;
                            Browse.Visibility = Visibility.Visible;
                            GenerateMetadata.Visibility = Visibility.Visible;
                            Output.Visibility = Visibility.Visible;
                            InsertXML.Visibility = Visibility.Collapsed;
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
                        InsertXML.Visibility = Visibility.Visible;
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
                        sw.WriteLine("py Cognos_CSV_Python.py");
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
                            sw.WriteLine("py Cognos_Process_Python.py");
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
            string connectionString = @"Data Source = " + Source123.Text.Replace("\\\\", "\\") + "; Integrated Security=true";
            SqlConnection sqlconnection = new SqlConnection(connectionString);
            sqlconnection.Open();
            string strconnection = "Data Source = " + Source123.Text.ToString() + "; Integrated Security=true";

            string table = "IF NOT EXISTS(SELECT name FROM master.dbo.sysdatabases WHERE Name='Cognos Metadata') CREATE DATABASE[Cognos Metadata]";
            InsertQuery1(table, strconnection);
        }
        public async void createsqltableUsage()
        {


            try
            {
                string connectionString = @"Data Source = "+Source123.Text.ToString()+ "; Integrated Security=true; Initial Catalog=Cognos Metadata";
                SqlConnection sqlconnection = new SqlConnection(connectionString);
                sqlconnection.Open();
                string strconnection = @"Data Source  = " + Source123.Text.ToString() + "; Integrated Security=true; Initial Catalog=Cognos Metadata";
                string table = "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='XMLData_Python') BEGIN DROP TABLE XMLData_Python END";
                table += "\n CREATE TABLE [dbo].[XMLData_Python](";
                table += "\n 	[XMLData] [xml] NULL,";
                table += "\n 	[index] [int] identity(1,1),";
                table += "\n ) ";
                InsertQuery1(table, strconnection);

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                //MessageBox.Show("Please check the SQL server Instance and try again");
            }


        }
        public async void createsqltableUsage1()
        {


            try
            {
                string connectionString = @"Data Source = " + Source123.Text.Replace("\\\\", "\\") + "; Integrated Security=true; Initial Catalog=Cognos Metadata";
                SqlConnection sqlconnection = new SqlConnection(connectionString);
                sqlconnection.Open();
                string strconnection = "Data Source = " + Source123.Text.ToString() + "; Integrated Security=true; Initial Catalog=Cognos Metadata";
                string table = "";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='CognosDataItems') BEGIN DROP TABLE CognosDataItems END";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='CognosFiltersExpression') BEGIN DROP TABLE CognosFiltersExpression END  ";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='CognosReportVariables') BEGIN DROP TABLE CognosReportVariables END";
                table += " IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='CognosReportModificationTime') BEGIN DROP TABLE CognosReportModificationTime END ";
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
            ProcessStart.Visibility = Visibility.Collapsed;
            InsertXML.Visibility = Visibility.Visible;
            ProcessImage.Visibility = Visibility.Collapsed;
            OutputImage.Visibility = Visibility.Collapsed;


        }
    }
}




