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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Security.Permissions;
using IronPython;
using Microsoft.Scripting.Hosting;
using System.Diagnostics;
using System.IO;
using System.Data.SqlClient;
using System.Threading;
using GetMetadata;

namespace GetMetaData
{
    /// <summary>
    /// Interaction logic for MainOptions.xaml
    /// </summary>
    public partial class MainOptions : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;
        

        Tableau Tableau = new Tableau();
        PowerBi PowerBIDialog = new PowerBi();
        Qlikview Qlikview = new Qlikview();

        Cognos_Options cognos = new Cognos_Options();
        OBIEE ob = new OBIEE();
        MicStr mstr = new MicStr();
        
        

        // MainOptions  windowop = new MainOptions();
        public MainOptions()
        {
            InitializeComponent();
            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            MyNotifyIcon.Icon = new System.Drawing.Icon(
                            @"Final.ico");
            MyNotifyIcon.MouseDoubleClick +=
                new System.Windows.Forms.MouseEventHandler(MyNotifyIcon_MouseDoubleClick);

           
           MyNotifyIcon.BalloonTipClicked += new EventHandler(MyNotifyIcon_BalloonTipClicked);
            
            //MyNotifyIcon.BalloonTipShown += new EventHandler(MyNotifyIcon_BalloonTipClicked);

            //MyNotifyIcon.BalloonTipClosed += new EventHandler(MyNotifyIcon_BalloonTipClicked);

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
            MyNotifyIcon.Text = "";
            MyNotifyIcon.BalloonTipTitle = "Minimize Sucessful";
            MyNotifyIcon.BalloonTipText = "Minimized the app ";
            MyNotifyIcon.ShowBalloonTip(5000);
           // Thread.Sleep(5000);
           //MyNotifyIcon.Dispose();
            




            //ShowInTaskbar = true;

        }
        private void MyNotifyIcon_MouseDoubleClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
            this.WindowState = WindowState.Maximized;
            
        }
        void MyNotifyIcon_BalloonTipClicked(object sender, EventArgs e)
        {
            this.WindowState = WindowState.Minimized;
            this.WindowState = WindowState.Maximized;
           // MyNotifyIcon.Visible = false;
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

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            
            
            this.Close();
            cognos.ShowDialog();
        }

        private void button2_Copy_Click(object sender, RoutedEventArgs e)
        {
            /*
            this.Close();
            cognos.ShowDialog();*/
        }

        private void button2_Copy1_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            Qlikview.ShowDialog();
        }

        private void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            Window1 window1 = new Window1();
            window1.ShowDialog();

        }


        private void Tableau_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            Tableau.ShowDialog();
        }

        private void PowerBI_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            PowerBIDialog.ShowDialog();
        }

        private void SSRS_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Crystal_Click(object sender, RoutedEventArgs e)
        {

        }

        private void DBM_Click(object sender, RoutedEventArgs e)
        {

        }


        private void OBIEE_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            ob.ShowDialog();
        }

        private void mstr_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            mstr.ShowDialog();
        }
    }
}
