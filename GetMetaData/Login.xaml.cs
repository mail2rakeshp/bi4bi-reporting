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


namespace GetMetaData
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        private System.Windows.Forms.NotifyIcon MyNotifyIcon;
        
        MainOptions window1 = new MainOptions();
        public Window1()
        {
            InitializeComponent();
            MyNotifyIcon = new System.Windows.Forms.NotifyIcon();
            MyNotifyIcon.Icon = new System.Drawing.Icon(
                            @"Final.ico");
            MyNotifyIcon.MouseDoubleClick +=
               new System.Windows.Forms.MouseEventHandler(MyNotifyIcon_MouseDoubleClick);
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
        private void Close_Click(object sender, RoutedEventArgs e)
        {
            // ShowInTaskbar = true;
            this.Close();
            //this.Hide();
            Application.Current.Shutdown();
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

    

        private void Get_Database_Click_1(object sender, RoutedEventArgs e)
        {
            if (UserName.Text.ToString().Equals("VICOE") &&( Password.Password.ToString().Equals("VICOETeam2022!") || PasswordShow.Text.ToString().Equals("VICOETeam2022!")))
            {

                UserName.Text = "";
                Password.Password = "";
                // window1.Show(); // Win10 tablet in tablet mode, use this, when sub Window is closed, the main window will be covered by the Start menu.
                
                this.Close();
                window1.ShowDialog();
            }
            else
            {
                MessageBox.Show("Invalid Credentials. Please Try again");
            }
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
        public void Window_Activated(object sender, EventArgs e)
        {
            this.Show();
        }
    }
}
