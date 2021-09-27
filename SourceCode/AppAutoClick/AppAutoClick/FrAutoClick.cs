using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static AppAutoClick.Helper.Win32;
using WindowScrape;
using WindowScrape.Types;
using AppAutoClick.Helper;

namespace AppAutoClick
{
    public partial class FrAutoClick : Form
    {
        private bool run = false;
        private long count = 0;
        private string softwareUsername = ConfigurationManager.AppSettings["SoftwareUsername"];
        private string softwarePassword = ConfigurationManager.AppSettings["SoftwarePassword"];
        private string pathFileExe = ConfigurationManager.AppSettings["PathFileExe"];
        private string pathCredential = ConfigurationManager.AppSettings["PathCredential"];
        private string spreadsheetId = ConfigurationManager.AppSettings["SpreadsheetId"];
        private string sheetName = ConfigurationManager.AppSettings["SheetName"];
        //private string nameWindowMain = "Calculator";
        private string nameWindowLogin = " Login to Pathogen Asset Control System";
        private string nameWindowReLogin = "PACS - inactivity lock";
        private string nameWindowRepositoryManagement = "Pathogen Asset Control System - [Repository Management]";

        private TimeSpan timeSleep;

        public FrAutoClick()
        {
            InitializeComponent();
            SetCountLabel(this.count);
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            StartAutoClicker();
        }
        private void btnStop_Click(object sender, EventArgs e)
        {
            this.run = false;
            this.count = 0;
            SetCountLabel(this.count);
            EnableControls();
        }

        private void StartAutoClicker()
        {
            var messageError = ValidData();
            if (string.IsNullOrEmpty(messageError))
            {
                timeSleep = new TimeSpan(int.Parse(txtHour.Text), int.Parse(txtMinute.Text), 0);
                if (!this.run)
                    AutoClickOnNewThread();
                this.run = true;
                DisableControls();
            }
            else
            {
                MessageBoxButtons errorButtons = MessageBoxButtons.OK;
                MessageBox.Show(messageError, "Error", errorButtons);
            }
        }
        
        private void DisableControls()
        {
            btnStart.Enabled = false;
            txtHour.Enabled = false;
            txtMinute.Enabled = false;
        }
        private void EnableControls()
        {
            btnStart.Enabled = true;
            txtHour.Enabled = true;
            txtMinute.Enabled = true;
        }

        private bool IsOpenSoftware(string classWindow, string nameWindow)
        {
            IntPtr hwnd = Win32.FindWindow(classWindow, nameWindow);
            return hwnd != IntPtr.Zero;
        }
        private string ValidData()
        {
            if (!File.Exists(pathFileExe))
            {
                return "Can't start, software path is incorrect";
            }
            var hourStr = txtHour.Text;
            var minuteStr = txtMinute.Text;
            if (string.IsNullOrEmpty(hourStr) || string.IsNullOrEmpty(minuteStr))
            {
                return "Hours and minutes cannot be left blank";
            }
            var hour = long.Parse(hourStr);
            var minute = long.Parse(minuteStr);
            if (hour == 0 && minute == 0)
            {
                return "Hours and minutes must have a value";
            }
            if (hour > 100 || minute > 59)
            {
                return "Hours must be less than 100 and minutes must be less than 60";
            }
            return string.Empty;
        }
        private void AutoClickOnNewThread()
        {
            Thread t = new Thread(AutoClick);
            t.IsBackground = true;
            t.Start();
        }
        private void AutoClick()
        {
            while (this.run)
            {
                try
                {
                    if (!IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowLogin) && 
                        !IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowReLogin) && 
                        !IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowRepositoryManagement) )
                    {
                        System.Diagnostics.Process.Start(pathFileExe);
                        while (IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowLogin))
                        {
                            Thread.Sleep(2000);
                        }
                    }

                    // login
                    if (IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowLogin))
                    {
                        IntPtr windowLogin = FindWindow(null, nameWindowLogin);

                        var inputLogins = EnumAllWindows(windowLogin, "WindowsForms10.EDIT.app.0.2bf8098_r6_ad1").ToList();
                        var inputUsername = inputLogins[2];
                        var inputPassword = inputLogins[0];

                        SendMessage(inputUsername, WM_SETTEXT, 0, softwareUsername);
                        Thread.Sleep(2000);
                        SendMessage(inputPassword, WM_SETTEXT, 0, softwarePassword);
                        Thread.Sleep(2000);


                        IntPtr btnOk = FindWindowEx(windowLogin, IntPtr.Zero, "WindowsForms10.Window.b.app.0.2bf8098_r6_ad1", "Ok");

                        SendMessage(btnOk, WM_LBUTTONDOWN, 0, IntPtr.Zero);
                        SendMessage(btnOk, WM_LBUTTONUP, 0, IntPtr.Zero);

                    }
                    else if (IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowReLogin))
                    {
                        IntPtr windowLogin = FindWindow(null, nameWindowReLogin);
                        var inputLogins = EnumAllWindows(windowLogin, "WindowsForms10.EDIT.app.0.2bf8098_r6_ad1").ToList();

                        var inputPassword = inputLogins[0];
                        SendMessage(inputPassword, WM_SETTEXT, 0, softwarePassword);
                        Thread.Sleep(2000);


                        IntPtr btnOk = FindWindowEx(windowLogin, IntPtr.Zero, "WindowsForms10.Window.b.app.0.2bf8098_r6_ad1", "Ok");

                        SendMessage(btnOk, WM_LBUTTONDOWN, 0, IntPtr.Zero);
                        SendMessage(btnOk, WM_LBUTTONUP, 0, IntPtr.Zero);
                    }

                    if (IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowRepositoryManagement))
                    {
                        IntPtr windowRepositoryManagement = FindWindow(null, nameWindowRepositoryManagement);

                        var pointRepositoryManagement = new POINT();
                        ClientToScreen(windowRepositoryManagement, ref pointRepositoryManagement);

                        LeftMouseClick(165, 47);
                    }



                    //SendMessage(btnCancel, MOUSEEVENTF_LEFTUP, 0, IntPtr.Zero);

                    //if(panes.Count > 2)
                    //{
                    //    IntPtr btnInstall = FindWindowEx(panes[1], IntPtr.Zero, "WindowsForms10.BUTTON.app.0.34f5582_r6_ad1", "Install");

                    //    SendMessage(btnInstall, BN_CLICKED, 0, IntPtr.Zero);

                    //    Thread.Sleep(2000);

                    //    IntPtr btnSelectBSP = FindWindowEx(hwnd, IntPtr.Zero, "WindowsForms10.BUTTON.app.0.34f5582_r6_ad1", "Select BSP");

                    //    SendMessage(btnSelectBSP, BN_CLICKED, 0, IntPtr.Zero);

                    //    Thread.Sleep(2000);

                    //    var dataExcel = ExcelHelper.ReadFileExcel(@"D:\DuAn\DUAN-AutoClicker\Test\Excels\File1.xlsx");

                    //    if(dataExcel.Count > 0)
                    //    {
                    //        var googleSheetsHelper = new GoogleSheetsHelper(pathCredential, spreadsheetId, sheetName, dataExcel);
                    //        googleSheetsHelper.WirteDatas();
                    //    }
                    //}

                    #region test
                    if (this.count == 0)
                        this.run = false;
                    #endregion

                    this.count++;
                    MethodInvoker countLabelUpdater = new MethodInvoker(() => {
                        SetCountLabel(this.count);
                    });
                    this.Invoke(countLabelUpdater);
                    Thread.Sleep(timeSleep);
                }
                catch (Exception ex)
                {
                    this.run = false;
                    EnableControls();
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void SetCountLabel(long count)
        {
            this.lbCount.Text = count + " Total Actions";
        }

        private void number_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
    }
}
