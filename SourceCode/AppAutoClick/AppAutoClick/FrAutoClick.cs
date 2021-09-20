﻿using System;
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
                //if (IsNotOpenSoftware(nameWindowMain))
                //{
                //    System.Diagnostics.Process.Start(pathFileExe);
                //}
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
        }
        private void EnableControls()
        {
            btnStart.Enabled = true;
        }

        private bool IsNotOpenSoftware(string classWindow, string nameWindow)
        {
            IntPtr hwnd = Win32.FindWindow(classWindow, nameWindow);
            return hwnd == IntPtr.Zero;
        }
        private string ValidData()
        {
            string messageError = string.Empty;
            if (!File.Exists(pathFileExe))
            {
                messageError = "Can't start, software path is incorrect";
            }
            return messageError;
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
                if (IsNotOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowLogin))
                {
                    System.Diagnostics.Process.Start(pathFileExe);
                }
                else
                {
                    IntPtr windowLogin = FindWindow(null, nameWindowLogin);

                    //var paneLogins = EnumAllWindows(windowLogin, "WindowsForms10.Window.8.app.0.2bf8098_r6_ad1").ToList();

                    //IntPtr paneLogin = FindWindowEx(paneLogins[4], IntPtr.Zero, "WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", "PACS");

                    //var paneLoginInputs = EnumAllWindows(paneLogin, "WindowsForms10.Window.b.app.0.2bf8098_r6_ad1").ToList();

                    //var inputUsername = FindWindowEx(paneLoginInputs[1], IntPtr.Zero, "WindowsForms10.EDIT.app.0.2bf8098_r6_ad1", null);
                    //var inputPassword = FindWindowEx(paneLoginInputs[0], IntPtr.Zero, "WindowsForms10.EDIT.app.0.2bf8098_r6_ad1", null);

                    var inputUsername = EnumAllWindows(windowLogin, "WindowsForms10.EDIT.app.0.2bf8098_r6_ad1").ToList()[1];
                    var inputPassword = EnumAllWindows(windowLogin, "WindowsForms10.EDIT.app.0.2bf8098_r6_ad1").ToList()[0];

                    SendMessage(inputUsername, WM_SETTEXT, 0, softwareUsername);
                    SendMessage(inputPassword, WM_SETTEXT, 0, softwarePassword);

                    IntPtr btnOk = FindWindowEx(windowLogin, IntPtr.Zero, "WindowsForms10.Window.b.app.0.2bf8098_r6_ad1", "Ok");

                    SendMessage(btnOk, BN_CLICKED, 0, IntPtr.Zero);

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

                }


                //this.run = false;
                this.count++;
                MethodInvoker countLabelUpdater = new MethodInvoker(() => {
                    SetCountLabel(this.count);
                });
                this.Invoke(countLabelUpdater);
                Thread.Sleep(1000);
            }
        }
        private void SetCountLabel(long count)
        {
            this.lbCount.Text = count + " Total Actions";
        }
    }
}
