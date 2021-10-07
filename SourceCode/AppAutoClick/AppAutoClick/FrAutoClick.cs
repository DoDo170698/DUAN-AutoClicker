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
using System.Diagnostics;

namespace AppAutoClick
{
    public partial class FrAutoClick : Form
    {
        private bool run = false;
        private long count = 0;
        private string programName = ConfigurationManager.AppSettings["ProgramName"];
        private string softwareUsername = ConfigurationManager.AppSettings["SoftwareUsername"];
        private string softwarePassword = ConfigurationManager.AppSettings["SoftwarePassword"];
        private string pathCredential = ConfigurationManager.AppSettings["PathCredential"];
        private string spreadsheetId = ConfigurationManager.AppSettings["SpreadsheetId"];
        private string sheetName = ConfigurationManager.AppSettings["SheetName"];
        private string pathFileExe = ConfigurationManager.AppSettings["PathFileExe"];
        private string pathFileExcel = ConfigurationManager.AppSettings["PathFileExcel"];
        private string fileNameExcel = ConfigurationManager.AppSettings["FileNameExcel"];
        //private string nameWindowMain = "Calculator";
        private string nameWindowLogin = " Login to Pathogen Asset Control System";
        private string nameWindowReLogin = "PACS - inactivity lock";
        private string nameWindowRepositoryManagement = "Pathogen Asset Control System - [Repository Management]";
        private string nameWindowSearchEngine = "Pathogen Asset Control System - [Search Engine]";
        private string nameWindowSaveAs = "Save As";

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

        private void CloseProgram()
        {
            Process cmd = new Process();
            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.RedirectStandardInput = true;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.UseShellExecute = false;
            cmd.Start();

            cmd.StandardInput.WriteLine($"taskkill/im {programName} /f");
            cmd.StandardInput.Flush();
            cmd.StandardInput.Close();
            cmd.WaitForExit();

            Thread.Sleep(2000);
        }

        private void AutoClick()
        {
            while (this.run)
            {
                try
                {
                    CloseProgram();
                    if (!IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowLogin) && 
                        !IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowReLogin) && 
                        !IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowRepositoryManagement) && 
                        !IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowSearchEngine))
                    {
                        System.Diagnostics.Process.Start(pathFileExe);
                        while (!IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowLogin))
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
                        Thread.Sleep(5000);

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
                        Thread.Sleep(5000);
                    }

                    if (IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowRepositoryManagement))
                    {
                        IntPtr windowRepositoryManagement = FindWindow(null, nameWindowRepositoryManagement);
                        ShowWindow(windowRepositoryManagement, SW_MAXIMIZE);
                        Thread.Sleep(4000);

                        BlockInput(true);
                        var pointRepositoryManagement = new POINT();
                        ClientToScreen(windowRepositoryManagement, ref pointRepositoryManagement);

                        LeftMouseClick(pointRepositoryManagement.x + 165, pointRepositoryManagement.y + 55);
                        Thread.Sleep(2000);

                        LeftMouseClick(pointRepositoryManagement.x + 20, pointRepositoryManagement.y + 100);
                        Thread.Sleep(2000);
                        BlockInput(false);
                    }

                    if (IsOpenSoftware("WindowsForms10.Window.8.app.0.2bf8098_r6_ad1", nameWindowSearchEngine))
                    {
                        IntPtr windowSearchEngine = FindWindow(null, nameWindowSearchEngine);

                        BlockInput(true);
                        var pointSearchEngine = new POINT();
                        ClientToScreen(windowSearchEngine, ref pointSearchEngine);

                        LeftMouseClick(pointSearchEngine.x + 45, pointSearchEngine.y + 390);
                        Thread.Sleep(2000);
                        BlockInput(false);

                        var btnSearchEngines = EnumAllWindows(windowSearchEngine, "WindowsForms10.Window.b.app.0.2bf8098_r6_ad1").ToList();

                        var btnExecute = btnSearchEngines[5];
                        SendMessage(btnExecute, WM_LBUTTONDOWN, 0, IntPtr.Zero);
                        SendMessage(btnExecute, WM_LBUTTONUP, 0, IntPtr.Zero);
                        Thread.Sleep(3000);

                        var staticSearchEngines = EnumAllWindows(windowSearchEngine, "WindowsForms10.STATIC.app.0.2bf8098_r6_ad1").ToList();

                        var staticDisplayRecord = staticSearchEngines[0];
                        StringBuilder record = new StringBuilder(65535);
                        while (!GetControlText(staticDisplayRecord).Contains("record"))
                        {
                            Thread.Sleep(2000);
                        }

                        SaveAsOnNewThread();

                        var btnExport = btnSearchEngines[11];
                        SendMessage(btnExport, WM_LBUTTONDOWN, 0, IntPtr.Zero);
                        SendMessage(btnExport, WM_LBUTTONUP, 0, IntPtr.Zero);

                    }


                    this.count++;
                    MethodInvoker countLabelUpdater = new MethodInvoker(() => {
                        SetCountLabel(this.count);
                    });
                    this.Invoke(countLabelUpdater);


                    CloseProgram();
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
        private void SaveAsOnNewThread()
        {
            Thread t = new Thread(SaveAs);
            t.IsBackground = true;
            t.Start();
        }

        private void SaveAs()
        {
            try
            {
                while (!IsOpenSoftware(null, nameWindowSaveAs))
                {
                    Thread.Sleep(2000);
                }

                IntPtr windowSaveAs = FindWindow(null, nameWindowSaveAs);
                var urlFileExcel = pathFileExcel + @"\" + fileNameExcel + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

                var editSaveAses = EnumAllWindows(windowSaveAs, "Edit").ToList();
                var editFileName = editSaveAses[0];
                SendMessage(editFileName, WM_KEYDOWN, 0x5A, 0x002C0001);
                SendMessage(editFileName, WM_CHAR, 0x7A, 0x002C0001);
                SendMessage(editFileName, WM_KEYUP, 0x5A, 0xC02C0001);
                SendMessage(editFileName, WM_SETTEXT, 0, urlFileExcel);
                Thread.Sleep(2000);

                //var toolbarWindow32s = EnumAllWindows(windowSaveAs, "ToolbarWindow32").ToList();
                //var toolbarAddress = toolbarWindow32s[5];
                //var address = GetControlText(toolbarAddress);
                //address = address.Replace("Address: ", "").Trim();

                IntPtr btnSave = FindWindowEx(windowSaveAs, IntPtr.Zero, "Button", "&Save");

                SendMessage(btnSave, WM_LBUTTONDOWN, 0, IntPtr.Zero);
                SendMessage(btnSave, WM_LBUTTONUP, 0, IntPtr.Zero);
                Thread.Sleep(3000);

                var dataExcel = new ExcelHelper(pathCredential, spreadsheetId, sheetName, urlFileExcel);
                dataExcel.ReadFileExcel();
                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                this.run = false;
                EnableControls();
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        private void number_Click(object sender, EventArgs e)
        {
            var textBox = (TextBox)sender;
            textBox.SelectAll();
            textBox.Focus();
        }
    }
}
