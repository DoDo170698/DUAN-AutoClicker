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
        private int run = 0; //0 đang tắt, 1 đang bật, 2 đang thực hiện chương trình
        private bool isSave = false;
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
        Thread autoClick;
        Thread autoSave;

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
            StopAutoClicker();
        }

        private void StartAutoClicker()
        {
            var messageError = ValidData();
            if (string.IsNullOrEmpty(messageError))
            {
                timeSleep = new TimeSpan(int.Parse(txtHour.Text), int.Parse(txtMinute.Text), 0);
                if (this.run == 0)
                    AutoClickOnNewThread();
                this.run = 1;
                DisableControls();
            }
            else
            {
                MessageBox.Show(messageError, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void StopAutoClicker()
        {
            if(this.run == 2)
            {
                var check = MessageBox.Show("Tiến trình đang thực hiện, bạn có muốn ngắt không?", "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                if(check == DialogResult.No)
                {
                    return;
                }
            }
            this.run = 0;
            this.isSave = false;
            if (this.autoClick != null)
                this.autoClick.Abort();
            if (this.autoSave != null)
                this.autoSave.Abort();
            EnableControls();
            CloseProgram();
        }
        
        private void DisableControls()
        {
            btnStart.Enabled = false;
            txtHour.Enabled = false;
            txtMinute.Enabled = false;
        }
        private void EnableControls()
        {
            this.count = 0;
            SetCountLabel(this.count);
            btnStart.Enabled = true;
            txtHour.Enabled = true;
            txtMinute.Enabled = true;

        }
        private void EnableControlsThread()
        {
            MethodInvoker enableControls = new MethodInvoker(() => {
                this.count = 0;
                SetCountLabel(this.count);
                btnStart.Enabled = true;
                txtHour.Enabled = true;
                txtMinute.Enabled = true;
            });
            this.Invoke(enableControls);
        }
        private string ValidData()
        {
            if (!File.Exists(pathFileExe))
            {
                return "Không thể bắt đầu, sai đường dẫn phần mềm";
            }
            var hourStr = txtHour.Text;
            var minuteStr = txtMinute.Text;
            if (string.IsNullOrEmpty(hourStr) || string.IsNullOrEmpty(minuteStr))
            {
                return "Giờ và phút không được để trống";
            }
            var hour = long.Parse(hourStr);
            var minute = long.Parse(minuteStr);
            if (hour == 0 && minute == 0)
            {
                return "Giờ và phút phải có giá trị";
            }
            if (hour > 100 || minute > 59)
            {
                return "Giờ không được lớn hơn 100 và phút không được lớn hơn 60";
            }
            if (!Directory.Exists(pathFileExcel))
            {
                DialogResult pathFileExcelResult = MessageBox.Show(string.Format("Không tồn tại đường dẫn đến folder '{0}', bạn có muốn tạo folder không?",pathFileExcel),
                    "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (pathFileExcelResult == DialogResult.Yes)
                {
                    Directory.CreateDirectory(pathFileExcel);
                }
                else if (pathFileExcelResult == DialogResult.No)
                {
                    return "Không thể khởi chạy";
                }
            }
            return string.Empty;
        }
        private bool IsOpenSoftware(string classWindow, string nameWindow)
        {
            IntPtr hwnd = Win32.FindWindow(classWindow, nameWindow);
            return hwnd != IntPtr.Zero;
        }
        private void AutoClickOnNewThread()
        {
            autoClick = new Thread(AutoClick);
            autoClick.IsBackground = true;
            autoClick.Start();
        }

        private void CloseProgram()
        {
            try
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
            catch (Exception ex)
            {
                MessageBox.Show("Thực hiện tắt chương trình bị gián đoạn", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LoggingHelper.Write(ex.Message);
            }
        }

        private void CloseProgramThread()
        {
            try
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
            catch (Exception ex)
            {
                if (ex is ThreadAbortException)
                    return;
                LoggingHelper.Write(ex.Message);
                throw new InvalidOperationException("Thực hiện tắt chương trình bị gián đoạn");
            }
        }

        private void AutoClick()
        {
            while (this.run != 0)
            {
                try
                {
                    this.run = 2;
                    this.isSave = false;
                    CloseProgramThread();
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

                        LeftMouseClick(pointSearchEngine.x + 45, pointSearchEngine.y + 380);
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




                    CloseProgramThread();
                    while (this.run != 0 && !this.isSave)
                    {
                        Thread.Sleep(1000);
                    }

                    this.run = 1;
                    this.count++;
                    MethodInvoker countLabelUpdater = new MethodInvoker(() => {
                        SetCountLabel(this.count);
                    });
                    this.Invoke(countLabelUpdater);

                    Thread.Sleep(timeSleep);
                }
                catch (Exception ex)
                {
                    this.run = 0;
                    this.isSave = false;
                    EnableControlsThread();
                    if (ex is InvalidOperationException)
                        MessageBox.Show(ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else if (ex is ThreadAbortException)
                        return;
                    else
                    {
                        MessageBox.Show("Thực hiện Auto bị gián đoạn", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoggingHelper.Write(ex.Message);
                    }
                }
            }
        }
        private void SaveAsOnNewThread()
        {
            autoSave = new Thread(SaveAs);
            autoSave.IsBackground = true;
            autoSave.Start();
        }

        private void SaveAs()
        {
            try
            {
                while (this.run != 0 && !IsOpenSoftware(null, nameWindowSaveAs))
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

                this.isSave = true;
            }
            catch (Exception ex)
            {
                this.run = 0;
                this.isSave = false;
                EnableControlsThread();
                if(ex is InvalidOperationException)
                    MessageBox.Show(ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else if (ex is ThreadAbortException)
                    return;
                else
                {
                    MessageBox.Show("Thực hiện Save file bị gián đoạn", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoggingHelper.Write(ex.Message);
                }
            }
        }

        private void SetCountLabel(long count)
        {
            this.lbCount.Text = count + " lần";
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
