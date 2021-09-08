using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;

namespace TestWinAPI
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class MainForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnGetWindows;
		private const int WM_CLOSE = 16;
		private const int BN_CLICKED = 245;
		private System.Windows.Forms.Button btnCloseCalc;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		/// <summary>
		/// The FindWindow API
		/// </summary>
		/// <param name="lpClassName">the class name for the window to search for</param>
		/// <param name="lpWindowName">the name of the window to search for</param>
		/// <returns></returns>
		[DllImport("User32.dll")]
		public static extern IntPtr FindWindow(String lpClassName,String lpWindowName);

		/// <summary>
		/// The SendMessage API
		/// </summary>
		/// <param name="hWnd">handle to the required window</param>
		/// <param name="msg">the system/Custom message to send</param>
		/// <param name="wParam">first message parameter</param>
		/// <param name="lParam">second message parameter</param>
		/// <returns></returns>
		[DllImport("user32.dll", CharSet=CharSet.Auto)]
		public static extern int SendMessage(IntPtr hWnd, int msg, int wParam, IntPtr lParam);
		
		/// <summary>
		/// The FindWindowEx API
		/// </summary>
		/// <param name="parentHandle">a handle to the parent window </param>
		/// <param name="childAfter">a handle to the child window to start search after</param>
		/// <param name="className">the class name for the window to search for</param>
		/// <param name="windowTitle">the name of the window to search for</param>
		/// <returns></returns>
		[DllImport("user32.dll", SetLastError = true)]
		public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className,  string  windowTitle);

		[DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
		public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

		public MainForm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnGetWindows = new System.Windows.Forms.Button();
			this.btnCloseCalc = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// btnGetWindows
			// 
			this.btnGetWindows.Location = new System.Drawing.Point(16, 24);
			this.btnGetWindows.Name = "btnGetWindows";
			this.btnGetWindows.Size = new System.Drawing.Size(120, 23);
			this.btnGetWindows.TabIndex = 0;
			this.btnGetWindows.Text = "Use Calculator";
			this.btnGetWindows.Click += new System.EventHandler(this.btnGetWindows_Click);
			// 
			// btnCloseCalc
			// 
			this.btnCloseCalc.Location = new System.Drawing.Point(16, 72);
			this.btnCloseCalc.Name = "btnCloseCalc";
			this.btnCloseCalc.Size = new System.Drawing.Size(120, 23);
			this.btnCloseCalc.TabIndex = 2;
			this.btnCloseCalc.Text = "Close the Calculator";
			this.btnCloseCalc.Click += new System.EventHandler(this.btnCloseCalc_Click);
			// 
			// MainForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(234, 119);
			this.Controls.Add(this.btnCloseCalc);
			this.Controls.Add(this.btnGetWindows);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "MainForm";
			this.Text = "Test Windows API";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new MainForm());
		}

		private void btnGetWindows_Click(object sender, System.EventArgs e)
		{
            IntPtr hwnd = IntPtr.Zero;
            IntPtr hwndChild = IntPtr.Zero;

            //Get a handle for the Calculator Application main window
            hwnd = FindWindow(null, "UniKey 4.2 RC4");
            if (hwnd == IntPtr.Zero)
            {
                if (MessageBox.Show("Couldn't find the UniKey 4.2 RC4 application. Do you want to start it?", "TestWinAPI", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(@"C:\Users\ThanhDo\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\UnikeyNT");
                }
            }
            else
            {

                //Get a handle for the "1" button
                hwndChild = FindWindowEx(hwnd, IntPtr.Zero, "Button", "Thông tin");

                //send BN_CLICKED message
                SendMessage(hwndChild, BN_CLICKED, 0, IntPtr.Zero);

                ////Get a handle for the "+" button
                //hwndChild = FindWindowEx(hwnd, IntPtr.Zero, "Button", "1");

                ////send BN_CLICKED message
                //SendMessage(hwndChild, BN_CLICKED, 0, IntPtr.Zero);

                ////Get a handle for the "+" button
                //hwndChild = FindWindowEx(hwnd, IntPtr.Zero, "Button", "+");

                ////send BN_CLICKED message
                //SendMessage(hwndChild, BN_CLICKED, 0, IntPtr.Zero);

                ////Get a handle for the "2" button
                //hwndChild = FindWindowEx(hwnd, IntPtr.Zero, "Button", "2");

                ////send BN_CLICKED message
                //SendMessage(hwndChild, BN_CLICKED, 0, IntPtr.Zero);

                ////Get a handle for the "=" button
                //hwndChild = FindWindowEx(hwnd, IntPtr.Zero, "Button", "=");

                ////send BN_CLICKED message
                //SendMessage(hwndChild, BN_CLICKED, 0, IntPtr.Zero);

            }


		}

		private void btnCloseCalc_Click(object sender, System.EventArgs e)
		{
			IntPtr hwnd = IntPtr.Zero;

			//Get a handle for the Calculator Application main window
			hwnd=FindWindow(null, "UniKey 4.2 RC4");

			//send WM_CLOSE system message
			if(hwnd!= IntPtr.Zero)
				SendMessage(hwnd,WM_CLOSE,0,IntPtr.Zero);
		}
	}
}
