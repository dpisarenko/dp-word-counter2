using AutoHotkey.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;

namespace dp_word_counter
{


    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            Console.WriteLine("timer_Tick: " + DateTime.Now);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Console.WriteLine("Heyo!");
            this.timer.Start();
        }
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(HandleRef hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }
        private void button1_Click(object sender, EventArgs e)
        {
            // var ahk = AutoHotkeyEngine.Instance;
            // ahk.
            Process[] processes  = Process.GetProcesses();
            Process scrivenerProcess = null;
            foreach (Process curProcess in processes)
            {
                Console.WriteLine("Name: " + curProcess.ProcessName + ", title: " + curProcess.MainWindowTitle);
                if (curProcess.MainWindowTitle.EndsWith("- Scrivener"))
                {
                    scrivenerProcess = curProcess;
                    break;
                }
            }
            if (scrivenerProcess == null)
            {
                Console.WriteLine("Scrivener not found");
                return;
            }

            var rect = new RECT();
            
            GetWindowRect(scrivenerProcess.Handle, out rect);

            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            Graphics graphics = Graphics.FromImage(bmp);
            graphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, new System.Drawing.Size(width, height), CopyPixelOperation.SourceCopy);

            bmp.Save("c:\\tmp\\test.png", ImageFormat.Png);

            Console.WriteLine("Heyo!");
        }
    }
}
