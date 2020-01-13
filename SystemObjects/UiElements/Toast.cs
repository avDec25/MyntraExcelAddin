using System;
using System.Drawing;
using System.Windows.Forms;

namespace MyntraExcelAddin.SystemObjects.UiElements
{
    public partial class Toast : Form
    {
        private Timer tmr;

        public Toast(string caption, string text, double timeout = 3.0)
        {
            InitializeComponent();
            
            this.Text = caption;
            this.toasttext.Text = text;

            Rectangle workingArea = Screen.GetWorkingArea(this);
            this.Location = new Point(workingArea.Right - Size.Width,
                                      workingArea.Bottom - Size.Height);

            Show();

            tmr = new Timer();
            tmr.Tick += delegate {
                Close();
            };
            tmr.Interval = (int)TimeSpan.FromSeconds(timeout).TotalMilliseconds;
            tmr.Start();
        }

        public void ChangeBackColor(int r, int g, int b)
        {
            BackColor = Color.FromArgb(r, g, b);
        }

        public static void Show(string text, string caption, double timeout = 3.0)
        {
            new Toast(text, caption, timeout);
        }
    }
}
