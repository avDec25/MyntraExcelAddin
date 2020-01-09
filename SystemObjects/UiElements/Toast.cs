using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyntraExcelAddin.SystemObjects.UiElements
{
    public partial class Toast : Form
    {
        private Timer tmr;

        public Toast(string caption, string text, double timeout)
        {
            InitializeComponent();

            this.Text = caption;
            this.toasttext.Text = text;

            Rectangle workingArea = Screen.GetWorkingArea(this);
            this.Location = new Point(workingArea.Right - Size.Width,
                                      workingArea.Bottom - Size.Height);

            this.Show();

            tmr = new Timer();
            tmr.Tick += delegate {
                this.Close();
            };
            tmr.Interval = (int)TimeSpan.FromSeconds(timeout).TotalMilliseconds;
            tmr.Start();
        }

        public static void Show(string text, string caption, double timeout = 3.0)
        {
            new Toast(text, caption, timeout);
        }
    }
}
