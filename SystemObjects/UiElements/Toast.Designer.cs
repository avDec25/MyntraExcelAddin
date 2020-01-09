namespace MyntraExcelAddin.SystemObjects.UiElements
{
    partial class Toast
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.toasttext = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // toasttext
            // 
            this.toasttext.AutoSize = true;
            this.toasttext.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toasttext.Location = new System.Drawing.Point(12, 9);
            this.toasttext.Name = "toasttext";
            this.toasttext.Size = new System.Drawing.Size(245, 32);
            this.toasttext.TabIndex = 0;
            this.toasttext.Text = "Message on the toast";
            // 
            // Toast
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.ClientSize = new System.Drawing.Size(800, 54);
            this.ControlBox = false;
            this.Controls.Add(this.toasttext);
            this.Name = "Toast";
            this.Text = "Toast";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Label toasttext;
    }
}