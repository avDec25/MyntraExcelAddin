using System;
using System.Collections.Generic;
using System.Windows.Forms;
using MyntraExcelAddin.Constant;
using MyntraExcelAddin.SystemObjects.UiElements;

namespace MyntraExcelAddin.Service
{
    class NotificationService
    {
        public void NotifyForEmptyCells(int row, List<int> cols)
        {
            String emptycols = "";
            String sep = "\n";
            foreach(int i in cols)
            {
                emptycols += sep + Header.Name[i];                
            }

            if (cols.Count != 0)
            {         
                MessageBox.Show("In Row " + row + ", Please fill the following Fields before proceeding: " + emptycols, "Data Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ProcessComplete(string process, string status)
        {
            if(status.Equals("success"))
            {
                Toast.Show(process, "Success! No errors.");
            } 
            else
            {
                Toast box = new Toast(process, "Failed! There are some Errors. Please correct them and try again.");
                box.ChangeBackColor(252, 3, 3);                
            }            
        }

    }
}
