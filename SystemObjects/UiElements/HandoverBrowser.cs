using MyntraExcelAddin.Entity;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using Newtonsoft.Json;
using MyntraExcelAddin.Service;
using System.Text;

namespace MyntraExcelAddin.SystemObjects.UiElements
{
    public partial class HandoverBrowser : Form
    {
        SheetUpdater sheetUpdater;
        ExternalServiceMessenger messenger;
        uint currentPage = 0;
        uint pageSize = 10;
        uint lastPage = 0;

        public HandoverBrowser(ExternalServiceMessenger messenger, SheetUpdater sheetUpdater)
        {
            InitializeComponent();
            this.messenger = messenger;
            this.sheetUpdater = sheetUpdater;
            ResetBrowser();
        }


        public static DataTable CreateDataTable<T>(IEnumerable<T> list)
        {
            Type type = typeof(T);
            var properties = type.GetProperties();

            DataTable dataTable = new DataTable();
            foreach (PropertyInfo info in properties)
            {
                dataTable.Columns.Add(new DataColumn(info.Name, Nullable.GetUnderlyingType(info.PropertyType) ?? info.PropertyType));
            }

            foreach (T entity in list)
            {
                object[] values = new object[properties.Length];
                for (int i = 0; i < properties.Length; i++)
                {
                    values[i] = properties[i].GetValue(entity);
                }

                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        public void ResetBrowser()
        {
            currentPage = 0;
            pageSize = uint.Parse(comboBox_pagesize.Text);

            Dictionary<string, List<string>> query = new Dictionary<string, List<string>>();
            query.Add("status", new List<string> { "DRAFT" });

            var tuple = messenger.GetFilteredHandovers(query, currentPage, pageSize);
            dataGridView1.DataSource = CreateDataTable<HandoverTableView>(tuple.Item1);
            lastPage = tuple.Item2;
            if(lastPage > 0)
            {
                --lastPage;
            }

            if(currentPage == lastPage)
            {
                previous.Enabled = false;
                next.Enabled = false;
            } 
            else
            {
                previous.Enabled = false;
                next.Enabled = true;
            }
        }

        private void next_Click(object sender, EventArgs e)
        {       
            ++currentPage;
            Dictionary<string, List<string>> query = new Dictionary<string, List<string>>();
            query.Add("status", new List<string> { "DRAFT" });
            var tuple = messenger.GetFilteredHandovers(query, currentPage, pageSize);
            dataGridView1.DataSource = CreateDataTable<HandoverTableView>(tuple.Item1);

            if(currentPage == lastPage) {
            	next.Enabled = false;
            	previous.Enabled = true;
            }
        }

        private void previous_Click(object sender, EventArgs e)
        {    
            --currentPage;
            Dictionary<string, List<string>> query = new Dictionary<string, List<string>>();
            query.Add("status", new List<string> { "DRAFT" });
            var tuple = messenger.GetFilteredHandovers(query, currentPage, pageSize);
            dataGridView1.DataSource = CreateDataTable<HandoverTableView>(tuple.Item1);

            if(currentPage == 0) {
            	next.Enabled = true;
            	previous.Enabled = false;
            }            
        }

        private void download_Click(object sender, EventArgs e)
        {
            List<string> allhids = new List<string>();
            for (int i = 0; i < dataGridView1.SelectedRows.Count; ++i)
            {
                var selectedRow = ((DataRowView)(dataGridView1.SelectedRows[i].DataBoundItem)).Row;
                allhids.Add(selectedRow.Field<string>("id"));
            }
            
            DialogResult dialogResult = MessageBox.Show("Handovers will be downloaded on your existing sheet, " +
                "this will overwrite any the data currently on your sheet; " +
                "Do you want to continue the Download?", "Download Handovers", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {
                string sep = "";
                string hids = "";
                foreach(String hid in allhids)
                {
                    hids += sep + hid;
                    sep = ",";
                }
                List<Handover> handoverlist = messenger.GetHandovers(hids);
                sheetUpdater.PutDownloadedHandoversOnSheet(handoverlist);
            }
            else if (dialogResult == DialogResult.No)
            {
                return;
            }
        }

        private void comboBox_pagesize_SelectedIndexChanged(object sender, EventArgs e)
        {
            ResetBrowser();
        }
    }
}
