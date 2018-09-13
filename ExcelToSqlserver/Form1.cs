using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelToSqlserver
{
    public partial class Form1 : Form
    {
        string filePath = string.Empty;
        DataTable dt = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_open_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            SetFileDialogProperty(ofd);
            DialogResult dr = ofd.ShowDialog();
            if (dr == DialogResult.OK)
            {
                filePath = ofd.FileName;
                txt_path.Text = filePath;
                try
                {
                    dt = ExcelToTable(filePath);
                    if (dt != null)
                    {
                        dg.DataSource = null;
                        dg.DataSource = dt;
                        btn_import.Enabled = true;
                    }
                }
                catch(Exception e1) { Show(e1.Message); btn_import.Enabled = false;btn_open.Enabled = true; return; }
                
            }
           
        }
        private void SetFileDialogProperty(OpenFileDialog dig)
        {
            dig.CheckFileExists = true;
            dig.CheckPathExists = true;
            dig.Filter = "*.xls| *.xls|*.xlsx|*.xlsx";
            dig.Multiselect = false;
            dig.Title = "选择EXCEL文件";
        }
        private DataTable ExcelToTable(string filePath)
        {
            DataTable dt = null;
            if (string.IsNullOrEmpty(filePath))
                return dt;
            if (filePath.ToLower().EndsWith("xls"))
            {
                dt = NpoiHelper.GetInstance().ImportFromExcel(filePath, "0", 0);
            } else if(filePath.ToLower().EndsWith("xlsx"))
                dt = NpoiHelper.GetInstance(false).ImportFromExcel(filePath,"0", 0);

            return dt;
        }

        private void btn_import_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_table.Text))
            {
                Show("请输入导入的目的数据表名");
                return;
            }
            string message;
            DbHelper.InsetBatch(dt, txt_table.Text.Trim(), out message);
            dt = null;
            btn_import.Enabled = false;
            btn_open.Enabled = true;
            txt_table.Clear();
            MessageBox.Show(message);
        }

        private void Show(string text)
        {
            MessageBox.Show(text, "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
