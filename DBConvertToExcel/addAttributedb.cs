using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DBConvertToExcel
{
    public partial class addAttributedb : Form
    {
        string saveFilePath;

        public string FilePath
        {
            get { return saveFilePath; }
            set { saveFilePath = value; }
        }
        public addAttributedb()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog savefileDialog = new SaveFileDialog();
            savefileDialog.Title = "保存文件";
            savefileDialog.Filter = "Excel文件（*.xls;*.xlsx)|*.xls;*xlsx|所有文件（*.*)|*.*";
            savefileDialog.RestoreDirectory = true;
            if (savefileDialog.ShowDialog() == DialogResult.OK)
            {
                saveFilePath = savefileDialog.FileName;
            }
            this.Close();
        }
    }
}
