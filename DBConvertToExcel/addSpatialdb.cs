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
    public partial class addSpatialdb : Form
    {
        string saveshpPath;
        public addSpatialdb()
        {
            InitializeComponent();
        }
        public string SHPPath
        {
            get { return saveshpPath; }
            set { saveshpPath = value; }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "保存文件";
            saveFileDialog.Filter = "SHP文件(.shp)|*.shp|所有文件(*.*)|*.*";
            saveFileDialog.RestoreDirectory = true;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                saveshpPath = saveFileDialog.FileName;
            }
            this.Close();
            
        }
    }
}
