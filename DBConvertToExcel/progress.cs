﻿using System;
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
    public partial class progress : Form
    {
       
        public progress()
        {
            InitializeComponent();
        }

        public void SetNotifyInfo(int percent, string message)
        {
            this.label1.Text = message;
            this.progressBar1.Value = percent;
        }

       
    }
}
