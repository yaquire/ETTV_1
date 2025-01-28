﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ETTV_1
{
    public partial class ProgressDialog : Form
    {
        public ProgressDialog()
        {
            InitializeComponent();
        }

        public void UpdateProgress(int progressValue, string progressText)
        {
            progressBar.Value = progressValue;
            progressLabel.Text = progressText;
        }



        private void ProgressDialog_Load(object sender, EventArgs e)
        {
           
        }
    }
}
