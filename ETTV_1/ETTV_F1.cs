using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.ApplicationServices;
using System.Collections;
using Autodesk.Revit.UI.Selection;


namespace ETTV_1
{
    public partial class ETTV_F1 : System.Windows.Forms.Form
    {
        public UIApplication uiapp1;
        public UIDocument uidoc1;
        public Autodesk.Revit.ApplicationServices.Application app;
        public Document doc1;

        public int Ex_Wall_Op;

        public ETTV_F1(ExternalCommandData commandData)
        {
            InitializeComponent();
            uiapp1 = commandData.Application;
            uidoc1 = uiapp1.ActiveUIDocument;
            app = uiapp1.Application;
            doc1 = uidoc1.Document;

           
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = false;
            
        }

        private void ETTV_F1_Load(object sender, EventArgs e)
        {
            button1.Enabled = false;            
            radioButton2.Checked = true;
            button2.Enabled = true;

           
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button2.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            Ex_Wall_Op = 1;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            Ex_Wall_Op = 2;
            this.Close();
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
        }
    }
}
