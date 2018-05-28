using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Maintainance_GeneralLines
{

    public partial class FrmDisplay : Form
    {
        public string EventStatus = "";
        public string currentdata = "";
        public int Line = 0;

        public FrmDisplay(string CurrentData,int Data)
        {
            InitializeComponent();
            currentdata = CurrentData;
            this.OK.DialogResult = DialogResult.OK;
            this.Cancel.DialogResult = DialogResult.Cancel;
            Line = Data;
        }

        private void FrmDisplay_Load(object sender, EventArgs e)
        {
            lblCurrent.Text = currentdata;
            if(Line == 10)
            {      
                comboBox1.Items.Add("None");
                comboBox1.Items.Add("Line 1 fails");
                comboBox1.Items.Add("Line 1 repd.");
                comboBox1.Items.Add("Line 2 fails");
                comboBox1.Items.Add("Line 2 repd.");
                comboBox1.Items.Add("Line 3 fails");
                comboBox1.Items.Add("Line 3 repd.");
                comboBox1.Items.Add("Line 4 fails");
                comboBox1.Items.Add("Line 4 repd.");
                comboBox1.Items.Add("Line 5 fails");
                comboBox1.Items.Add("Line 5 repd.");
                comboBox1.Items.Add("Line 6 fails");
                comboBox1.Items.Add("Line 6 repd.");
                comboBox1.Items.Add("Line 7 fails");
                comboBox1.Items.Add("Line 7 repd.");
                comboBox1.Items.Add("Line 8 fails");
                comboBox1.Items.Add("Line 8 repd.");
                comboBox1.Items.Add("Line 9 fails");
                comboBox1.Items.Add("Line 9 repd.");
                comboBox1.Items.Add("Line 10 fails");
                comboBox1.Items.Add("Line 10 repd.");
            }
            else if (Line == 9)
            {
                comboBox1.Items.Add("None");
                comboBox1.Items.Add("Line 1 fails");
                comboBox1.Items.Add("Line 1 repd.");
                comboBox1.Items.Add("Line 2 fails");
                comboBox1.Items.Add("Line 2 repd.");
                comboBox1.Items.Add("Line 3 fails");
                comboBox1.Items.Add("Line 3 repd.");
                comboBox1.Items.Add("Line 4 fails");
                comboBox1.Items.Add("Line 4 repd.");
                comboBox1.Items.Add("Line 5 fails");
                comboBox1.Items.Add("Line 5 repd.");
                comboBox1.Items.Add("Line 6 fails");
                comboBox1.Items.Add("Line 6 repd.");
                comboBox1.Items.Add("Line 7 fails");
                comboBox1.Items.Add("Line 7 repd.");
                comboBox1.Items.Add("Line 8 fails");
                comboBox1.Items.Add("Line 8 repd.");
                comboBox1.Items.Add("Line 9 fails");
                comboBox1.Items.Add("Line 9 repd.");
                //comboBox1.Items.Add("Line 10 fails");
                //comboBox1.Items.Add("Line 10 repd.");
            }
            else if (Line == 8)
            {
                comboBox1.Items.Add("None");
                comboBox1.Items.Add("Line 1 fails");
                comboBox1.Items.Add("Line 1 repd.");
                comboBox1.Items.Add("Line 2 fails");
                comboBox1.Items.Add("Line 2 repd.");
                comboBox1.Items.Add("Line 3 fails");
                comboBox1.Items.Add("Line 3 repd.");
                comboBox1.Items.Add("Line 4 fails");
                comboBox1.Items.Add("Line 4 repd.");
                comboBox1.Items.Add("Line 5 fails");
                comboBox1.Items.Add("Line 5 repd.");
                comboBox1.Items.Add("Line 6 fails");
                comboBox1.Items.Add("Line 6 repd.");
                comboBox1.Items.Add("Line 7 fails");
                comboBox1.Items.Add("Line 7 repd.");
                comboBox1.Items.Add("Line 8 fails");
                comboBox1.Items.Add("Line 8 repd.");
                //comboBox1.Items.Add("Line 9 fails");
                //comboBox1.Items.Add("Line 9 repd.");
                //comboBox1.Items.Add("Line 10 fails");
                //comboBox1.Items.Add("Line 10 repd.");
            }
            else if (Line == 7)
            {
                comboBox1.Items.Add("None");
                comboBox1.Items.Add("Line 1 fails");
                comboBox1.Items.Add("Line 1 repd.");
                comboBox1.Items.Add("Line 2 fails");
                comboBox1.Items.Add("Line 2 repd.");
                comboBox1.Items.Add("Line 3 fails");
                comboBox1.Items.Add("Line 3 repd.");
                comboBox1.Items.Add("Line 4 fails");
                comboBox1.Items.Add("Line 4 repd.");
                comboBox1.Items.Add("Line 5 fails");
                comboBox1.Items.Add("Line 5 repd.");
                comboBox1.Items.Add("Line 6 fails");
                comboBox1.Items.Add("Line 6 repd.");
                comboBox1.Items.Add("Line 7 fails");
                comboBox1.Items.Add("Line 7 repd.");
                //comboBox1.Items.Add("Line 8 fails");
                //comboBox1.Items.Add("Line 8 repd.");
                //comboBox1.Items.Add("Line 9 fails");
                //comboBox1.Items.Add("Line 9 repd.");
                //comboBox1.Items.Add("Line 10 fails");
                //comboBox1.Items.Add("Line 10 repd.");
            }
            else if (Line == 6)
            {
                comboBox1.Items.Add("None");
                comboBox1.Items.Add("Line 1 fails");
                comboBox1.Items.Add("Line 1 repd.");
                comboBox1.Items.Add("Line 2 fails");
                comboBox1.Items.Add("Line 2 repd.");
                comboBox1.Items.Add("Line 3 fails");
                comboBox1.Items.Add("Line 3 repd.");
                comboBox1.Items.Add("Line 4 fails");
                comboBox1.Items.Add("Line 4 repd.");
                comboBox1.Items.Add("Line 5 fails");
                comboBox1.Items.Add("Line 5 repd.");
                comboBox1.Items.Add("Line 6 fails");
                comboBox1.Items.Add("Line 6 repd.");
                //comboBox1.Items.Add("Line 7 fails");
                //comboBox1.Items.Add("Line 7 repd.");
                //comboBox1.Items.Add("Line 8 fails");
                //comboBox1.Items.Add("Line 8 repd.");
                //comboBox1.Items.Add("Line 9 fails");
                //comboBox1.Items.Add("Line 9 repd.");
                //comboBox1.Items.Add("Line 10 fails");
                //comboBox1.Items.Add("Line 10 repd.");
            }
            else if (Line == 5)
            {
                comboBox1.Items.Add("None");
                comboBox1.Items.Add("Line 1 fails");
                comboBox1.Items.Add("Line 1 repd.");
                comboBox1.Items.Add("Line 2 fails");
                comboBox1.Items.Add("Line 2 repd.");
                comboBox1.Items.Add("Line 3 fails");
                comboBox1.Items.Add("Line 3 repd.");
                comboBox1.Items.Add("Line 4 fails");
                comboBox1.Items.Add("Line 4 repd.");
                comboBox1.Items.Add("Line 5 fails");
                comboBox1.Items.Add("Line 5 repd.");
                //comboBox1.Items.Add("Line 6 fails");
                //comboBox1.Items.Add("Line 6 repd.");
                //comboBox1.Items.Add("Line 7 fails");
                //comboBox1.Items.Add("Line 7 repd.");
                //comboBox1.Items.Add("Line 8 fails");
                //comboBox1.Items.Add("Line 8 repd.");
                //comboBox1.Items.Add("Line 9 fails");
                //comboBox1.Items.Add("Line 9 repd.");
                //comboBox1.Items.Add("Line 10 fails");
                //comboBox1.Items.Add("Line 10 repd.");
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            string line = this.comboBox1.Text;
            EventStatus = line;
        }


    }
}
