using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.OleDb;


namespace Maintainance_GeneralLines
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == 0)
            {
                Line5 line5 = new Line5();
                line5.ShowDialog();
            }
            else if (comboBox1.SelectedIndex == 1)
            {
                Form6 line6 = new Form6();
                line6.ShowDialog();
            }
            else if (comboBox1.SelectedIndex == 2)
            {
                Line7 line7 = new Line7();
                line7.ShowDialog();
            }
            else if (comboBox1.SelectedIndex == 3)
            {
                Line8 line8 = new Line8();
                line8.ShowDialog();
            }
            else if (comboBox1.SelectedIndex == 4)
            {
                Line9 line9 = new Line9();
                line9.ShowDialog();
            }
            else if (comboBox1.SelectedIndex == 5)
            {
                Line10 line10 = new Line10();
                line10.ShowDialog();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
