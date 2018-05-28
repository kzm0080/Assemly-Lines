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
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
        }

        /// Global Variables

        List<clsLine6> clsLine7 = new List<clsLine6>();
        List<clsLine6StatTotals> clsLine7StatTotals = new List<clsLine6StatTotals>();
        GlobalData objGlobalData = new GlobalData();

        int k = 1;
        int kremaining = 0;
        string RemainingTime = "00:00:00";

        /// End Global Variables

        // Read Excel data from selected sheet
        public System.Data.DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            System.Data.DataTable dtexcel = new System.Data.DataTable();
            if (fileExt.CompareTo(".xls") == 0)//compare the extension of the file
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//for below excel 2007
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";//for above excel 2007
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    string sheetname = "1 Month (Data)";
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + sheetname + "$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
            return dtexcel;
        }

        // Select excel file to import data
        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty; int test = 0;

            objGlobalData.Line1 = 0;
            objGlobalData.Line2 = 0;
            objGlobalData.Line3 = 0;
            objGlobalData.Line4 = 0;
            objGlobalData.Line5 = 0;
            objGlobalData.Line6 = 0;
            //objGlobalData.Line7 = 0;
            //objGlobalData.Line8 = 0;
            //objGlobalData.Line9 = 0;
            //objGlobalData.Line10 = 0;

            objGlobalData.importcount = Convert.ToInt32(textBox1.Text);

            if (objGlobalData.importcount <= 0)
            { 
                objGlobalData.importcount = objGlobalData.Totalimportcount;
                textBox1.Text = (objGlobalData.importcount + 1).ToString();
            }
            else
                objGlobalData.importcount = objGlobalData.importcount - 1; // to match the records

            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        System.Data.DataTable dtExcel = new System.Data.DataTable();
                        dtExcel = ReadExcel(filePath, fileExt);
                        if (dtExcel.Rows.Count > 0)
                        {
                            dtExcel.Rows[0].Delete();
                            dtExcel.Rows[1].Delete();
                            dtExcel.AcceptChanges();

                            dataGridView1.DataSource = null;

                            clsLine7 = new List<clsLine6>();
                            for (int i = 0; i < dtExcel.Rows.Count; i++)
                            {
                                clsLine6 line7 = new clsLine6();                                

                                if (i >= 1)
                                {
                                    line7.Record = i + 1;
                                    line7.Time = dtExcel.Rows[i][0].ToString();
                                    line7.EventStatus = dtExcel.Rows[i][1].ToString();
                                    line7.Line1 = dtExcel.Rows[i][2].ToString();
                                    line7.Line2 = dtExcel.Rows[i][3].ToString();
                                    line7.Line3 = dtExcel.Rows[i][4].ToString();
                                    line7.Line4 = dtExcel.Rows[i][5].ToString();
                                    line7.Line5 = dtExcel.Rows[i][6].ToString();
                                    line7.Line6 = dtExcel.Rows[i][7].ToString();
                                    //line7.Line7 = dtExcel.Rows[i][8].ToString();

                                    if (i < objGlobalData.importcount)
                                    {

                                        TimeSpan ts1 = new TimeSpan(int.Parse(line7.Time.Split(':')[0]),    // hours
                                            int.Parse(line7.Time.Split(':')[1]),    // minutes
                                            int.Parse(line7.Time.Split(':')[2]));

                                        string endtime = dtExcel.Rows[i + 1][0].ToString();

                                        TimeSpan ts2 = new TimeSpan(int.Parse(endtime.Split(':')[0]),
                                            int.Parse(endtime.Split(':')[1]), int.Parse(endtime.Split(':')[2]));

                                        TimeSpan ds = (ts2 - ts1);

                                        line7.Wait = ds.ToString();
                                    }
                                    else
                                        line7.Wait = "00:00:00";
                                }
                                else
                                {
                                    line7.Record = i + 1;
                                    line7.Time = "00:00:00";
                                    line7.EventStatus = dtExcel.Rows[i][1].ToString();
                                    line7.Line1 = dtExcel.Rows[i][2].ToString();
                                    line7.Line2 = dtExcel.Rows[i][3].ToString();
                                    line7.Line3 = dtExcel.Rows[i][4].ToString();
                                    line7.Line4 = dtExcel.Rows[i][5].ToString();
                                    line7.Line5 = dtExcel.Rows[i][6].ToString();
                                    line7.Line6 = dtExcel.Rows[i][7].ToString();
                                    //line7.Line7 = dtExcel.Rows[i][8].ToString();
                                    line7.Wait = "00:00:00";


                                }
                                if (i <= objGlobalData.importcount)
                                {

                                    if (line7.EventStatus == "Line 1 fails" || line7.EventStatus == "Line 2 fails" || line7.EventStatus == "Line 3 fails" || line7.EventStatus == "Line 4 fails" || line7.EventStatus == "Line 5 fails" || line7.EventStatus == "Line 6 fails"
                                                 || line7.EventStatus == "Line 1 repd." || line7.EventStatus == "Line 2 repd." || line7.EventStatus == "Line 3 repd." || line7.EventStatus == "Line 4 repd." || line7.EventStatus == "Line 5 repd." || line7.EventStatus == "Line 6 repd." || line7.EventStatus == "" || line7.EventStatus == "None")
                                    {

                                    }
                                    else
                                    {
                                        if (cmbAuto.SelectedIndex == 0)
                                        {
                                            DialogResult dr = new DialogResult();
                                            FrmDisplay frmd = new FrmDisplay(line7.EventStatus, 6);
                                            dr = frmd.ShowDialog();
                                            if (dr == DialogResult.OK)
                                            {

                                                if (frmd.EventStatus == "")
                                                {
                                                    DialogResult dr1 = new DialogResult();
                                                    FrmDisplay frmd1 = new FrmDisplay(line7.EventStatus, 6);
                                                    dr1 = frmd1.ShowDialog();
                                                    if (dr1 == DialogResult.OK)
                                                    {
                                                        line7.EventStatus = frmd1.EventStatus;

                                                        if (line7.EventStatus == "")
                                                        {
                                                            MessageBox.Show("Please correct the data in Event Status at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                                            return;
                                                        }
                                                    }
                                                }
                                                else
                                                    line7.EventStatus = frmd.EventStatus;

                                                if (line7.EventStatus == "Line 1 fails" || line7.EventStatus == "Line 2 fails" || line7.EventStatus == "Line 3 fails" || line7.EventStatus == "Line 4 fails" || line7.EventStatus == "Line 5 fails" || line7.EventStatus == "Line 6 fails" 
                                                || line7.EventStatus == "Line 1 repd." || line7.EventStatus == "Line 2 repd." || line7.EventStatus == "Line 3 repd." || line7.EventStatus == "Line 4 repd." || line7.EventStatus == "Line 5 repd." || line7.EventStatus == "Line 6 repd." || line7.EventStatus == "" || line7.EventStatus == "None")
                                                {

                                                }
                                                else
                                                {
                                                    MessageBox.Show("Please correct the data in Event Status at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                                    return;
                                                }

                                            }
                                            else
                                            {
                                                MessageBox.Show("Please correct the data in Event Status at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                                return;
                                            }

                                        }
                                        else
                                        {
                                            MessageBox.Show("Please correct the data in Event Status at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }

                                    }

                                    if (line7.EventStatus == "Line 1 fails")
                                    {
                                        if (objGlobalData.Line1 == 1)
                                        {
                                            MessageBox.Show("Multiple Failures for Line1 with out repair at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line1 = 1;
                                    }
                                    else if (line7.EventStatus == "Line 1 repd.")
                                    {
                                        if (objGlobalData.Line1 == 0)
                                        {
                                            MessageBox.Show("Multiple Repairs for Line1 with out fail at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line1 = 0;
                                    }

                                    else if (line7.EventStatus == "Line 2 fails")
                                    {
                                        if (objGlobalData.Line2 == 1)
                                        {
                                            MessageBox.Show("Multiple Failures for Line2 with out repair at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line2 = 1;
                                    }
                                    else if (line7.EventStatus == "Line 2 repd.")
                                    {
                                        if (objGlobalData.Line2 == 0)
                                        {
                                            MessageBox.Show("Multiple Repairs for Line2 with out fail at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line2 = 0;
                                    }
                                    else if (line7.EventStatus == "Line 3 fails")
                                    {
                                        if (objGlobalData.Line3 == 1)
                                        {
                                            MessageBox.Show("Multiple Failures for Line3 with out repair at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line3 = 1;
                                    }
                                    else if (line7.EventStatus == "Line 3 repd.")
                                    {
                                        if (objGlobalData.Line3 == 0)
                                        {
                                            MessageBox.Show("Multiple Repairs for Line3 with out fail at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line3 = 0;
                                    }
                                    else if (line7.EventStatus == "Line 4 fails")
                                    {
                                        if (objGlobalData.Line4 == 1)
                                        {
                                            MessageBox.Show("Multiple Failures for Line4 with out repair at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line4 = 1;
                                    }
                                    else if (line7.EventStatus == "Line 4 repd.")
                                    {
                                        if (objGlobalData.Line4 == 0)
                                        {
                                            MessageBox.Show("Multiple Repairs for Line4 with out fail at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line4 = 0;
                                    }
                                    else if (line7.EventStatus == "Line 5 fails")
                                    {
                                        if (objGlobalData.Line5 == 1)
                                        {
                                            MessageBox.Show("Multiple Failures for Line5 with out repair at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line5 = 1;
                                    }
                                    else if (line7.EventStatus == "Line 5 repd.")
                                    {
                                        if (objGlobalData.Line5 == 0)
                                        {
                                            MessageBox.Show("Multiple Repairs for Line5 with out fail at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line5 = 0;
                                    }
                                    else if (line7.EventStatus == "Line 6 fails")
                                    {
                                        if (objGlobalData.Line6 == 1)
                                        {
                                            MessageBox.Show("Multiple Failures for Line6 with out repair at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line6 = 1;
                                    }
                                    else if (line7.EventStatus == "Line 6 repd.")
                                    {
                                        if (objGlobalData.Line6 == 0)
                                        {
                                            MessageBox.Show("Multiple Repairs for Line6 with out fail at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                            return;
                                        }
                                        else
                                            objGlobalData.Line6 = 0;
                                    }
                                    //else if (line7.EventStatus == "Line 7 fails")
                                    //{
                                    //    if (objGlobalData.Line7 == 1)
                                    //    {
                                    //        MessageBox.Show("Multiple Failures for Line7 with out repair at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                    //        return;
                                    //    }
                                    //    else
                                    //        objGlobalData.Line7 = 1;
                                    //}

                                    //else if (line7.EventStatus == "Line 7 repd.")
                                    //{
                                    //    if (objGlobalData.Line7 == 0)
                                    //    {
                                    //        MessageBox.Show("Multiple Repairs for Line7 with out fail at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                    //        return;
                                    //    }
                                    //    else
                                    //        objGlobalData.Line7 = 0;
                                    //}
                                    //else if (line7.EventStatus == "Line 8 fails")
                                    //{
                                    //    if (objGlobalData.Line8 == 1)
                                    //    {
                                    //        MessageBox.Show("Multiple Failures for Line8 with out repair at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                    //        return;
                                    //    }
                                    //    else
                                    //        objGlobalData.Line8 = 1;
                                    //}

                                    //else if (line7.EventStatus == "Line 8 repd.")
                                    //{
                                    //    if (objGlobalData.Line8 == 0)
                                    //    {
                                    //        MessageBox.Show("Multiple Repairs for Line8 with out fail at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                    //        return;
                                    //    }
                                    //    else
                                    //        objGlobalData.Line8 = 0;
                                    //}
                                    //else if (line7.EventStatus == "Line 9 fails")
                                    //{
                                    //    if (objGlobalData.Line9 == 1)
                                    //    {
                                    //        MessageBox.Show("Multiple Failures for Line9 with out repair at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                    //        return;
                                    //    }
                                    //    else
                                    //        objGlobalData.Line9 = 1;
                                    //}

                                    //else if (line7.EventStatus == "Line 9 repd.")
                                    //{
                                    //    if (objGlobalData.Line9 == 0)
                                    //    {
                                    //        MessageBox.Show("Multiple Repairs for Line9 with out fail at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                    //        return;
                                    //    }
                                    //    else
                                    //        objGlobalData.Line9 = 0;
                                    //}
                                    //else if (line7.EventStatus == "Line 10 fails")
                                    //{
                                    //    if (objGlobalData.Line10 == 1)
                                    //    {
                                    //        MessageBox.Show("Multiple Failures for Line10 with out repair at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                    //        return;
                                    //    }
                                    //    else
                                    //        objGlobalData.Line10 = 1;
                                    //}
                                    //else if (line7.EventStatus == "Line 10 repd.")
                                    //{
                                    //    if (objGlobalData.Line10 == 0)
                                    //    {
                                    //        MessageBox.Show("Multiple Repairs for Line10 with out fail at Row Number : " + (line7.Record + 3).ToString() + " For Time : " + line7.Time.ToString() + "");
                                    //        return;
                                    //    }
                                    //    else
                                    //        objGlobalData.Line10 = 0;
                                    //}


                                    objGlobalData.TotalTime = new TimeSpan(int.Parse(line7.Time.Split(':')[0]),    // hours
                                            int.Parse(line7.Time.Split(':')[1]),    // minutes
                                            int.Parse(line7.Time.Split(':')[2]));
                                    clsLine7.Add(line7);
                                }

                            }

                            if (objGlobalData.Line1 != 0 || objGlobalData.Line2 != 0 || objGlobalData.Line3 != 0 || objGlobalData.Line4 != 0 || objGlobalData.Line5 != 0 || objGlobalData.Line6 != 0 || objGlobalData.Line7 != 0 || objGlobalData.Line8 != 0 || objGlobalData.Line9 != 0 || objGlobalData.Line10 != 0)
                            {
                                if (objGlobalData.Line1 != 0)
                                {
                                    MessageBox.Show("Please repair line1 and import again");
                                    return;
                                }
                                else if (objGlobalData.Line2 != 0)
                                {
                                    MessageBox.Show("Please repair line2 and import again");
                                    return;
                                }
                                else if (objGlobalData.Line3 != 0)
                                {
                                    MessageBox.Show("Please repair line3 and import again");
                                    return;
                                }
                                else if (objGlobalData.Line4 != 0)
                                {
                                    MessageBox.Show("Please repair line4 and import again");
                                    return;
                                }
                                else if (objGlobalData.Line5 != 0)
                                {
                                    MessageBox.Show("Please repair line5 and import again");
                                    return;
                                }
                                else if (objGlobalData.Line6 != 0)
                                {
                                    MessageBox.Show("Please repair line6 and import again");
                                    return;
                                }
                                else if (objGlobalData.Line7 != 0)
                                {
                                    MessageBox.Show("Please repair line7 and import again");
                                    return;
                                }
                                else if (objGlobalData.Line8 != 0)
                                {
                                    MessageBox.Show("Please repair line8 and import again");
                                    return;
                                }
                                else if (objGlobalData.Line9 != 0)
                                {
                                    MessageBox.Show("Please repair line9 and import again");
                                    return;
                                }
                                else if (objGlobalData.Line10 != 0)
                                {
                                    MessageBox.Show("Please repair line10 and import again");
                                    return;
                                }

                            }

                            dataGridView1.DataSource = clsLine7;
                            dataGridView1.Columns["Wait"].Visible = false;

                        }
                        else
                        {
                            MessageBox.Show("Unable to import data from excel");
                            return;
                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        List<clsLine6Statistics> clsLine7Stats = new List<Maintainance_GeneralLines.clsLine6Statistics>();
        clsLine6StatTotals clsLine7StatTotal = new clsLine6StatTotals();

        // To caluculate totals for the given lines
        private void btnWaittime_Click(object sender, EventArgs e)
        {
            clsLine7StatTotal = new clsLine6StatTotals();
            clsLine7Stats = new List<Maintainance_GeneralLines.clsLine6Statistics>();

            TimeSpan TsZero = new TimeSpan(int.Parse(RemainingTime.Split(':')[0]),    // hours
                                            int.Parse(RemainingTime.Split(':')[1]),    // minutes
                                            int.Parse(RemainingTime.Split(':')[2]));

            // Line1 Totals
            clsLine7StatTotal.Line1DownTotal = RemainingTime;
            clsLine7StatTotal.Line1UpTotal = RemainingTime;
            clsLine7StatTotal.Line1DownKTotal = kremaining;
            clsLine7StatTotal.Line1UpKTotal = kremaining;
            clsLine7StatTotal.Line1WaitingTotal = RemainingTime;
            clsLine7StatTotal.Line1Down = TsZero;
            clsLine7StatTotal.Line1Up = TsZero;
            clsLine7StatTotal.Line1Wait = TsZero;

            // Line2 Totals
            clsLine7StatTotal.Line2DownTotal = RemainingTime;
            clsLine7StatTotal.Line2UpTotal = RemainingTime;
            clsLine7StatTotal.Line2DownKTotal = kremaining;
            clsLine7StatTotal.Line2UpKTotal = kremaining;
            clsLine7StatTotal.Line2WaitingTotal = RemainingTime;
            clsLine7StatTotal.Line2Down = TsZero;
            clsLine7StatTotal.Line2Up = TsZero;
            clsLine7StatTotal.Line2Wait = TsZero;

            // Line3 Totals
            clsLine7StatTotal.Line3DownTotal = RemainingTime;
            clsLine7StatTotal.Line3UpTotal = RemainingTime;
            clsLine7StatTotal.Line3DownKTotal = kremaining;
            clsLine7StatTotal.Line3UpKTotal = kremaining;
            clsLine7StatTotal.Line3WaitingTotal = RemainingTime;
            clsLine7StatTotal.Line3Down = TsZero;
            clsLine7StatTotal.Line3Up = TsZero;
            clsLine7StatTotal.Line3Wait = TsZero;


            // Line4 Totals
            clsLine7StatTotal.Line4DownTotal = RemainingTime;
            clsLine7StatTotal.Line4UpTotal = RemainingTime;
            clsLine7StatTotal.Line4DownKTotal = kremaining;
            clsLine7StatTotal.Line4UpKTotal = kremaining;
            clsLine7StatTotal.Line4WaitingTotal = RemainingTime;
            clsLine7StatTotal.Line4Down = TsZero;
            clsLine7StatTotal.Line4Up = TsZero;
            clsLine7StatTotal.Line4Wait = TsZero;


            // Line5 Totals
            clsLine7StatTotal.Line5DownTotal = RemainingTime;
            clsLine7StatTotal.Line5UpTotal = RemainingTime;
            clsLine7StatTotal.Line5DownKTotal = kremaining;
            clsLine7StatTotal.Line5UpKTotal = kremaining;
            clsLine7StatTotal.Line5WaitingTotal = RemainingTime;
            clsLine7StatTotal.Line5Down = TsZero;
            clsLine7StatTotal.Line5Up = TsZero;
            clsLine7StatTotal.Line5Wait = TsZero;


            // Line6 Totals
            clsLine7StatTotal.Line6DownTotal = RemainingTime;
            clsLine7StatTotal.Line6UpTotal = RemainingTime;
            clsLine7StatTotal.Line6DownKTotal = kremaining;
            clsLine7StatTotal.Line6UpKTotal = kremaining;
            clsLine7StatTotal.Line6WaitingTotal = RemainingTime;
            clsLine7StatTotal.Line6Down = TsZero;
            clsLine7StatTotal.Line6Up = TsZero;
            clsLine7StatTotal.Line6Wait = TsZero;


            // Line7 Totals
            //clsLine7StatTotal.Line7DownTotal = RemainingTime;
            //clsLine7StatTotal.Line7UpTotal = RemainingTime;
            //clsLine7StatTotal.Line7DownKTotal = kremaining;
            //clsLine7StatTotal.Line7UpKTotal = kremaining;
            //clsLine7StatTotal.Line7WaitingTotal = RemainingTime;
            //clsLine7StatTotal.Line7Down = TsZero;
            //clsLine7StatTotal.Line7Up = TsZero;
            //clsLine7StatTotal.Line7Wait = TsZero;

            if (clsLine7.Count > 0)
            {
                int i = 1;
                foreach (clsLine6 line7 in clsLine7)
                {
                    clsLine6Statistics clsLine7Statistics = new Maintainance_GeneralLines.clsLine6Statistics();
                    if (!string.IsNullOrEmpty(line7.Time))
                    {

                        clsLine7Statistics.Number = i.ToString();
                        TimeSpan TsTime = new TimeSpan(int.Parse(line7.Time.Split(':')[0]),    // hours
                                            int.Parse(line7.Time.Split(':')[1]),    // minutes
                                            int.Parse(line7.Time.Split(':')[2]));

                        TimeSpan TsWait = new TimeSpan(int.Parse(line7.Wait.Split(':')[0]),    // hours
                                            int.Parse(line7.Wait.Split(':')[1]),    // minutes
                                            int.Parse(line7.Wait.Split(':')[2]));

                        // Line 1 Statistics
                        if (line7.EventStatus == "Line 1 fails")
                        {
                            clsLine7Statistics.Line1Down = line7.Time;
                            clsLine7Statistics.Line1DownK = k;
                            clsLine7Statistics.Line1Up = RemainingTime;
                            clsLine7Statistics.Line1UpK = kremaining;

                            clsLine7StatTotal.Line1Down = clsLine7StatTotal.Line1Down + TsTime;
                            clsLine7StatTotal.Line1DownKTotal = clsLine7StatTotal.Line1DownKTotal + k;
                        }
                        else
                        {
                            clsLine7Statistics.Line1Down = RemainingTime;
                            clsLine7Statistics.Line1DownK = kremaining;
                            clsLine7Statistics.Line1Up = RemainingTime;
                            clsLine7Statistics.Line1UpK = kremaining;
                        }
                        if (line7.EventStatus == "Line 1 repd.")
                        {
                            clsLine7Statistics.Line1Down = RemainingTime;
                            clsLine7Statistics.Line1DownK = kremaining;
                            clsLine7Statistics.Line1Up = line7.Time;
                            clsLine7Statistics.Line1UpK = k;

                            clsLine7StatTotal.Line1Up = clsLine7StatTotal.Line1Up + TsTime;
                            clsLine7StatTotal.Line1UpKTotal = clsLine7StatTotal.Line1UpKTotal + k;

                        }
                        if (line7.Line1 == "Dn(waiting)")
                        {
                            clsLine7Statistics.Line1Waiting = line7.Wait;

                            clsLine7StatTotal.Line1Wait = clsLine7StatTotal.Line1Wait + TsWait;
                        }
                        else
                            clsLine7Statistics.Line1Waiting = RemainingTime;


                        // Line 2 Statistics
                        if (line7.EventStatus == "Line 2 fails")
                        {
                            clsLine7Statistics.Line2Down = line7.Time;
                            clsLine7Statistics.Line2DownK = k;
                            clsLine7Statistics.Line2Up = RemainingTime;
                            clsLine7Statistics.Line2UpK = kremaining;


                            clsLine7StatTotal.Line2Down = clsLine7StatTotal.Line2Down + TsTime;
                            clsLine7StatTotal.Line2DownKTotal = clsLine7StatTotal.Line2DownKTotal + k;

                        }
                        else
                        {
                            clsLine7Statistics.Line2Down = RemainingTime;
                            clsLine7Statistics.Line2DownK = kremaining;
                            clsLine7Statistics.Line2Up = RemainingTime;
                            clsLine7Statistics.Line2UpK = kremaining;
                        }

                        if (line7.EventStatus == "Line 2 repd.")
                        {
                            clsLine7Statistics.Line2Down = RemainingTime;
                            clsLine7Statistics.Line2DownK = kremaining;
                            clsLine7Statistics.Line2Up = line7.Time;
                            clsLine7Statistics.Line2UpK = k;

                            clsLine7StatTotal.Line2Up = clsLine7StatTotal.Line2Up + TsTime;
                            clsLine7StatTotal.Line2UpKTotal = clsLine7StatTotal.Line2UpKTotal + k;
                        }
                        if (line7.Line2 == "Dn(waiting)")
                        {
                            clsLine7Statistics.Line2Waiting = line7.Wait;

                            clsLine7StatTotal.Line2Wait = clsLine7StatTotal.Line2Wait + TsWait;

                        }
                        else
                            clsLine7Statistics.Line2Waiting = RemainingTime;


                        // Line 3 Statistics
                        if (line7.EventStatus == "Line 3 fails")
                        {
                            clsLine7Statistics.Line3Down = line7.Time;
                            clsLine7Statistics.Line3DownK = k;
                            clsLine7Statistics.Line3Up = RemainingTime;
                            clsLine7Statistics.Line3UpK = kremaining;


                            clsLine7StatTotal.Line3Down = clsLine7StatTotal.Line3Down + TsTime;
                            clsLine7StatTotal.Line3DownKTotal = clsLine7StatTotal.Line3DownKTotal + k;

                        }
                        else
                        {
                            clsLine7Statistics.Line3Down = RemainingTime;
                            clsLine7Statistics.Line3DownK = kremaining;
                            clsLine7Statistics.Line3Up = RemainingTime;
                            clsLine7Statistics.Line3UpK = kremaining;
                        }
                        if (line7.EventStatus == "Line 3 repd.")
                        {
                            clsLine7Statistics.Line3Down = RemainingTime;
                            clsLine7Statistics.Line3DownK = kremaining;
                            clsLine7Statistics.Line3Up = line7.Time;
                            clsLine7Statistics.Line3UpK = k;

                            clsLine7StatTotal.Line3Up = clsLine7StatTotal.Line3Up + TsTime;
                            clsLine7StatTotal.Line3UpKTotal = clsLine7StatTotal.Line3UpKTotal + k;

                        }
                        if (line7.Line3 == "Dn(waiting)")
                        {
                            clsLine7Statistics.Line3Waiting = line7.Wait;

                            clsLine7StatTotal.Line3Wait = clsLine7StatTotal.Line3Wait + TsWait;

                        }
                        else
                            clsLine7Statistics.Line3Waiting = RemainingTime;

                        // Line 4 Statistics
                        if (line7.EventStatus == "Line 4 fails")
                        {
                            clsLine7Statistics.Line4Down = line7.Time;
                            clsLine7Statistics.Line4DownK = k;
                            clsLine7Statistics.Line4Up = RemainingTime;
                            clsLine7Statistics.Line4UpK = kremaining;


                            clsLine7StatTotal.Line4Down = clsLine7StatTotal.Line4Down + TsTime;
                            clsLine7StatTotal.Line4DownKTotal = clsLine7StatTotal.Line4DownKTotal + k;

                        }
                        else
                        {
                            clsLine7Statistics.Line4Down = RemainingTime;
                            clsLine7Statistics.Line4DownK = kremaining;
                            clsLine7Statistics.Line4Up = RemainingTime;
                            clsLine7Statistics.Line4UpK = kremaining;
                        }
                        if (line7.EventStatus == "Line 4 repd.")
                        {
                            clsLine7Statistics.Line4Down = RemainingTime;
                            clsLine7Statistics.Line4DownK = kremaining;
                            clsLine7Statistics.Line4Up = line7.Time;
                            clsLine7Statistics.Line4UpK = k;

                            clsLine7StatTotal.Line4Up = clsLine7StatTotal.Line4Up + TsTime;
                            clsLine7StatTotal.Line4UpKTotal = clsLine7StatTotal.Line4UpKTotal + k;

                        }
                        if (line7.Line4 == "Dn(waiting)")
                        {
                            clsLine7Statistics.Line4Waiting = line7.Wait;

                            clsLine7StatTotal.Line4Wait = clsLine7StatTotal.Line4Wait + TsWait;

                        }
                        else
                            clsLine7Statistics.Line4Waiting = RemainingTime;


                        // Line 5 Statistics
                        if (line7.EventStatus == "Line 5 fails")
                        {
                            clsLine7Statistics.Line5Down = line7.Time;
                            clsLine7Statistics.Line5DownK = k;
                            clsLine7Statistics.Line5Up = RemainingTime;
                            clsLine7Statistics.Line5UpK = kremaining;


                            clsLine7StatTotal.Line5Down = clsLine7StatTotal.Line5Down + TsTime;
                            clsLine7StatTotal.Line5DownKTotal = clsLine7StatTotal.Line5DownKTotal + k;

                        }
                        else
                        {
                            clsLine7Statistics.Line5Down = RemainingTime;
                            clsLine7Statistics.Line5DownK = kremaining;
                            clsLine7Statistics.Line5Up = RemainingTime;
                            clsLine7Statistics.Line5UpK = kremaining;
                        }
                        if (line7.EventStatus == "Line 5 repd.")
                        {
                            clsLine7Statistics.Line5Down = RemainingTime;
                            clsLine7Statistics.Line5DownK = kremaining;
                            clsLine7Statistics.Line5Up = line7.Time;
                            clsLine7Statistics.Line5UpK = k;

                            clsLine7StatTotal.Line5Up = clsLine7StatTotal.Line5Up + TsTime;
                            clsLine7StatTotal.Line5UpKTotal = clsLine7StatTotal.Line5UpKTotal + k;

                        }
                        if (line7.Line5 == "Dn(waiting)")
                        {
                            clsLine7Statistics.Line5Waiting = line7.Wait;

                            clsLine7StatTotal.Line5Wait = clsLine7StatTotal.Line5Wait + TsWait;

                        }
                        else
                            clsLine7Statistics.Line5Waiting = RemainingTime;

                        // Line 6 Statistics
                        if (line7.EventStatus == "Line 6 fails")
                        {
                            clsLine7Statistics.Line6Down = line7.Time;
                            clsLine7Statistics.Line6DownK = k;
                            clsLine7Statistics.Line6Up = RemainingTime;
                            clsLine7Statistics.Line6UpK = kremaining;


                            clsLine7StatTotal.Line6Down = clsLine7StatTotal.Line6Down + TsTime;
                            clsLine7StatTotal.Line6DownKTotal = clsLine7StatTotal.Line6DownKTotal + k;

                        }
                        else
                        {
                            clsLine7Statistics.Line6Down = RemainingTime;
                            clsLine7Statistics.Line6DownK = kremaining;
                            clsLine7Statistics.Line6Up = RemainingTime;
                            clsLine7Statistics.Line6UpK = kremaining;
                        }
                        if (line7.EventStatus == "Line 6 repd.")
                        {
                            clsLine7Statistics.Line6Down = RemainingTime;
                            clsLine7Statistics.Line6DownK = kremaining;
                            clsLine7Statistics.Line6Up = line7.Time;
                            clsLine7Statistics.Line6UpK = k;

                            clsLine7StatTotal.Line6Up = clsLine7StatTotal.Line6Up + TsTime;
                            clsLine7StatTotal.Line6UpKTotal = clsLine7StatTotal.Line6UpKTotal + k;

                        }
                        if (line7.Line6 == "Dn(waiting)")
                        {
                            clsLine7Statistics.Line6Waiting = line7.Wait;

                            clsLine7StatTotal.Line6Wait = clsLine7StatTotal.Line6Wait + TsWait;

                        }
                        else
                            clsLine7Statistics.Line6Waiting = RemainingTime;


                        // Line 7 Statistics
                        //if (line7.EventStatus == "Line 7 fails")
                        //{
                        //    clsLine7Statistics.Line7Down = line7.Time;
                        //    clsLine7Statistics.Line7DownK = k;
                        //    clsLine7Statistics.Line7Up = RemainingTime;
                        //    clsLine7Statistics.Line7UpK = kremaining;


                        //    clsLine7StatTotal.Line7Down = clsLine7StatTotal.Line7Down + TsTime;
                        //    clsLine7StatTotal.Line7DownKTotal = clsLine7StatTotal.Line7DownKTotal + k;

                        //}
                        //else
                        //{
                        //    clsLine7Statistics.Line7Down = RemainingTime;
                        //    clsLine7Statistics.Line7DownK = kremaining;
                        //    clsLine7Statistics.Line7Up = RemainingTime;
                        //    clsLine7Statistics.Line7UpK = kremaining;
                        //}
                        //if (line7.EventStatus == "Line 7 repd.")
                        //{
                        //    clsLine7Statistics.Line7Down = RemainingTime;
                        //    clsLine7Statistics.Line7DownK = kremaining;
                        //    clsLine7Statistics.Line7Up = line7.Time;
                        //    clsLine7Statistics.Line7UpK = k;

                        //    clsLine7StatTotal.Line7Up = clsLine7StatTotal.Line7Up + TsTime;
                        //    clsLine7StatTotal.Line7UpKTotal = clsLine7StatTotal.Line7UpKTotal + k;

                        //}
                        //if (line7.Line7 == "Dn(waiting)")
                        //{
                        //    clsLine7Statistics.Line7Waiting = line7.Wait;

                        //    clsLine7StatTotal.Line7Wait = clsLine7StatTotal.Line7Wait + TsWait;

                        //}
                        //else
                        //    clsLine7Statistics.Line7Waiting = RemainingTime;
                    }

                    clsLine7Stats.Add(clsLine7Statistics);
                    i = i + 1;
                }
                clsLine6Statistics clsLine7StatisticsTot = new clsLine6Statistics();

                // Line 1 Stat Totals 
                clsLine7StatisticsTot.Line1Down = Math.Floor(clsLine7StatTotal.Line1Down.TotalHours).ToString() + ":" + clsLine7StatTotal.Line1Down.Minutes.ToString() + ":" + clsLine7StatTotal.Line1Down.Seconds.ToString();
                clsLine7StatisticsTot.Line1DownK = clsLine7StatTotal.Line1DownKTotal;
                clsLine7StatisticsTot.Line1Up = Math.Floor(clsLine7StatTotal.Line1Up.TotalHours).ToString() + ":" + clsLine7StatTotal.Line1Up.Minutes.ToString() + ":" + clsLine7StatTotal.Line1Up.Seconds.ToString();
                clsLine7StatisticsTot.Line1UpK = clsLine7StatTotal.Line1UpKTotal;
                clsLine7StatisticsTot.Line1Waiting = Math.Floor(clsLine7StatTotal.Line1Wait.TotalHours).ToString() + ":" + clsLine7StatTotal.Line1Wait.Minutes.ToString() + ":" + clsLine7StatTotal.Line1Wait.Seconds.ToString();


                // Line 2 Stat Totals 
                clsLine7StatisticsTot.Line2Down = Math.Floor(clsLine7StatTotal.Line2Down.TotalHours).ToString() + ":" + clsLine7StatTotal.Line2Down.Minutes.ToString() + ":" + clsLine7StatTotal.Line2Down.Seconds.ToString();
                clsLine7StatisticsTot.Line2DownK = clsLine7StatTotal.Line2DownKTotal;
                clsLine7StatisticsTot.Line2Up = Math.Floor(clsLine7StatTotal.Line2Up.TotalHours).ToString() + ":" + clsLine7StatTotal.Line2Up.Minutes.ToString() + ":" + clsLine7StatTotal.Line2Up.Seconds.ToString();
                clsLine7StatisticsTot.Line2UpK = clsLine7StatTotal.Line2UpKTotal;
                clsLine7StatisticsTot.Line2Waiting = Math.Floor(clsLine7StatTotal.Line2Wait.TotalHours).ToString() + ":" + clsLine7StatTotal.Line2Wait.Minutes.ToString() + ":" + clsLine7StatTotal.Line2Wait.Seconds.ToString();

                // Line 3 Stat Totals 
                clsLine7StatisticsTot.Line3Down = Math.Floor(clsLine7StatTotal.Line3Down.TotalHours).ToString() + ":" + clsLine7StatTotal.Line3Down.Minutes.ToString() + ":" + clsLine7StatTotal.Line3Down.Seconds.ToString();
                clsLine7StatisticsTot.Line3DownK = clsLine7StatTotal.Line3DownKTotal;
                clsLine7StatisticsTot.Line3Up = Math.Floor(clsLine7StatTotal.Line3Up.TotalHours).ToString() + ":" + clsLine7StatTotal.Line3Up.Minutes.ToString() + ":" + clsLine7StatTotal.Line3Up.Seconds.ToString();
                clsLine7StatisticsTot.Line3UpK = clsLine7StatTotal.Line3UpKTotal;
                clsLine7StatisticsTot.Line3Waiting = Math.Floor(clsLine7StatTotal.Line3Wait.TotalHours).ToString() + ":" + clsLine7StatTotal.Line3Wait.Minutes.ToString() + ":" + clsLine7StatTotal.Line3Wait.Seconds.ToString();

                // Line 4 Stat Totals 
                clsLine7StatisticsTot.Line4Down = Math.Floor(clsLine7StatTotal.Line4Down.TotalHours).ToString() + ":" + clsLine7StatTotal.Line4Down.Minutes.ToString() + ":" + clsLine7StatTotal.Line4Down.Seconds.ToString();
                clsLine7StatisticsTot.Line4DownK = clsLine7StatTotal.Line4DownKTotal;
                clsLine7StatisticsTot.Line4Up = Math.Floor(clsLine7StatTotal.Line4Up.TotalHours).ToString() + ":" + clsLine7StatTotal.Line4Up.Minutes.ToString() + ":" + clsLine7StatTotal.Line4Up.Seconds.ToString();
                clsLine7StatisticsTot.Line4UpK = clsLine7StatTotal.Line4UpKTotal;
                clsLine7StatisticsTot.Line4Waiting = Math.Floor(clsLine7StatTotal.Line4Wait.TotalHours).ToString() + ":" + clsLine7StatTotal.Line4Wait.Minutes.ToString() + ":" + clsLine7StatTotal.Line4Wait.Seconds.ToString();

                // Line 5 Stat Totals 
                clsLine7StatisticsTot.Line5Down = Math.Floor(clsLine7StatTotal.Line5Down.TotalHours).ToString() + ":" + clsLine7StatTotal.Line5Down.Minutes.ToString() + ":" + clsLine7StatTotal.Line5Down.Seconds.ToString();
                clsLine7StatisticsTot.Line5DownK = clsLine7StatTotal.Line5DownKTotal;
                clsLine7StatisticsTot.Line5Up = Math.Floor(clsLine7StatTotal.Line5Up.TotalHours).ToString() + ":" + clsLine7StatTotal.Line5Up.Minutes.ToString() + ":" + clsLine7StatTotal.Line5Up.Seconds.ToString();
                clsLine7StatisticsTot.Line5UpK = clsLine7StatTotal.Line5UpKTotal;
                clsLine7StatisticsTot.Line5Waiting = Math.Floor(clsLine7StatTotal.Line5Wait.TotalHours).ToString() + ":" + clsLine7StatTotal.Line5Wait.Minutes.ToString() + ":" + clsLine7StatTotal.Line5Wait.Seconds.ToString();

                // Line 6 Stat Totals 
                clsLine7StatisticsTot.Line6Down = Math.Floor(clsLine7StatTotal.Line6Down.TotalHours).ToString() + ":" + clsLine7StatTotal.Line6Down.Minutes.ToString() + ":" + clsLine7StatTotal.Line6Down.Seconds.ToString();
                clsLine7StatisticsTot.Line6DownK = clsLine7StatTotal.Line6DownKTotal;
                clsLine7StatisticsTot.Line6Up = Math.Floor(clsLine7StatTotal.Line6Up.TotalHours).ToString() + ":" + clsLine7StatTotal.Line6Up.Minutes.ToString() + ":" + clsLine7StatTotal.Line6Up.Seconds.ToString();
                clsLine7StatisticsTot.Line6UpK = clsLine7StatTotal.Line6UpKTotal;
                clsLine7StatisticsTot.Line6Waiting = Math.Floor(clsLine7StatTotal.Line6Wait.TotalHours).ToString() + ":" + clsLine7StatTotal.Line6Wait.Minutes.ToString() + ":" + clsLine7StatTotal.Line6Wait.Seconds.ToString();

                // Line 7 Stat Totals 
                //clsLine7StatisticsTot.Line7Down = Math.Floor(clsLine7StatTotal.Line7Down.TotalHours).ToString() + ":" + clsLine7StatTotal.Line7Down.Minutes.ToString() + ":" + clsLine7StatTotal.Line7Down.Seconds.ToString();
                //clsLine7StatisticsTot.Line7DownK = clsLine7StatTotal.Line7DownKTotal;
                //clsLine7StatisticsTot.Line7Up = Math.Floor(clsLine7StatTotal.Line7Up.TotalHours).ToString() + ":" + clsLine7StatTotal.Line7Up.Minutes.ToString() + ":" + clsLine7StatTotal.Line7Up.Seconds.ToString();
                //clsLine7StatisticsTot.Line7UpK = clsLine7StatTotal.Line7UpKTotal;
                //clsLine7StatisticsTot.Line7Waiting = Math.Floor(clsLine7StatTotal.Line7Wait.TotalHours).ToString() + ":" + clsLine7StatTotal.Line7Wait.Minutes.ToString() + ":" + clsLine7StatTotal.Line7Wait.Seconds.ToString();

                // Assign to global variables for statistics

                // Line 1
                objGlobalData.Line1Down = clsLine7StatTotal.Line1Down;
                objGlobalData.Line1Up = clsLine7StatTotal.Line1Up;
                objGlobalData.Line1Waiting = clsLine7StatTotal.Line1Wait;
                objGlobalData.Line1DownK = clsLine7StatTotal.Line1DownKTotal;
                objGlobalData.Line1UpK = clsLine7StatTotal.Line1UpKTotal;

                // Line 2
                objGlobalData.Line2Down = clsLine7StatTotal.Line2Down;
                objGlobalData.Line2Up = clsLine7StatTotal.Line2Up;
                objGlobalData.Line2Waiting = clsLine7StatTotal.Line2Wait;
                objGlobalData.Line2DownK = clsLine7StatTotal.Line2DownKTotal;
                objGlobalData.Line2UpK = clsLine7StatTotal.Line2UpKTotal;

                // Line 3
                objGlobalData.Line3Down = clsLine7StatTotal.Line3Down;
                objGlobalData.Line3Up = clsLine7StatTotal.Line3Up;
                objGlobalData.Line3Waiting = clsLine7StatTotal.Line3Wait;
                objGlobalData.Line3DownK = clsLine7StatTotal.Line3DownKTotal;
                objGlobalData.Line3UpK = clsLine7StatTotal.Line3UpKTotal;

                // Line 4
                objGlobalData.Line4Down = clsLine7StatTotal.Line4Down;
                objGlobalData.Line4Up = clsLine7StatTotal.Line4Up;
                objGlobalData.Line4Waiting = clsLine7StatTotal.Line4Wait;
                objGlobalData.Line4DownK = clsLine7StatTotal.Line4DownKTotal;
                objGlobalData.Line4UpK = clsLine7StatTotal.Line4UpKTotal;

                // Line 5
                objGlobalData.Line5Down = clsLine7StatTotal.Line5Down;
                objGlobalData.Line5Up = clsLine7StatTotal.Line5Up;
                objGlobalData.Line5Waiting = clsLine7StatTotal.Line5Wait;
                objGlobalData.Line5DownK = clsLine7StatTotal.Line5DownKTotal;
                objGlobalData.Line5UpK = clsLine7StatTotal.Line5UpKTotal;

                // Line 6
                objGlobalData.Line6Down = clsLine7StatTotal.Line6Down;
                objGlobalData.Line6Up = clsLine7StatTotal.Line6Up;
                objGlobalData.Line6Waiting = clsLine7StatTotal.Line6Wait;
                objGlobalData.Line6DownK = clsLine7StatTotal.Line6DownKTotal;
                objGlobalData.Line6UpK = clsLine7StatTotal.Line6UpKTotal;

                // Line 7
                //objGlobalData.Line7Down = clsLine7StatTotal.Line7Down;
                //objGlobalData.Line7Up = clsLine7StatTotal.Line7Up;
                //objGlobalData.Line7Waiting = clsLine7StatTotal.Line7Wait;
                //objGlobalData.Line7DownK = clsLine7StatTotal.Line7DownKTotal;
                //objGlobalData.Line7UpK = clsLine7StatTotal.Line7UpKTotal;



                clsLine7StatisticsTot.Number = i.ToString();

                clsLine7Stats.Add(clsLine7StatisticsTot);
                dataGridView2.Visible = true;
                if (dataGridView2.Rows.Count > 0)
                    dataGridView2.DataSource = null;
                dataGridView2.DataSource = clsLine7Stats;
            }
        }

        List<clsLine6MonthTotal> clsLine7MonthTotal = new List<Maintainance_GeneralLines.clsLine6MonthTotal>();
        private void btnTotals_Click(object sender, EventArgs e)
        {
            clsLine7MonthTotal = new List<Maintainance_GeneralLines.clsLine6MonthTotal>();
            for (int i = 1; i <= 7; i++)
            {
                clsLine6MonthTotal clsLine7MonthTot = new Maintainance_GeneralLines.clsLine6MonthTotal();
                TimeSpan TsDown = TimeSpan.Zero;
                TimeSpan TsUp = TimeSpan.Zero;
                TimeSpan TsWait = TimeSpan.Zero;
                if (i == 1) // Line1
                {
                    clsLine7MonthTot.Line = "Line" + i.ToString();
                    TsDown = objGlobalData.Line1Up - objGlobalData.Line1Down;
                    TsUp = objGlobalData.TotalTime - TsDown;
                    TsWait = objGlobalData.Line1Waiting;


                    double MTTF = getRemaining(TsUp, objGlobalData.Line1DownK);
                    double lambda = Math.Round(1 / MTTF, 5);

                    double MTTR = getRemaining(TsDown, objGlobalData.Line1UpK);
                    double Mu = Math.Round(1 / MTTR, 5);

                    clsLine7MonthTot.TotalDown = Math.Floor(TsDown.TotalHours).ToString() + ":" + TsDown.Minutes.ToString() + ":" + TsDown.Seconds.ToString();
                    clsLine7MonthTot.TotalUp = Math.Floor(TsUp.TotalHours).ToString() + ":" + TsUp.Minutes.ToString() + ":" + TsUp.Seconds.ToString();
                    clsLine7MonthTot.TotalWaiting = Math.Floor(TsWait.TotalHours).ToString() + ":" + TsWait.Minutes.ToString() + ":" + TsWait.Seconds.ToString();
                    clsLine7MonthTot.MTTF = MTTF.ToString();
                    clsLine7MonthTot.Lambda = lambda.ToString();
                    clsLine7MonthTot.MTTR = MTTR.ToString();
                    clsLine7MonthTot.MU = Mu.ToString();

                    clsLine7MonthTotal.Add(clsLine7MonthTot);

                }
                else if (i == 2) // Line2
                {
                    clsLine7MonthTot.Line = "Line" + i.ToString();
                    TsDown = objGlobalData.Line2Up - objGlobalData.Line2Down;
                    TsUp = objGlobalData.TotalTime - TsDown;
                    TsWait = objGlobalData.Line2Waiting;


                    double MTTF = getRemaining(TsUp, objGlobalData.Line2DownK);
                    double lambda = Math.Round(1 / MTTF, 5);

                    double MTTR = getRemaining(TsDown, objGlobalData.Line2UpK);
                    double Mu = Math.Round(1 / MTTR, 5);

                    clsLine7MonthTot.TotalDown = Math.Floor(TsDown.TotalHours).ToString() + ":" + TsDown.Minutes.ToString() + ":" + TsDown.Seconds.ToString();
                    clsLine7MonthTot.TotalUp = Math.Floor(TsUp.TotalHours).ToString() + ":" + TsUp.Minutes.ToString() + ":" + TsUp.Seconds.ToString();
                    clsLine7MonthTot.TotalWaiting = Math.Floor(TsWait.TotalHours).ToString() + ":" + TsWait.Minutes.ToString() + ":" + TsWait.Seconds.ToString();
                    clsLine7MonthTot.MTTF = MTTF.ToString();
                    clsLine7MonthTot.Lambda = lambda.ToString();
                    clsLine7MonthTot.MTTR = MTTR.ToString();
                    clsLine7MonthTot.MU = Mu.ToString();

                    clsLine7MonthTotal.Add(clsLine7MonthTot);

                }
                else if (i == 3) // Line3
                {
                    clsLine7MonthTot.Line = "Line" + i.ToString();
                    TsDown = objGlobalData.Line3Up - objGlobalData.Line3Down;
                    TsUp = objGlobalData.TotalTime - TsDown;
                    TsWait = objGlobalData.Line3Waiting;


                    double MTTF = getRemaining(TsUp, objGlobalData.Line3DownK);
                    double lambda = Math.Round(1 / MTTF, 5);

                    double MTTR = getRemaining(TsDown, objGlobalData.Line3UpK);
                    double Mu = Math.Round(1 / MTTR, 5);

                    clsLine7MonthTot.TotalDown = Math.Floor(TsDown.TotalHours).ToString() + ":" + TsDown.Minutes.ToString() + ":" + TsDown.Seconds.ToString();
                    clsLine7MonthTot.TotalUp = Math.Floor(TsUp.TotalHours).ToString() + ":" + TsUp.Minutes.ToString() + ":" + TsUp.Seconds.ToString();
                    clsLine7MonthTot.TotalWaiting = Math.Floor(TsWait.TotalHours).ToString() + ":" + TsWait.Minutes.ToString() + ":" + TsWait.Seconds.ToString();
                    clsLine7MonthTot.MTTF = MTTF.ToString();
                    clsLine7MonthTot.Lambda = lambda.ToString();
                    clsLine7MonthTot.MTTR = MTTR.ToString();
                    clsLine7MonthTot.MU = Mu.ToString();

                    clsLine7MonthTotal.Add(clsLine7MonthTot);

                }
                else if (i == 4) // Line4
                {
                    clsLine7MonthTot.Line = "Line" + i.ToString();
                    TsDown = objGlobalData.Line4Up - objGlobalData.Line4Down;
                    TsUp = objGlobalData.TotalTime - TsDown;
                    TsWait = objGlobalData.Line4Waiting;


                    double MTTF = getRemaining(TsUp, objGlobalData.Line4DownK);
                    double lambda = Math.Round(1 / MTTF, 5);

                    double MTTR = getRemaining(TsDown, objGlobalData.Line4UpK);
                    double Mu = Math.Round(1 / MTTR, 5);

                    clsLine7MonthTot.TotalDown = Math.Floor(TsDown.TotalHours).ToString() + ":" + TsDown.Minutes.ToString() + ":" + TsDown.Seconds.ToString();
                    clsLine7MonthTot.TotalUp = Math.Floor(TsUp.TotalHours).ToString() + ":" + TsUp.Minutes.ToString() + ":" + TsUp.Seconds.ToString();
                    clsLine7MonthTot.TotalWaiting = Math.Floor(TsWait.TotalHours).ToString() + ":" + TsWait.Minutes.ToString() + ":" + TsWait.Seconds.ToString();
                    clsLine7MonthTot.MTTF = MTTF.ToString();
                    clsLine7MonthTot.Lambda = lambda.ToString();
                    clsLine7MonthTot.MTTR = MTTR.ToString();
                    clsLine7MonthTot.MU = Mu.ToString();

                    clsLine7MonthTotal.Add(clsLine7MonthTot);

                }
                else if (i == 5) // Line5
                {
                    clsLine7MonthTot.Line = "Line" + i.ToString();
                    TsDown = objGlobalData.Line5Up - objGlobalData.Line5Down;
                    TsUp = objGlobalData.TotalTime - TsDown;
                    TsWait = objGlobalData.Line5Waiting;


                    double MTTF = getRemaining(TsUp, objGlobalData.Line5DownK);
                    double lambda = Math.Round(1 / MTTF, 5);

                    double MTTR = getRemaining(TsDown, objGlobalData.Line5UpK);
                    double Mu = Math.Round(1 / MTTR, 5);

                    clsLine7MonthTot.TotalDown = Math.Floor(TsDown.TotalHours).ToString() + ":" + TsDown.Minutes.ToString() + ":" + TsDown.Seconds.ToString();
                    clsLine7MonthTot.TotalUp = Math.Floor(TsUp.TotalHours).ToString() + ":" + TsUp.Minutes.ToString() + ":" + TsUp.Seconds.ToString();
                    clsLine7MonthTot.TotalWaiting = Math.Floor(TsWait.TotalHours).ToString() + ":" + TsWait.Minutes.ToString() + ":" + TsWait.Seconds.ToString();
                    clsLine7MonthTot.MTTF = MTTF.ToString();
                    clsLine7MonthTot.Lambda = lambda.ToString();
                    clsLine7MonthTot.MTTR = MTTR.ToString();
                    clsLine7MonthTot.MU = Mu.ToString();

                    clsLine7MonthTotal.Add(clsLine7MonthTot);

                }
                else if (i == 6) // Line6
                {
                    clsLine7MonthTot.Line = "Line" + i.ToString();
                    TsDown = objGlobalData.Line6Up - objGlobalData.Line6Down;
                    TsUp = objGlobalData.TotalTime - TsDown;
                    TsWait = objGlobalData.Line6Waiting;


                    double MTTF = getRemaining(TsUp, objGlobalData.Line6DownK);
                    double lambda = Math.Round(1 / MTTF, 5);

                    double MTTR = getRemaining(TsDown, objGlobalData.Line6UpK);
                    double Mu = Math.Round(1 / MTTR, 5);

                    clsLine7MonthTot.TotalDown = Math.Floor(TsDown.TotalHours).ToString() + ":" + TsDown.Minutes.ToString() + ":" + TsDown.Seconds.ToString();
                    clsLine7MonthTot.TotalUp = Math.Floor(TsUp.TotalHours).ToString() + ":" + TsUp.Minutes.ToString() + ":" + TsUp.Seconds.ToString();
                    clsLine7MonthTot.TotalWaiting = Math.Floor(TsWait.TotalHours).ToString() + ":" + TsWait.Minutes.ToString() + ":" + TsWait.Seconds.ToString();
                    clsLine7MonthTot.MTTF = MTTF.ToString();
                    clsLine7MonthTot.Lambda = lambda.ToString();
                    clsLine7MonthTot.MTTR = MTTR.ToString();
                    clsLine7MonthTot.MU = Mu.ToString();

                    clsLine7MonthTotal.Add(clsLine7MonthTot);

                }
                //else if (i == 7) // Line7
                //{
                //    clsLine7MonthTot.Line = "Line" + i.ToString();
                //    TsDown = objGlobalData.Line7Up - objGlobalData.Line7Down;
                //    TsUp = objGlobalData.TotalTime - TsDown;
                //    TsWait = objGlobalData.Line7Waiting;


                //    double MTTF = getRemaining(TsUp, objGlobalData.Line7DownK);
                //    double lambda = Math.Round(1 / MTTF, 5);

                //    double MTTR = getRemaining(TsDown, objGlobalData.Line7UpK);
                //    double Mu = Math.Round(1 / MTTR, 5);

                //    clsLine7MonthTot.TotalDown = Math.Floor(TsDown.TotalHours).ToString() + ":" + TsDown.Minutes.ToString() + ":" + TsDown.Seconds.ToString();
                //    clsLine7MonthTot.TotalUp = Math.Floor(TsUp.TotalHours).ToString() + ":" + TsUp.Minutes.ToString() + ":" + TsUp.Seconds.ToString();
                //    clsLine7MonthTot.TotalWaiting = Math.Floor(TsWait.TotalHours).ToString() + ":" + TsWait.Minutes.ToString() + ":" + TsWait.Seconds.ToString();
                //    clsLine7MonthTot.MTTF = MTTF.ToString();
                //    clsLine7MonthTot.Lambda = lambda.ToString();
                //    clsLine7MonthTot.MTTR = MTTR.ToString();
                //    clsLine7MonthTot.MU = Mu.ToString();

                //    clsLine7MonthTotal.Add(clsLine7MonthTot);

                //}

            }
            dataGridView3.DataSource = clsLine7MonthTotal;
            dataGridView3.Visible = true;
        }

        public double getRemaining(TimeSpan UPtime, double count)
        {

            int l1dtimehr = Convert.ToInt32(Math.Floor(UPtime.TotalHours));
            int l1dtimemin = Convert.ToInt32(UPtime.Minutes);
            int l1dtimesec = Convert.ToInt32(UPtime.Seconds);


            double mttfhr, mttfsec, mttfmin = 0.0;

            mttfhr = l1dtimehr / count;
            mttfmin = l1dtimemin / (60 * count);
            mttfsec = l1dtimesec / (3600 * count);

            double mttf = mttfhr + mttfmin + mttfsec;
            return Math.Round(mttf, 5);

        }
        
        private void Form2_Load(object sender, EventArgs e)
        {
            textBox1.Text = kremaining.ToString();
            cmbAuto.SelectedIndex = 0;
        }

        Random r = new Random(); int range = 1;
        List<YearSimulation> LineY = new List<YearSimulation>();
        List<YearSimDisplay> LineYResult = new List<YearSimDisplay>();
        private void btnYear_Click(object sender, EventArgs e)
        {
            if (clsLine7MonthTotal.Count <= 0)
            {
                MessageBox.Show("Please import and simulate the excel data first");
                return;
            }


            int Ycount = 1026; LineYResult = new List<YearSimDisplay>();

            for (int i = 1; i <= 6; i++)
            {
                double MTTF = 0.0, MTTR = 0.0, Lamda = 0.0, MU = 0.0;

                MTTF = Convert.ToDouble(clsLine7MonthTotal[i - 1].MTTF);
                MTTR = Convert.ToDouble(clsLine7MonthTotal[i - 1].MTTR);
                Lamda = Convert.ToDouble(clsLine7MonthTotal[i - 1].Lambda);
                MU = Convert.ToDouble(clsLine7MonthTotal[i - 1].MU);

                LineY = new List<YearSimulation>();
                YearSimDisplay YResult = new YearSimDisplay();

                double PrevArrivalTime = 0.0, PrevCompletionTime = 0.0, CumInterArrival = 0.0, CumServiceTime = 0.0;
                double YMTTF = 0.0, YMTTR = 0.0, YLamda = 0.0, YMU = 0.0;
                for (int J = 0; J < Ycount; J++)
                {
                    double rand = r.NextDouble() * range;
                    YearSimulation Ysim = new YearSimulation();
                    Ysim.Iteration = J + 1;
                    Ysim.InterArrivalTime = Math.Round((-((System.Math.Log(1 - rand)) / Convert.ToDouble(Lamda))), 5);
                    Ysim.ServiceTime = Math.Round((-((System.Math.Log(1 - rand)) / Convert.ToDouble(MU))), 5);
                    CumInterArrival = CumInterArrival + Ysim.InterArrivalTime;
                    CumServiceTime = CumServiceTime + Ysim.ServiceTime;
                    Ysim.ArrivalTime = Math.Round(PrevArrivalTime + Ysim.InterArrivalTime, 5);
                    PrevArrivalTime = Ysim.ArrivalTime;
                    if (J == 0)
                    {
                        Ysim.ServiceStartTime = Math.Round(Ysim.ArrivalTime, 5);
                    }
                    else
                    {
                        if (PrevCompletionTime > Ysim.ArrivalTime)
                            Ysim.ServiceStartTime = Math.Round(PrevCompletionTime, 5);
                        else
                            Ysim.ServiceStartTime = Math.Round(Ysim.ArrivalTime, 5);
                    }
                    Ysim.WaitingTime = Math.Round(Ysim.ServiceStartTime - Ysim.ArrivalTime, 5);
                    Ysim.CompleteTime = Math.Round(Ysim.ServiceStartTime + Ysim.ServiceTime, 5);
                    PrevCompletionTime = Math.Round(Ysim.CompleteTime, 5);
                    Ysim.TimeinSystem = Math.Round(Ysim.CompleteTime - Ysim.ArrivalTime, 5);

                    LineY.Add(Ysim);
                }

                YMTTF = (Math.Round(CumInterArrival / Ycount, 5));
                YMTTR = (Math.Round(CumServiceTime / Ycount, 5));
                YLamda = Math.Round(1 / YMTTF, 5);
                YMU = Math.Round(1 / YMTTR, 5);

                if (i == 1)
                {

                    lbl1MMTTF.Text = clsLine7MonthTotal[i - 1].MTTF;
                    lbl1MMTTR.Text = clsLine7MonthTotal[i - 1].MTTR;
                    lbl1MLamda.Text = clsLine7MonthTotal[i - 1].Lambda;
                    lbl1MMu.Text = clsLine7MonthTotal[i - 1].MU;


                    lbl1YMTTF.Text = YMTTF.ToString();
                    lbl1YMTTR.Text = YMTTR.ToString();
                    lbl1YLamda.Text = YLamda.ToString();
                    lbl1YMu.Text = YMU.ToString();

                    dgvLine1.DataSource = LineY;

                }
                else if (i == 2)
                {
                    lbl2MMTTF.Text = clsLine7MonthTotal[i - 1].MTTF;
                    lbl2MMTTR.Text = clsLine7MonthTotal[i - 1].MTTR;
                    lbl2MLamda.Text = clsLine7MonthTotal[i - 1].Lambda;
                    lbl2MMu.Text = clsLine7MonthTotal[i - 1].MU;


                    lbl2YMTTF.Text = YMTTF.ToString();
                    lbl2YMTTR.Text = YMTTR.ToString();
                    lbl2YLamda.Text = YLamda.ToString();
                    lbl2YMu.Text = YMU.ToString();

                    dgvLine2.DataSource = LineY;

                }
                else if (i == 3)
                {
                    lbl3MMTTF.Text = clsLine7MonthTotal[i - 1].MTTF;
                    lbl3MMTTR.Text = clsLine7MonthTotal[i - 1].MTTR;
                    lbl3MLamda.Text = clsLine7MonthTotal[i - 1].Lambda;
                    lbl3MMu.Text = clsLine7MonthTotal[i - 1].MU;

                    lbl3YMTTF.Text = YMTTF.ToString();
                    lbl3YMTTR.Text = YMTTR.ToString();
                    lbl3YLamda.Text = YLamda.ToString();
                    lbl3YMu.Text = YMU.ToString();

                    dgvLine3.DataSource = LineY;
                }
                else if (i == 4)
                {
                    lbl4MMTTF.Text = clsLine7MonthTotal[i - 1].MTTF;
                    lbl4MMTTR.Text = clsLine7MonthTotal[i - 1].MTTR;
                    lbl4MLamda.Text = clsLine7MonthTotal[i - 1].Lambda;
                    lbl4MMu.Text = clsLine7MonthTotal[i - 1].MU;

                    lbl4YMTTF.Text = YMTTF.ToString();
                    lbl4YMTTR.Text = YMTTR.ToString();
                    lbl4YLamda.Text = YLamda.ToString();
                    lbl4YMu.Text = YMU.ToString();

                    dgvLine4.DataSource = LineY;

                }
                else if (i == 5)
                {
                    lbl5MMTTF.Text = clsLine7MonthTotal[i - 1].MTTF;
                    lbl5MMTTR.Text = clsLine7MonthTotal[i - 1].MTTR;
                    lbl5MLamda.Text = clsLine7MonthTotal[i - 1].Lambda;
                    lbl5MMu.Text = clsLine7MonthTotal[i - 1].MU;

                    lbl5YMTTF.Text = YMTTF.ToString();
                    lbl5YMTTR.Text = YMTTR.ToString();
                    lbl5YLamda.Text = YLamda.ToString();
                    lbl5YMu.Text = YMU.ToString();

                    dgvLine5.DataSource = LineY;

                }
                else if (i == 6)
                {
                    lbl6MMTTF.Text = clsLine7MonthTotal[i - 1].MTTF;
                    lbl6MMTTR.Text = clsLine7MonthTotal[i - 1].MTTR;
                    lbl6MLamda.Text = clsLine7MonthTotal[i - 1].Lambda;
                    lbl6MMu.Text = clsLine7MonthTotal[i - 1].MU;

                    lbl6YMTTF.Text = YMTTF.ToString();
                    lbl6YMTTR.Text = YMTTR.ToString();
                    lbl6YLamda.Text = YLamda.ToString();
                    lbl6YMu.Text = YMU.ToString();

                    dgvLine6.DataSource = LineY;
                }
                //else if (i == 7)
                //{
                //    lbl7MMTTF.Text = clsLine7MonthTotal[i - 1].MTTF;
                //    lbl7MMTTR.Text = clsLine7MonthTotal[i - 1].MTTR;
                //    lbl7MLamda.Text = clsLine7MonthTotal[i - 1].Lambda;
                //    lbl7MMu.Text = clsLine7MonthTotal[i - 1].MU;

                //    lbl7YMTTF.Text = YMTTF.ToString();
                //    lbl7YMTTR.Text = YMTTR.ToString();
                //    lbl7YLamda.Text = YLamda.ToString();
                //    lbl7YMu.Text = YMU.ToString();

                //    dgvLine7.DataSource = LineY;

                //}
                //else if (i == 8)
                //{
                //    lbl8MMTTF.Text = clsLine7MonthTotal[i - 1].MTTF;
                //    lbl8MMTTR.Text = clsLine7MonthTotal[i - 1].MTTR;
                //    lbl8MLamda.Text = clsLine7MonthTotal[i - 1].Lambda;
                //    lbl8MMu.Text = clsLine7MonthTotal[i - 1].MU;

                //    lbl8YMTTF.Text = YMTTF.ToString();
                //    lbl8YMTTR.Text = YMTTR.ToString();
                //    lbl8YLamda.Text = YLamda.ToString();
                //    lbl8YMu.Text = YMU.ToString();

                //    dgvLine8.DataSource = LineY;

                //}
                //else if (i == 9)
                //{
                //    lbl9MMTTF.Text = clsLine7MonthTotal[i - 1].MTTF;
                //    lbl9MMTTR.Text = clsLine7MonthTotal[i - 1].MTTR;
                //    lbl9MLamda.Text = clsLine7MonthTotal[i - 1].Lambda;
                //    lbl9MMu.Text = clsLine7MonthTotal[i - 1].MU;

                //    lbl9YMTTF.Text = YMTTF.ToString();
                //    lbl9YMTTR.Text = YMTTR.ToString();
                //    lbl9YLamda.Text = YLamda.ToString();
                //    lbl9YMu.Text = YMU.ToString();

                //    dgvLine9.DataSource = LineY;

                //}
                //else if (i == 10)
                //{
                //    lbl10MMTTF.Text = clsLine7MonthTotal[i - 1].MTTF;
                //    lbl10MMTTR.Text = clsLine7MonthTotal[i - 1].MTTR;
                //    lbl10MLamda.Text = clsLine7MonthTotal[i - 1].Lambda;
                //    lbl10MMu.Text = clsLine7MonthTotal[i - 1].MU;

                //    lbl10YMTTF.Text = YMTTF.ToString();
                //    lbl10YMTTR.Text = YMTTR.ToString();
                //    lbl10YLamda.Text = YLamda.ToString();
                //    lbl10YMu.Text = YMU.ToString();

                //    dgvLine10.DataSource = LineY;

                //}

                YResult.Line = "Line" + i.ToString();
                YResult.YearMTTF = YMTTF;
                YResult.YearMTTR = YMTTR;
                YResult.YearLamda = YLamda;
                YResult.YearMU = YMU;
                YResult.MonthMTTF = MTTF;
                YResult.MonthMTTR = MTTR;
                YResult.MonthLamda = Lamda;
                YResult.MonthMU = MU;

                LineYResult.Add(YResult);

            }
            dgvYear.DataSource = LineYResult;
        }
    }
}
