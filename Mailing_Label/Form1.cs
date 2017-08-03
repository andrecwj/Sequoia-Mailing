using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;



namespace Mailing_Label
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            fillcombo();
            comboBox3.SelectedIndex = 0;
        }


        private class Item
        {
            public string Name;
            public int Value;
            public Item(string name, int value)
            {
                Name = name;
                Value = value;
            }
        }

        string connectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=SgSeqDemo;Persist Security Info=True;User ID=sa;Password=bigtree";

        void fillcombo()
        {
            SqlConnection sqcon = new SqlConnection(connectionString);
            try
            {
                sqcon.Open();
                if (checkBox1.Checked == false)
                {
                    string Query = "Select DISTINCT Class_Desc FROM Customer_Class";
                    SqlCommand createCommand = new SqlCommand(Query, sqcon);
                    SqlDataReader dr = createCommand.ExecuteReader();
                    while (dr.Read())
                    {
                        string name = dr.GetString(0);
                        comboBox3.Items.Add(name);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex + " Please try again.");
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_Click(object sender, EventArgs e)
        {
            {
                try
                {
                    string box = comboBox1.SelectedItem.ToString();
                }
                catch
                {
                    MessageBox.Show("Select an option on the left drop-down box first.");
                }
            }
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            string s = comboBox1.Text;
            string t = comboBox2.Text;

            string v = comboBox1.Text;

            int u = string.CompareOrdinal(s, t);

            if (u > 0)
            {
                MessageBox.Show("The selected item on the left is " + v + ". Only alphabets after " + v + " are allowed to be entered on the right. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                string name2 = "ALL";
                this.comboBox3.Items.Clear();
                comboBox3.Items.Add(name2);
                comboBox3.SelectedIndex = 0;

                SqlConnection sqcon = new SqlConnection(connectionString);
                try
                {
                    sqcon.Open();
                    string Query = "Select DISTINCT Class_Desc FROM Customer_Class WHERE Class_Isactive='True'";
                    SqlCommand createCommand = new SqlCommand(Query, sqcon);
                    SqlDataReader dr = createCommand.ExecuteReader();
                    while (dr.Read())
                    {
                        string name = dr.GetString(0);
                        comboBox3.Items.Add(name);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occured. Please try again, or contact support for assistance. Error: " + ex);
                }
            }
            else if (checkBox1.Checked == false)
            {
                string name2 = "ALL";
                this.comboBox3.Items.Clear();
                comboBox3.Items.Add(name2);
                comboBox3.SelectedIndex = 0;

                SqlConnection sqcon = new SqlConnection(connectionString);
                try
                {
                    sqcon.Open();
                    string Query = "Select DISTINCT Class_Desc FROM Customer_Class";
                    SqlCommand createCommand = new SqlCommand(Query, sqcon);
                    SqlDataReader dr = createCommand.ExecuteReader();
                    while (dr.Read())
                    {
                        string name = dr.GetString(0);
                        comboBox3.Items.Add(name);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occured. Please try again, or contact support for assistance. Error: " + ex);
                }
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            string w = comboBox1.Text;
            string x = comboBox2.Text;

            string v = comboBox2.Text;

            int u = string.CompareOrdinal(w, x);

            if (v == "")
            {
                string box2 = comboBox1.SelectedItem.ToString();
            }

            if (u > 0 && v != "")
            {
                MessageBox.Show("The selected item on the right is " + v + ". Only alphabets after " + v + " are allowed to be entered on the left. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }


        private void preferencesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form cs = new Settings();
            cs.Show();

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("An error occured. Please try again, or contact support for assistance. Error: " + ex);
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string cb1 = comboBox1.Text;
            string cb2 = comboBox2.Text;
            string cb3 = comboBox3.Text;

            if (cb3 == "ALL" && checkBox1.Checked == false)
            {
                try
                {
                    string sql = null;
                    string data = null;
                    string data1 = null;
                    int i = 0;
                    int j = 0;
                    int k = 0;
                    int l = 0;

                    Excel.Application xlApp;
                    Excel.Workbook xlWorkbook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;

                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

                    SqlConnection sqcon = new SqlConnection(connectionString);
                    sqcon.Open();
                    sql = "Select Cust_name FROM Customer WHERE Cust_name BETWEEN'" + cb1 + "'AND'" + cb2 + "ZZZZZZZZ' ORDER BY Cust_name ASC";
                    SqlDataAdapter dscmd = new SqlDataAdapter(sql, sqcon);
                    DataSet ds = new DataSet();
                    dscmd.Fill(ds);

                    for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                    {
                        for (j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
                        {
                            data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                            xlWorkSheet.Cells[i + 1, j + 1] = data;
                        }
                    }

                    xlWorkbook.SaveAs(@"M:\SequoiaPOS\Mailing_Name.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkbook.Close(true, misValue, misValue);

                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

                    SqlConnection sqcon1 = new SqlConnection(connectionString);
                    sqcon1.Open();
                    sql = "select Cust_address, Cust_address1, Cust_PostCode from Customer WHERE Cust_name BETWEEN'" + cb1 + "'AND'" + cb2 + "ZZZZZZZZ' ORDER BY Cust_name ASC";
                    SqlDataAdapter postcode = new SqlDataAdapter(sql, sqcon1);
                    DataSet ds1 = new DataSet();
                    postcode.Fill(ds1);

                    for (k = 0; k <= ds1.Tables[0].Rows.Count - 1; k++)
                    {
                        for (l = 0; l <= ds1.Tables[0].Columns.Count - 1; l++)
                        {
                            data1 = ds1.Tables[0].Rows[k].ItemArray[l].ToString();
                            xlWorkSheet.Cells[k + 1, l + 1] = data1;
                        }
                    }

                    xlWorkbook.SaveAs(@"M:\SequoiaPOS\Mailing_Address_Postcode.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkbook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkbook);
                    releaseObject(xlApp);
                }

                catch (Exception ex)
                {
                    MessageBox.Show("An error occured. Please try again, or contact support for assistance. Error: " + ex);
                }
            }

            else if (cb3 == "ALL" && checkBox2.Checked == false)
            {
                try
                {
                    string sql = null;
                    string sql1 = null;
                    string data = null;
                    string data1 = null;
                    int i = 0;
                    int j = 0;
                    int k = 0;
                    int l = 0;

                    Excel.Application xlApp;
                    Excel.Workbook xlWorkbook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;

                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

                    SqlConnection sqcon = new SqlConnection(connectionString);
                    sqcon.Open();
                    sql = "Select Cust_name FROM Customer WHERE Cust_name BETWEEN'" + cb1 + "'AND'" + cb2 + "ZZZZZZZZ' ORDER BY Cust_name ASC";
                    SqlDataAdapter dscmd = new SqlDataAdapter(sql, sqcon);
                    DataSet ds = new DataSet();
                    dscmd.Fill(ds);

                    for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                    {
                        for (j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
                        {
                            data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                            xlWorkSheet.Cells[i + 1, j + 1] = data;
                        }
                    }

                    xlWorkbook.SaveAs(@"M:\SequoiaPOS\Mailing_Name.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkbook.Close(true, misValue, misValue);

                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

                    SqlConnection sqcon1 = new SqlConnection(connectionString);
                    sqcon1.Open();
                    sql1 = "select Cust_address, Cust_address1, Cust_PostCode from Customer WHERE Cust_name BETWEEN'" + cb1 + "'AND'" + cb2 + "ZZZZZZZZ' ORDER BY Cust_name ASC";
                    SqlDataAdapter postcode = new SqlDataAdapter(sql1, sqcon1);
                    DataSet ds1 = new DataSet();
                    postcode.Fill(ds1);

                    for (k = 0; k <= ds1.Tables[0].Rows.Count - 1; k++)
                    {
                        for (l = 0; l <= ds1.Tables[0].Columns.Count - 1; l++)
                        {
                            data1 = ds1.Tables[0].Rows[k].ItemArray[l].ToString();
                            xlWorkSheet.Cells[k + 1, l + 1] = data1;
                        }
                    }

                    xlWorkbook.SaveAs(@"M:\SequoiaPOS\Mailing_Address_Postcode.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkbook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkbook);
                    releaseObject(xlApp);

                }

                catch (Exception ex)
                {
                    MessageBox.Show("An error occured. Please try again, or contact support for assistance. Error: " + ex);
                }
            }

            else if (cb3 == "ALL" && checkBox2.Checked == true)
            {
                try
                {
                    string sql = null;
                    string sql1 = null;
                    string data = null;
                    string data1 = null;
                    int i = 0;
                    int j = 0;
                    int k = 0;
                    int l = 0;

                    Excel.Application xlApp;
                    Excel.Workbook xlWorkbook;
                    Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;

                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

                    SqlConnection sqcon = new SqlConnection(connectionString);
                    sqcon.Open();
                    sql = "Select Cust_name FROM Customer WHERE Cust_name BETWEEN'" + cb1 + "'AND'" + cb2 + "ZZZZZZZZ' AND Cust_isactive = 'True' ORDER BY Cust_name ASC";
                    SqlDataAdapter dscmd = new SqlDataAdapter(sql, sqcon);
                    DataSet ds = new DataSet();
                    dscmd.Fill(ds);

                    for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                    {
                        for (j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
                        {
                            data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                            xlWorkSheet.Cells[i + 1, j + 1] = data;
                        }
                    }

                    xlWorkbook.SaveAs(@"M:\SequoiaPOS\Mailing_Name.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkbook.Close(true, misValue, misValue);
                    xlApp = new Microsoft.Office.Interop.Excel.Application();
                    xlWorkbook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

                    SqlConnection sqcon1 = new SqlConnection(connectionString);
                    sqcon1.Open();
                    sql1 = "select Cust_address, Cust_address1, Cust_PostCode from Customer WHERE Cust_name BETWEEN '" + cb1 + "' AND '" + cb2 + "ZZZZZZZZ' AND Cust_isactive = 'True' ORDER BY Cust_name ASC";
                    SqlDataAdapter postcode = new SqlDataAdapter(sql1, sqcon1);
                    DataSet ds1 = new DataSet();
                    postcode.Fill(ds1);

                    for (k = 0; k <= ds1.Tables[0].Rows.Count - 1; k++)
                    {
                        for (l = 0; l <= ds1.Tables[0].Columns.Count - 1; l++)
                        {
                            data1 = ds1.Tables[0].Rows[k].ItemArray[l].ToString();
                            xlWorkSheet.Cells[k + 1, l + 1] = data1;
                        }
                    }

                    xlWorkbook.SaveAs(@"M:\SequoiaPOS\Mailing_Address_Postcode.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkbook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkbook);
                    releaseObject(xlApp);
                }

                catch (Exception ex)
                {
                    MessageBox.Show("An error occured. Please try again, or contact support for assistance. Error: " + ex);
                }
            }

            else
            {
                if (checkBox2.Checked == false)
                {
                    try
                    {
                        string sql = null;
                        string sql1 = null;
                        string data = null;
                        string data1 = null;
                        int i = 0;
                        int j = 0;
                        int k = 0;
                        int l = 0;

                        Excel.Application xlApp;
                        Excel.Workbook xlWorkbook;
                        Excel.Worksheet xlWorkSheet;
                        object misValue = System.Reflection.Missing.Value;

                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkbook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

                        SqlConnection sqcon = new SqlConnection(connectionString);
                        sqcon.Open();
                        sql = "Select Customer.Cust_name, Customer_Class.Class_Desc FROM Customer INNER JOIN Customer_Class ON Customer.Cust_Class=Customer_Class.Class_Code WHERE Cust_name BETWEEN '" + cb1 + "' AND '" + cb2 + "ZZZZZZZZ' AND Customer_Class.Class_Desc='" + cb3 + "' ORDER BY Cust_name ASC";
                        SqlDataAdapter dscmd = new SqlDataAdapter(sql, sqcon);
                        DataSet ds = new DataSet();
                        dscmd.Fill(ds);

                        for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            for (j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
                            {
                                data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                                xlWorkSheet.Cells[i + 1, j + 1] = data;
                            }
                        }

                        xlWorkbook.SaveAs(@"M:\SequoiaPOS\Mailing_Name.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkbook.Close(true, misValue, misValue);

                        SqlConnection sqcon1 = new SqlConnection(connectionString);
                        sqcon1.Open();
                        sql1 = "select Cust_address, Cust_address1, Cust_PostCode from Customer INNER JOIN Customer_Class ON Customer.Cust_Class=Customer_Class.Class_Code WHERE Cust_name BETWEEN '" + cb1 + "' AND '" + cb2 + "ZZZZZZZZ' AND Customer_Class.Class_Desc='" + cb3 + "' ORDER BY Cust_name ASC";
                        SqlDataAdapter postcode = new SqlDataAdapter(sql1, sqcon1);
                        DataSet ds1 = new DataSet();
                        postcode.Fill(ds1);

                        for (k = 0; k <= ds1.Tables[0].Rows.Count - 1; k++)
                        {
                            for (l = 0; l <= ds1.Tables[0].Columns.Count - 1; l++)
                            {
                                data1 = ds1.Tables[0].Rows[k].ItemArray[l].ToString();
                                xlWorkSheet.Cells[k + 1, l + 1] = data1;
                            }
                        }

                        xlWorkbook.SaveAs(@"M:\SequoiaPOS\Mailing_Address_Postcode.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkbook.Close(true, misValue, misValue);
                        xlApp.Quit();

                        releaseObject(xlWorkSheet);
                        releaseObject(xlWorkbook);
                        releaseObject(xlApp);
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("An error occured. Please try again, or contact support for assistance. Error: " + ex);
                    }
                }

                if (checkBox2.Checked == true)
                {
                    try
                    {
                        string sql = null;
                        string sql1 = null;
                        string data = null;
                        string data1 = null;
                        int i = 0;
                        int j = 0;
                        int k = 0;
                        int l = 0;

                        Excel.Application xlApp;
                        Excel.Workbook xlWorkbook;
                        Excel.Worksheet xlWorkSheet;
                        object misValue = System.Reflection.Missing.Value;

                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkbook = xlApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

                        SqlConnection sqcon = new SqlConnection(connectionString);
                        sqcon.Open();
                        sql = "Select Customer.Cust_name, Customer_Class.Class_Desc FROM Customer INNER JOIN Customer_Class ON Customer.Cust_Class=Customer_Class.Class_Code WHERE Cust_name BETWEEN '" + cb1 + "' AND '" + cb2 + "ZZZZZZZZ' AND Customer_Class.Class_Desc='" + cb3 + "' AND Cust_isactive = 'True' ORDER BY Cust_name ASC";
                        SqlDataAdapter dscmd = new SqlDataAdapter(sql, sqcon);
                        DataSet ds = new DataSet();
                        dscmd.Fill(ds);

                        for (i = 0; i <= ds.Tables[0].Rows.Count - 1; i++)
                        {
                            for (j = 0; j <= ds.Tables[0].Columns.Count - 1; j++)
                            {
                                data = ds.Tables[0].Rows[i].ItemArray[j].ToString();
                                xlWorkSheet.Cells[i + 1, j + 1] = data;
                            }
                        }

                        xlWorkbook.SaveAs(@"M:\SequoiaPOS\Mailing_Name.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkbook.Close(true, misValue, misValue);
                        SqlConnection sqcon1 = new SqlConnection(connectionString);
                        sqcon1.Open();
                        sql1 = "select Cust_address, Cust_address1, Cust_PostCode from Customer INNER JOIN Customer_Class ON Customer.Cust_Class=Customer_Class.Class_Code WHERE Cust_name BETWEEN '" + cb1 + "' AND '" + cb2 + "ZZZZZZZZ' AND Customer_Class.Class_Desc='" + cb3 + "' AND Cust_isactive = 'True' ORDER BY Cust_name ASC";
                        SqlDataAdapter postcode = new SqlDataAdapter(sql1, sqcon1);
                        DataSet ds1 = new DataSet();
                        postcode.Fill(ds1);

                        for (k = 0; k <= ds1.Tables[0].Rows.Count - 1; k++)
                        {
                            for (l = 0; l <= ds1.Tables[0].Columns.Count - 1; l++)
                            {
                                data1 = ds1.Tables[0].Rows[k].ItemArray[l].ToString();
                                xlWorkSheet.Cells[k + 1, l + 1] = data1;
                            }
                        }

                        xlWorkbook.SaveAs(@"M:\SequoiaPOS\Mailing_Address_Postcode.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkbook.Close(true, misValue, misValue);
                        xlApp.Quit();

                        releaseObject(xlWorkSheet);
                        releaseObject(xlWorkbook);
                        releaseObject(xlApp);


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("An error occured. Please try again, or contact support for assistance. Error: " + ex);
                    }
                }
            }
        }
    }
}

