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



namespace Mailing_Label
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            fillcombo();
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
                    //createCommand.ExecuteNonQuery();
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
                MessageBox.Show("The selected item on the left is " + v + ", which is lower in the descending alphabetical order. Please edit your choice and try again.");
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
                this.comboBox3.Items.Clear();
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
                    MessageBox.Show("Error: " + ex + " Please try again.");
                }
            }
            else if (checkBox1.Checked == false)
            {
                this.comboBox3.Items.Clear();
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
                    MessageBox.Show("Error: " + ex + " Please try again.");
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

            if (u > 0 && v != "" )
            {
                MessageBox.Show("The selected item on the right is " + v + ", which is lower in the descending alphabetical order. Please edit your choice and try again.");
            }
        }
    }
}

