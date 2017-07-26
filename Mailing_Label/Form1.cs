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
    }
}
