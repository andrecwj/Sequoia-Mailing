using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Mailing_Label
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string num = numericUpDown1.Value.ToString();

            if (num == "0")
            {
                MessageBox.Show("The last transaction must be at least a month old. Please try again.");
            }
            else
            {
                Form main = new Form1();
                this.Close();
                MessageBox.Show("Changes saved", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);

                System.IO.Directory.CreateDirectory(@"M:\SequoiaPOS\SequoiaMailing");

                string[] lines = { num };
                System.IO.File.WriteAllLines(@"M:\SequoiaPOS\SequoiaMailing\MailingSettings.txt", lines);
            }
        }
    }
}
