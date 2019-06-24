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
using _crypt = CryptIO;

namespace Autodb
{
    public partial class SettingsWindow : Form
    {
        public SettingsWindow()
        {
            InitializeComponent();
            textBox1.Text = Properties.Settings.Default.defaultPath;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            Properties.Settings.Default.defaultPath = folderBrowserDialog1.SelectedPath;
            Properties.Settings.Default.Save();
            textBox1.Text = Properties.Settings.Default.defaultPath;
        }

        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.color = "White";
        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.color = "Gray";
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            SqlCommand sqlCommand = new SqlCommand("Update USERS_ set USERS_PASS = '"+textBox3.Text+"' where USERS_PASS = '"+textBox2.Text+"'");
            sqlCommand.ExecuteNonQuery();
        }
    }
    
}
