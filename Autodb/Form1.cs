using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Media;
using System.IO;
using System.Text;
using Microsoft.Win32;
using System.Threading.Tasks;
using System.Windows.Forms;
using crypting = CryptIO;

namespace Autodb
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            FileInfo file = new FileInfo(Application.StartupPath + "/CryptIO.dll");
            if (!file.Exists)
            {
                MessageBox.Show("Динамическая библиотека CryptIO неустановлена. \n Переустановите приложение.");
                Close();
            }
            RegistryKey currentUserKey = Registry.CurrentUser;
            RegistryKey helloKey = currentUserKey.OpenSubKey("AS", true);
            string login = helloKey.GetValue("Activated").ToString();
            if (login != "1")
            {
                MessageBox.Show("Вы не активировали данный программный продукт");
                button1.Enabled = false;
                button2.Enabled = false;
            }
            helloKey.Close();
            notifyIcon1.ShowBalloonTip(12);
        }

        private void Label6_Click(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                db.startConnection("Data Source = " + comboBox2.Text.ToString() + "; " +
                    "Initial Catalog = " + textBox10.Text + "; " +
                    "Persist Security Info = True; User ID = " + textBox9.Text + "; Password=\"" + textBox8.Text + "\"");
            }
            catch (Exception ex) when (ex is NullReferenceException || ex is System.Data.SqlClient.SqlException)
            {
                MessageBox.Show("Невозможно подключиться к серверу БД. Проверьте правильность данных сервера\n");
            }
            string passwordHash = crypting.crypt.getHash(textBox2.Text);

            try
            {
                SqlCommand auth = new SqlCommand("SELECT dbo.authUser('" + textBox1.Text + "', '" + passwordHash + "')", db.connection);
                string result = auth.ExecuteScalar().ToString();
                if (result != "")
                {
                    db.userId = result;
                    //authorized
                    var snd = new SoundPlayer(Properties.Resources.fit);
                    snd.Play();
                    snd.Dispose();
                    this.Hide();
                    MainFormcs mainFormcs = new MainFormcs();
                    mainFormcs.Show();
                    mainFormcs.FormClosed += MainFormcs_FormClosed;
                }
                else
                {
                    var snd = new SoundPlayer(Properties.Resources.fit);
                    snd.Play();
                    snd.Dispose();
                    MessageBox.Show("Ошибка. Проверьте введеные данные");                    
                }                
            }
            catch(Exception ex)
            {
                MessageBox.Show("Ошибка. Проверьте введеные данные" + ex.Message);
            }            
        }

        private void MainFormcs_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Show();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            db.fillServers(comboBox2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                db.startConnection("Data Source = " + comboBox2.Text.ToString() + "; " +
                    "Initial Catalog = " + textBox10.Text + "; " +
                    "Persist Security Info = True; User ID = " + textBox9.Text + "; Password=\"" + textBox8.Text + "\"");
                try
                {
                    string result = new SqlCommand("SELECT top(1) ID_ROLE from dbo.ROLE_ where GUEST>0", db.connection).ExecuteScalar().ToString();
                    
                    new SqlCommand("EXEC INSERT_USERS '" + textBox4.Text + "', '" + crypting.crypt.getHash(textBox3.Text) + "', '" + result + "'", db.connection).ExecuteNonQuery();
                    new SqlCommand("EXEC INSERT_SOTR_1 '" + textBox5.Text + "', '" + textBox6.Text + "', '" + textBox7.Text + "', '"
                        + maskedTextBox1.Text + "', '" + result + "'", db.connection).ExecuteScalar();
                    var snd = new SoundPlayer(Properties.Resources.fit);
                    snd.Play();
                    snd.Dispose();
                    MessageBox.Show("Вы зарагестрированы");
                }
                catch(Exception ex) when (ex is NullReferenceException || ex is System.Data.SqlClient.SqlException)
                {
                    MessageBox.Show("Невозможно создать учетную запись" + ex.ToString());
                }
            }
            catch (Exception ex) when (ex is NullReferenceException || ex is System.Data.SqlClient.SqlException)
            {
                MessageBox.Show("Невозможно подключиться к серверу БД. Проверьте правильность данных сервера\n" + ex.ToString());
            }            
            
        }
    }
}
