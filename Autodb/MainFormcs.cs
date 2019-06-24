using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Data.SqlClient;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace Autodb
{
    public partial class MainFormcs : Form
    {
        private void update()
        {
            ZakTable.DataSource = _Tables.table
                (
                    "select " +
                    "zak.id_klient as 'КлиентCODE', " +
                    "klient_fam as 'Клиент', " +
                    "zak.id_machine as 'МашинаCODE', " +
                    "nom_machine as 'Машина', " +
                    "zak.id_sotr as 'СотрудникCODE', " +
                    "fam_sotr as 'Сотрудник', " +
                    "zak.id_uslugi as 'УслугаCODE', " +
                    "uslugi_naim as 'Услуга', " +
                    "zak_summ as 'Сумма', " +
                    "date_zak as 'Дата' " +
                    "from zak  " +
                    "inner join klient on zak.id_klient = klient.ID_KLIENT " +
                    "inner join machine on zak.ID_MACHINE = MACHINE.ID_MACHINE " +
                    "inner join sotr on zak.ID_SOTR = sotr.ID_SOTR " +
                    "inner join USLUGI on zak.ID_USLUGI = USLUGI.ID_USLUGI"
                );
            ZakTable.Columns["КлиентCODE"].Visible = false;
            ZakTable.Columns["МашинаCODE"].Visible = false;
            ZakTable.Columns["СотрудникCODE"].Visible = false;
            ZakTable.Columns["УслугаCODE"].Visible = false;

            dataGridView1.DataSource = _Tables.table
                (
                    "select "+
                    "id_machine, "+
                    "nom_machine, "+
                    "machine.MODEL_ID, "+
                    "model.MARK_ID, "+
                    "model.MODEL_NAIM, "+
                    "mark.MARKA "+
                    "from machine "+
                    "inner join model on machine.MODEL_ID = model.id_model "+
                    "inner join mark on model.MARK_ID = mark.id_mark "
                );

            dataGridView4.DataSource = _Tables.table
                (
                    "select " +
                    "* " +
                    
                    "from klient "
                );

            dataGridView5.DataSource = _Tables.table
                (
                    "select " +
                    "* " +

                    "from dolj "
                );


            comboBox7.DataSource = _Tables.table
                (
                    "Select id_model, MODEL_NAIM from model"
                );
            comboBox7.ValueMember = "id_model";
            comboBox7.DisplayMember = "MODEL_NAIM";

            comboBox6.DataSource = _Tables.table
                (
                    "Select id_mark, MARKA from mark"
                );
            comboBox6.ValueMember = "id_mark";
            comboBox6.DisplayMember = "MARKA";
            

            dataGridView3.DataSource = _Tables.table
                (
                    "select " +
                    "id_zak, " +
                    "zak.id_klient as 'КлиентCODE', " +
                    "klient_fam as 'Клиент', " +
                    "zak.id_machine as 'МашинаCODE', " +
                    "nom_machine as 'Машина', " +
                    "zak.id_sotr as 'СотрудникCODE', " +
                    "fam_sotr as 'Сотрудник', " +
                    "zak.id_uslugi as 'УслугаCODE', " +
                    "uslugi_naim as 'Услуга', " +
                    "zak_summ as 'Сумма', " +
                    "date_zak as 'Дата', " +
                    "stat as 'Статус заказа' " +
                    "from zak  " +
                    "inner join klient on zak.id_klient = klient.ID_KLIENT " +
                    "inner join machine on zak.ID_MACHINE = MACHINE.ID_MACHINE " +
                    "inner join sotr on zak.ID_SOTR = sotr.ID_SOTR " +
                    "inner join USLUGI on zak.ID_USLUGI = USLUGI.ID_USLUGI"
                );
            dataGridView3.Columns["КлиентCODE"].Visible = false;
            dataGridView3.Columns["МашинаCODE"].Visible = false;
            dataGridView3.Columns["СотрудникCODE"].Visible = false;
            dataGridView3.Columns["id_zak"].Visible = false;
            dataGridView3.Columns["УслугаCODE"].Visible = false;
            comboBox1.DataSource = _Tables.table
                (
                    "Select id_klient, klient_fam from klient"
                );
            comboBox1.ValueMember = "id_klient";
            comboBox1.DisplayMember = "klient_fam";
            //----------------------------------------------
            comboBox2.DataSource = _Tables.table
                (
                    "Select id_machine, nom_machine from machine"
                );
            comboBox2.ValueMember = "id_machine";
            comboBox2.DisplayMember = "nom_machine";
            //----------------------------------------------
            comboBox3.DataSource = _Tables.table
                (
                    "Select id_sotr, fam_sotr from sotr"
                );
            comboBox3.ValueMember = "id_sotr";
            comboBox3.DisplayMember = "fam_sotr";
            //----------------------------------------------
            comboBox4.DataSource = _Tables.table
                (
                    "Select id_uslugi, uslugi_naim, uslugi_stoimost from uslugi"
                );
            comboBox4.ValueMember = "id_uslugi";
            comboBox4.DisplayMember = "uslugi_naim";
            foreach (DataRow row in _Tables.table
                (
                    "Select id_uslugi, uslugi_naim, uslugi_stoimost from uslugi"
                ).Rows)
            {
                if (row["id_uslugi"] == comboBox4.SelectedValue)
                {
                    label1.Text = row["uslugi_stoimost"].ToString();
                }
            }
            foreach (DataRow row in _Tables.table
    (
        "Select id_uslugi, uslugi_naim, uslugi_stoimost from uslugi"
    ).Rows)
            {
                //MessageBox.Show(row["id_uslugi"].ToString());
                //MessageBox.Show(comboBox4.SelectedValue.ToString());
                if (row["id_uslugi"].ToString() == comboBox4.SelectedValue.ToString())
                {
                    label1.Text = row["uslugi_stoimost"].ToString();
                }
            }
            dataGridView2.DataSource = _Tables.table("SELECT DATE_OF_LOG as 'Дата', MESSAGE_LOG as 'Описание' from log_table");
            //ZakTable.Rows[0].Selected = true;
            SqlCommand sqlCommand = new SqlCommand("SELECT DATE_OF_LOG as 'Дата', MESSAGE_LOG as 'Описание' from log_table", db.connection);
            SqlDependency sqlDependency = new SqlDependency(sqlCommand);
            sqlDependency.OnChange += SqlDependency_OnChange;
           
        }

        private void SqlDependency_OnChange(object sender, SqlNotificationEventArgs e)
        {
            dataGridView2.Rows.Clear();
            dataGridView2.DataSource = _Tables.table("SELECT DATE_OF_LOG as 'Дата', MESSAGE_LOG as 'Описание' from log_table");
        }

        public MainFormcs()
        {
            InitializeComponent();
            if (Properties.Settings.Default.color == "Gray")
            {
                tabControl1.TabPages[0].BackColor = Color.Navy;
                tabControl1.TabPages[1].BackColor = Color.Navy;
                tabControl1.TabPages[2].BackColor = Color.Navy;
                tabControl1.TabPages[3].BackColor = Color.Navy;
                BackColor = Color.Navy;
            }
            update();

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {
            
        }

        private void ZakTable_SelectionChanged(object sender, EventArgs e)
        {
            //foreach (var row in  comboBox1.Items)
            //{
            //    MessageBox.Show(row.ToString());
            //}
            /*MessageBox.Show(ZakTable.SelectedRows[1].Cells["КлиентCODE"].Value.ToString());

            comboBox1.SelectedValue = ZakTable.SelectedRows[0].Cells["КлиентCODE"].Value.ToString();
            comboBox2.SelectedValue = ZakTable.SelectedRows[0].Cells["МашинаCODE"].Value.ToString();
            comboBox3.SelectedValue = ZakTable.SelectedRows[0].Cells["СотрудникCODE"].Value.ToString();
            comboBox4.SelectedValue = ZakTable.SelectedRows[0].Cells["УслугаCODE"].Value.ToString();*/
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            SqlCommand sqlCommand = new SqlCommand("EXEC INSERT_ZAK "+comboBox1.SelectedValue+","+ comboBox2.SelectedValue +","+ comboBox3.SelectedValue +","+ comboBox4.SelectedValue +","+ label1.Text +",'"+ DateTime.Now.ToString("d") +"'",db.connection);
            //MessageBox.Show(sqlCommand.CommandText);
            sqlCommand.ExecuteNonQuery();
            update();
        }

        private void ComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (DataRow row in _Tables.table
                (
                    "Select id_uslugi, uslugi_naim, uslugi_stoimost from uslugi"
                ).Rows)
            {
                //MessageBox.Show(row["id_uslugi"].ToString());
                //MessageBox.Show(comboBox4.SelectedValue.ToString());
                if (row["id_uslugi"].ToString() == comboBox4.SelectedValue.ToString())
                {
                    label1.Text = row["uslugi_stoimost"].ToString();
                }
            }
        }

        private void ComboBox4_Click(object sender, EventArgs e)
        {
            foreach (DataRow row in _Tables.table
                (
                    "Select id_uslugi, uslugi_naim, uslugi_stoimost from uslugi"
                ).Rows)
            {
                if (row["id_uslugi"] == comboBox4.SelectedValue)
                {
                    label1.Text = row["uslugi_stoimost"].ToString();
                }
            }
        }

        private void TabPage1_Click(object sender, EventArgs e)
        {

        }

        private void wordandpdf(int a, string cb1text, string cb2text, string cb3text, string cb4text, string l1text)
        {
            Word.Application winword = new Word.Application(); //создаем COM-объект Word 
            object missing = System.Reflection.Missing.Value;
            winword.Visible = false;
            Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            document.Content.Text = "Чек" + Environment.NewLine + "Клиент: " + cb1text + Environment.NewLine + "Номер машины: " + cb2text + Environment.NewLine + "Сотрудник: " + cb3text + Environment.NewLine + "Услуга: " + cb4text +
            Environment.NewLine + "Время оформления заказа " + DateTime.Now + Environment.NewLine + "Цена: " + l1text;
            winword.Visible = true;
            Word.Document DocWord = winword.Application.ActiveDocument;
            switch (a)
            {
                case 1:
                    DocWord.SaveAs2(filename);
                    filename = null;
                    break;
                case 2:
                    DocWord.SaveAs2(filename, Word.WdSaveFormat.wdFormatPDF);
                    filename = null;
                    break;
            }
        }

        private string filename;

        private async void Button11_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = Properties.Settings.Default.defaultPath;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                saveFileDialog1.ShowDialog();
                filename = saveFileDialog1.FileName;
                string cb1 = comboBox1.Text;
                string cb2 = comboBox2.Text;
                string cb3 = comboBox3.Text;
                string cb4 = comboBox4.Text;
                string lb1 = label1.Text;
                await Task.Run(() => wordandpdf(1, cb1, cb2, cb3, cb4, lb1));//ma code
                filename = null;
            }
        }

        private async void Button12_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = Properties.Settings.Default.defaultPath;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filename = saveFileDialog1.FileName;
                string cb1 = comboBox1.Text;
                string cb2 = comboBox2.Text;
                string cb3 = comboBox3.Text;
                string cb4 = comboBox4.Text;
                string lb1 = label1.Text;
                await Task.Run(() => wordandpdf(2, cb1, cb2, cb3, cb4, lb1));
                filename = null;
            }            
        }

        private async void Button13_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = Properties.Settings.Default.defaultPath;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                saveFileDialog1.ShowDialog();
                filename = saveFileDialog1.FileName;
                string cb1 = comboBox1.Text;
                string cb2 = comboBox2.Text;
                string cb3 = comboBox3.Text;
                string cb4 = comboBox4.Text;
                string lb1 = label1.Text;
                await Task.Run(() => excel(cb1, cb2, cb3, cb4, lb1));
            }
        }

        private void excel(string cb1text, string cb2text, string cb3text, string cb4text, string l1text)
        {
            //Объявляем приложение 
            Excel.Application ex = new Excel.Application();
            //Отобразить Excel 
            ex.Visible = true;
            //Количество листов в рабочей книге 
            ex.SheetsInNewWorkbook = 1;
            //Добавить рабочую книгу 
            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
            //Отключить отображение окон с сообщениями 
            ex.DisplayAlerts = false;
            //Получаем первый лист документа (счет начинается с 1) 
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
            //Название листа (вкладки снизу) 
            sheet.Name = "Чек";
            //Пример заполнения ячеек 
            sheet.Cells[1, 1] = "Клиент: ";
            sheet.Cells[1, 2] = cb1text;
            sheet.Cells[2, 1] = "Номер машины: ";
            sheet.Cells[2, 2] = cb2text;
            sheet.Cells[3, 1] = "Сотрудник:";
            sheet.Cells[3, 2] = cb3text;
            sheet.Cells[4, 1] = "Услуга: ";
            sheet.Cells[4, 2] = cb4text;
            sheet.Cells[5, 1] = "Время оформления заказа ";
            sheet.Cells[5, 2] = DateTime.Now;
            sheet.Cells[6, 1] = "Цена: ";
            sheet.Cells[6, 2] = l1text;

            ex.Application.ActiveWorkbook.SaveAs(filename, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            filename = null;
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            SettingsWindow f = new SettingsWindow();
            f.ShowDialog();
        }

        private void DataGridView3_SelectionChanged(object sender, EventArgs e)
        {            
            try
            {
                if (dataGridView3.SelectedRows[0].Cells["Статус заказа"].Value.ToString() == "False")
                {
                    checkBox1.Checked = false;
                }
                else
                {
                    checkBox1.Checked = true;
                }
            }
            catch
            {

            }
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                //MessageBox.Show(dataGridView3.SelectedRows[0].Cells["КлиентCODE"].Value.ToString());
                if (checkBox1.Checked == true)
                {
                    SqlCommand sqlCommand = new SqlCommand("update zak set stat = 1 where id_zak ="
                        + dataGridView3.SelectedRows[0].Cells["id_zak"].Value.ToString(), db.connection);
                    sqlCommand.ExecuteNonQuery();
                    update();
                }
                else
                {
                    SqlCommand sqlCommand = new SqlCommand("update zak set stat = 0 where id_zak ="
                        + dataGridView3.SelectedRows[0].Cells["id_zak"].Value.ToString(), db.connection);
                    sqlCommand.ExecuteNonQuery();
                    update();
                }
            }
            catch
            {

            }
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void DataGridView1_CellContentClick(object sender, EventArgs e)
        {
            try
            {
                cueTextBox2.Text = dataGridView1.SelectedRows[0].Cells["nom_machine"].Value.ToString();
                comboBox7.SelectedValue = dataGridView1.SelectedRows[0].Cells["MODEL_ID"].Value.ToString();
                comboBox6.SelectedValue = dataGridView1.SelectedRows[0].Cells["MARK_ID"].Value.ToString();
            }
            catch
            {

            }
        }

        private void DataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                cueTextBox3.Text = dataGridView4.SelectedRows[0].Cells["KLIENT_FAM"].Value.ToString();
                cueTextBox4.Text = dataGridView4.SelectedRows[0].Cells["KLIENT_IM"].Value.ToString();
                cueTextBox5.Text = dataGridView4.SelectedRows[0].Cells["KLIENT_OTCH"].Value.ToString();
                cueTextBox6.Text = dataGridView4.SelectedRows[0].Cells["KLIENT_NUM"].Value.ToString();
            }
            catch
            {

            }            
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            SqlCommand sqlCommand = new SqlCommand("EXEC INSERT_KLIENT '"+cueTextBox3.Text+"', '"+cueTextBox4.Text+"','"+cueTextBox5.Text+"','"+cueTextBox6.Text+"'", db.connection);
            sqlCommand.ExecuteNonQuery();
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            SqlCommand sqlCommand = new SqlCommand("EXEC INSERT_Machine '"+cueTextBox2.Text+"', "+ 0 +", "+comboBox7.SelectedValue+"", db.connection);
            sqlCommand.ExecuteNonQuery();
        }

        private void Button14_Click(object sender, EventArgs e)
        {
            SqlCommand sqlCommand = new SqlCommand("EXEC INSERT_Machine '" + cueTextBox2.Text + "', " + 0 + ", " + comboBox7.Text + "," + dataGridView1.SelectedRows[0].Cells["id_machine"].Value.ToString());
            sqlCommand.ExecuteNonQuery();
        }

        private void Button15_Click(object sender, EventArgs e)
        {
            SqlCommand sqlCommand = new SqlCommand("exec delete_machine " + dataGridView1.SelectedRows[0].Cells["id_machine"].Value.ToString(), db.connection);
            sqlCommand.ExecuteNonQuery();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            SqlCommand sqlCommand = new SqlCommand("exec insert_dolj ' " + cueTextBox1.Text +"'", db.connection);
            sqlCommand.ExecuteNonQuery();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            SqlCommand sqlCommand = new SqlCommand("exec update_dolj " + dataGridView1.SelectedRows[0].Cells["id_dolj"].Value.ToString() +"," + cueTextBox1.Text + "'", db.connection);
            sqlCommand.ExecuteNonQuery();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            SqlCommand sqlCommand = new SqlCommand("exec delete_dolj " + dataGridView1.SelectedRows[0].Cells["id_dolj"].Value.ToString(), db.connection);
            sqlCommand.ExecuteNonQuery();
        }
    }
}


//------------------------------------------------------


