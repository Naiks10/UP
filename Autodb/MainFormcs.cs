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
            //ZakTable.Rows[0].Selected = true;
        }

        public MainFormcs()
        {
            InitializeComponent();
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

        private void wordandpdf(int a)
        {
            Word.Application winword = new Word.Application(); //создаем COM-объект Word 
            object missing = System.Reflection.Missing.Value;
            winword.Visible = false;
            Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            document.Content.Text = "Чек" + Environment.NewLine + "Клиент: " + comboBox1.Text + Environment.NewLine + "Номер машины: " + comboBox2.Text + Environment.NewLine + "Сотрудник: " + comboBox3.Text + Environment.NewLine + "Услуга: " + comboBox4.Text +
            Environment.NewLine + "Время оформления заказа " + DateTime.Now + Environment.NewLine + "Цена: " + label1.Text;
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

        private void Button11_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
            filename = saveFileDialog1.FileName;
            wordandpdf(1);
            filename = null;
        }

        private void Button12_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
            filename = saveFileDialog1.FileName;
            wordandpdf(2);
            filename = null;
        }

        private void Button13_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
            filename = saveFileDialog1.FileName;
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
            sheet.Cells[1, 2] = comboBox1.Text;
            sheet.Cells[2, 1] = "Номер машины: ";
            sheet.Cells[2, 2] = comboBox2.Text;
            sheet.Cells[3, 1] = "Сотрудник:";
            sheet.Cells[3, 2] = comboBox3.Text;
            sheet.Cells[4, 1] = "Услуга: ";
            sheet.Cells[4, 2] = comboBox4.Text;
            sheet.Cells[5, 1] = "Время оформления заказа ";
            sheet.Cells[5, 2] = DateTime.Now;
            sheet.Cells[6, 1] = "Цена: ";
            sheet.Cells[6, 2] = label1.Text;

            ex.Application.ActiveWorkbook.SaveAs(filename, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            filename = null;
        }
    }
}


//------------------------------------------------------


