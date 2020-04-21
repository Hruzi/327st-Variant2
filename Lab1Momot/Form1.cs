using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lab1Momot
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Update();


        }
        private void Update()
        {

            FileDBF dbf = new FileDBF();
            var dt1 = dbf.Execute(@"SELECT ACCT_NBR AS Номер_Аккаунта, SYMBOL AS Символ, SHARES AS Акции, PUR_PRICE AS Стоимость_Покупки, PUR_DATE AS Дата_покупки FROM holdings.dbf ");
            var dt2 = dbf.Execute(@"SELECT INDUSTRY AS Индустрия, SYMBOL AS Символ,CO_NAME AS Название_Компании, EXCHANGE AS Биржа, CUR_PRICE AS Цена, YRL_HIGH AS Максимальная_цена, YRL_LOW AS Минимальная_цена, P_E_RATIO AS Соотношение,
            BETA AS Бета, PROJ_GRTH AS Рост,  PRICE_CHG AS Изменение_цены, SAFETY AS Безопасность, RATING AS Рейтинг, RANK AS Ранк, OUTLOOK AS Прогноз, RCMNDATION AS Рекомендации, RISK AS Риски FROM master.dbf ");
            var dt3 = dbf.Execute(@"SELECT IND_CODE AS Код, IND_NAME AS Сокращенное_Название, LONG_NAME AS Полное_название  FROM industry.dbf ");
           
            dataGridView1.DataSource = dt1;
            dataGridView2.DataSource = dt2;
            dataGridView3.DataSource = dt3;

           

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        public class FileDBF
        {
            private OleDbConnection _connection = null;

            private const string putFileName = @"C:\Users\Hruzi\Desktop\3\3_2Momot3kurs2semestr\Programming Tecnology\lab1\DBDEMOS"; // сюда пишите ПОЛНЫЙ ПУТЬ к ПАПКЕ.

            public DataTable Execute(string command)
            {
                DataTable dt = null;
                if (_connection != null)
                {
                    try
                    {
                        _connection.Open();
                        dt = new DataTable();
                        OleDbCommand oCmd = _connection.CreateCommand();
                        oCmd.CommandText = command;
                        dt.Load(oCmd.ExecuteReader());
                        _connection.Close();
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }
                }
                return dt;
            }

            public DataTable GetAll(string dbpath)
            {

                return Execute("SELECT * FROM " + dbpath); ;
            }

            public FileDBF()
            {
                _connection = new OleDbConnection();
                _connection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + putFileName + "; Extended Properties=dBASE IV;";
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedCells[0].ColumnIndex == 1)
            {
                FileDBF fld = new FileDBF();
                string num = dataGridView2.SelectedCells[0].Value.ToString();

                var dt3 = fld.Execute(@"SELECT ACCT_NBR AS Номер_Аккаунта, SYMBOL AS Символ, SHARES AS Акции, PUR_PRICE AS Стоимость_Покупки, PUR_DATE AS Дата_покупки FROM holdings.dbf WHERE SYMBOL = '" + num + "';");
                dataGridView4.DataSource = dt3;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView3.SelectedCells[0].ColumnIndex == 0)
            {
                FileDBF fld = new FileDBF();
                int num = Convert.ToInt32(dataGridView3.SelectedCells[0].Value.ToString());
                var dt5 = fld.Execute(@"SELECT INDUSTRY AS Индустрия, SYMBOL AS Символ,CO_NAME AS Название_Компании, EXCHANGE AS Биржа, CUR_PRICE AS Цена, YRL_HIGH AS Максимальная_цена, YRL_LOW AS Минимальная_цена, P_E_RATIO AS Соотношение,
                BETA AS Бета, PROJ_GRTH AS Рост,  PRICE_CHG AS Изменение_цены, SAFETY AS Безопасность, RATING AS Рейтинг, RANK AS Ранк, OUTLOOK AS Прогноз, RCMNDATION AS Рекомендации, RISK AS Риски FROM master.dbf  WHERE INDUSTRY = " + num + ";");
                dataGridView5.DataSource = dt5;
            }
            
        }
    }
}
