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
using System.Data.OleDb;

namespace Journal
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void загрузкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Создаём соединение
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0; Data Source=JournalBD.mdb"; //строка соединения
            OleDbConnection dbConnection = new OleDbConnection(connectionString); //создаём соединение

            //Выполняем запрос к БД
            dbConnection.Open(); //открываем соединение
            string query = "SELECT * FROM Журнал"; //строка запроса
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection); //команда
            OleDbDataReader dbReader = dbCommand.ExecuteReader(); //считываем данные

            //Проверяем данные
            if (dbReader.HasRows == false)
            {
                MessageBox.Show("Данные не найдены!", "Ошибка");
            }
            else
            {
                //Запишем данные в таблицу формы
                while (dbReader.Read())
                {
                    //Выводим данные
                    dataGridView1.Rows.Add(dbReader["Номер_в_журнале"], dbReader["ФИО_учащегося"], dbReader["Предмет"], dbReader["Оценка"], dbReader["Дата"]);
                }
            }

            //Закрываем соединение
            dbReader.Close();
            dbConnection.Close();

        }

        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Проверим количество выбранных строк
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!", "Внимание");
                return;

            }

            //Заполним выбранную строку
            int index = dataGridView1.SelectedRows[0].Index;

            //Проверим данные таблицы
            if (dataGridView1.Rows[index].Cells[0].Value == null ||
               dataGridView1.Rows[index].Cells[1].Value == null ||
               dataGridView1.Rows[index].Cells[2].Value == null ||
               dataGridView1.Rows[index].Cells[3].Value == null ||
               dataGridView1.Rows[index].Cells[4].Value == null)
            {
                MessageBox.Show("Не все данные введены", "Внимание!");
                return;
            }

            //Считаем данные
            string Nomer = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string FIO = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string Predmet = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string Ocenka = dataGridView1.Rows[index].Cells[3].Value.ToString();
            string Data = dataGridView1.Rows[index].Cells[4].Value.ToString();

            //Создаём соединение
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0; Data Source=JournalBD.mdb"; //строка соединения
            OleDbConnection dbConnection = new OleDbConnection(connectionString); //создаём соединение

            //Выполняем запрос к БД
            dbConnection.Open(); //открываем соединение
            string query = "INSERT INTO Журнал VALUES(" + Nomer + ", '" + FIO + "', '" + Predmet + "', '" + Ocenka + "', '" + Data + "')"; //строка запроса
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection); //команда

            //Выполняем запрос
            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса", "Ошибка!");
            else
                MessageBox.Show("Данные добавлены!", "Внимание!");

            //Закрываем соединение с БД
            dbConnection.Close();


        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Проверим количество выбранных строк
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!", "Внимание");
                return;
            }

            //Заполним выбранную строку
            int index = dataGridView1.SelectedRows[0].Index;

            //Проверим данные таблицы
            if (dataGridView1.Rows[index].Cells[0].Value == null)
            {
                MessageBox.Show("Не все данные введены", "Внимание!");
                return;
            }

            //Считаем данные
            string Nomer = dataGridView1.Rows[index].Cells[0].Value.ToString();

            //Создаём соединение
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0; Data Source=JournalBD.mdb"; //строка соединения
            OleDbConnection dbConnection = new OleDbConnection(connectionString); //создаём соединение

            //Выполняем запрос к БД
            dbConnection.Open(); //открываем соединение
            string query = "DELETE FROM Журнал WHERE Номер_в_журнале = " + Nomer; //строка запроса
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection); //команда

            //Выполняем запрос
            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса", "Ошибка!");
            else
            {
                MessageBox.Show("Данные удалены!", "Внимание!");
                //Удаляем данные из таблицы в форме
                dataGridView1.Rows.RemoveAt(index);
            }

            //Закрываем соединение с БД
            dbConnection.Close();
        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Проверим количество выбранных строк
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Выберите одну строку!", "Внимание");
                return;
            }

            //Заполним выбранную строку
            int index = dataGridView1.SelectedRows[0].Index;

            //Проверим данные таблицы
            if (dataGridView1.Rows[index].Cells[0].Value == null ||
               dataGridView1.Rows[index].Cells[1].Value == null ||
               dataGridView1.Rows[index].Cells[2].Value == null ||
               dataGridView1.Rows[index].Cells[3].Value == null ||
               dataGridView1.Rows[index].Cells[4].Value == null)
            {
                MessageBox.Show("Не все данные введены", "Внимание!");
                return;
            }

            //Считаем данные
            string Nomer = dataGridView1.Rows[index].Cells[0].Value.ToString();
            string FIO = dataGridView1.Rows[index].Cells[1].Value.ToString();
            string Predmet = dataGridView1.Rows[index].Cells[2].Value.ToString();
            string Ocenka = dataGridView1.Rows[index].Cells[3].Value.ToString();
            string Data = dataGridView1.Rows[index].Cells[4].Value.ToString();

            //Создаём соединение
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0; Data Source=JournalBD.mdb"; //строка соединения
            OleDbConnection dbConnection = new OleDbConnection(connectionString); //создаём соединение

            //Выполняем запрос к БД
            dbConnection.Open(); //открываем соединение
            string query = "UPDATE Журнал SET ФИО_учащегося = '" + FIO + "', Предмет = '" + Predmet + "', Оценка = '" + Ocenka + "', " +
                "Дата = '" + Data + "' WHERE Номер_в_журнале = " + Nomer; //строка запроса
            OleDbCommand dbCommand = new OleDbCommand(query, dbConnection); //команда

            //Выполняем запрос
            if (dbCommand.ExecuteNonQuery() != 1)
                MessageBox.Show("Ошибка выполнения запроса", "Ошибка!");
            else
            {
                MessageBox.Show("Данные изменены!", "Внимание!");
            }

            //Закрываем соединение с БД
            dbConnection.Close();

        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
