using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace _471_Stepanenko_Laba_14
{
    public partial class Form1 : Form
    {
        public static string connectString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", @"C:\Users\USER\source\repos\471 Stepanenko Laba 14\471 Stepanenko Laba 14\stepanenko_database.mdb");
        public Form1()
        {
            InitializeComponent();
        }

        private void getData_Click(object sender, EventArgs e)
        {
            GetData();
        }
        private void deleteOne_Click(object sender, EventArgs e)
        {
            DeleteOne();
            GetData();
        }
        private void GetData()
        {
            using (OleDbConnection myConnection = new OleDbConnection(connectString))
            {
                try
                {
                    myConnection.Open();
                    string query = "select * from Contacts";
                    OleDbCommand command = new OleDbCommand(query, myConnection);
                    OleDbDataReader reader = command.ExecuteReader();
                    listBox1.Items.Clear();
                    listBox2.Items.Clear();
                    listBox3.Items.Clear();
                    listBox4.Items.Clear();
                    listBox5.Items.Clear();
                    listBox6.Items.Clear();
                    while (reader.Read())
                    {
                        listBox1.Items.Add(String.Format("{0} {1}.{2}.", reader[1].ToString(), reader[2].ToString()[0], reader[3].ToString()[0]));
                        listBox2.Items.Add(reader[4].ToString());
                        listBox3.Items.Add(reader[5].ToString());
                        listBox4.Items.Add(reader[6].ToString());
                        listBox5.Items.Add(reader[7].ToString());
                        listBox6.Items.Add(reader[0].ToString());
                    }
                    reader.Close();
                    myConnection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void DeleteOne()
        {
            using (OleDbConnection myConnection = new OleDbConnection(connectString))
            {
                try
                {
                    myConnection.Open();
                    string index = listBox6.SelectedItem.ToString();
                    string query = String.Format("DELETE FROM Contacts WHERE Код = {0}", index);

                    OleDbCommand command = new OleDbCommand(query, myConnection);
                    command.ExecuteNonQuery();
                    myConnection.Close();
                }
                catch
                {
                    MessageBox.Show("Щоб видалити поле, потрібно виділити поле id", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void newContact_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }

        private void deleteAll_Click(object sender, EventArgs e)
        {
            DeleteAll();
            GetData();
        }
        private void DeleteAll()
        {
            using (OleDbConnection myConnection = new OleDbConnection(connectString))
            {
                try
                {
                    myConnection.Open();
                    string query = "DELETE FROM Contacts";

                    // создаем объект OleDbCommand для выполнения запроса к БД MS Access
                    OleDbCommand command = new OleDbCommand(query, myConnection);

                    // выполняем запрос к MS Access
                    command.ExecuteNonQuery();
                    myConnection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void searchButton_Click(object sender, EventArgs e)
        {
            using (OleDbConnection myConnection = new OleDbConnection(connectString))
            {
                try
                {
                    if(searchBox.Text == "")
                    {
                        MessageBox.Show("Ви не ввели значення пошуку", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        myConnection.Open();
                        string query = String.Format("select * from Contacts where city = \"{0}\"", searchBox.Text);
                        OleDbCommand command = new OleDbCommand(query, myConnection);
                        OleDbDataReader reader = command.ExecuteReader();
                        listBox1.Items.Clear();
                        listBox2.Items.Clear();
                        listBox3.Items.Clear();
                        listBox4.Items.Clear();
                        listBox5.Items.Clear();
                        listBox6.Items.Clear();
                        while (reader.Read())
                        {
                            listBox1.Items.Add(String.Format("{0} {1}.{2}.", reader[1].ToString(), reader[2].ToString()[0], reader[3].ToString()[0]));
                            listBox2.Items.Add(reader[4].ToString());
                            listBox3.Items.Add(reader[5].ToString());
                            listBox4.Items.Add(reader[6].ToString());
                            listBox5.Items.Add(reader[7].ToString());
                            listBox6.Items.Add(reader[0].ToString());
                        }
                        reader.Close();
                        myConnection.Close();
                        if(listBox1.Items.Count == 0)
                        {
                            MessageBox.Show("Такого міста немає");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
