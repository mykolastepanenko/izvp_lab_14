using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace _471_Stepanenko_Laba_14
{
    public partial class Form2 : Form
    {
        public static string connectString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", @"C:\Users\USER\source\repos\471 Stepanenko Laba 14\471 Stepanenko Laba 14\stepanenko_database.mdb");
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OleDbConnection myConnection = new OleDbConnection(connectString))
            {
                try
                {
                    myConnection.Open();
                    if (
                        textBox1.Text == "" ||
                        textBox2.Text == "" ||
                        textBox3.Text == "" ||
                        textBox4.Text == "" ||
                        textBox5.Text == "" ||
                        textBox6.Text == "" ||
                        textBox7.Text == ""
                        )
                    {
                        MessageBox.Show("Ви не ввели всі значення нового контакту", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        string query = String.Format("insert into Contacts (name, lastname, secondname, city, phone, address, email) values (\"{0}\", \"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\",\"{6}\")", textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text, textBox7.Text);
                        OleDbCommand command = new OleDbCommand(query, myConnection);
                        command.ExecuteNonQuery();
                    }

                    myConnection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
