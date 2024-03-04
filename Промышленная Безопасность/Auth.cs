using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Промышленная_Безопасность
{
    public partial class Auth : Form
    {
        public Auth()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text;
            string pass = textBox2.Text;

            try
            {
                var db = new DBContext();
                var table = db.Execute(
                    $"Select * from [Аккаунты] where [Имя]=N'{name}' and [Пароль]=N'{pass}'");
                if (table.Count != 0 && string.Equals(table[0][2], name) && string.Equals(table[0][3], pass))
                {
                    this.Hide();
                    bool isAdmin = Convert.ToBoolean(table[0][1]);
                    Form1 form = new Form1(this.Close, name, isAdmin);
                    try
                    {
                        form.Show();
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("Неверное имя пользователя и/или пароль");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
