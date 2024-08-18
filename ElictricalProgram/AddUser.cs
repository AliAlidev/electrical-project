using System;
using System.Windows.Forms;

namespace Elictrical_Program
{
    public partial class AddUser : Form
    {
        public AddUser()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || comboBox1.Text == "")
            {
                MessageBox.Show("يجب ادخال كافة البيانات");
            }
            else
            {
                string user_name = textBox1.Text;
                string user_password = DBFunctions.Base64Encode(textBox2.Text);
                bool is_active = checkBox1.Checked;
                string last_login = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                int user_role = comboBox1.SelectedIndex;
                if (DBFunctions.isUserExist(user_name))
                {
                    MessageBox.Show("اسم المستخدم موجود مسبقاً");
                }
                else
                {
                    string qrt = "insert into Users(user_name, user_password, is_active, last_login, user_role) values('" + user_name + "','" + user_password
                        + "','" + is_active + "','" + last_login + "'," + user_role + ")";
                    DBFunctions.executeCommand(qrt);
                    MessageBox.Show("تم ادخال المستخدم بنجاح");
                    textBox1.Clear();
                    textBox2.Clear();
                    comboBox1.ResetText();
                    checkBox1.Checked = false;
                }
            }


        }
    }
}


