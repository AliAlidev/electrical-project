using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Elictrical_Program
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string user_name = textBox1.Text;
            string user_password = textBox2.Text;
            if (user_name == "" || user_password == "")
            {
                MessageBox.Show("يجب ادخال بيانات المستخدم");
            }
            else
            {
                int is_valid_user = DBFunctions.isValidLogin(user_name, user_password);
                if(is_valid_user == 2)
                {
                    DBFunctions.is_valid_user = false;
                    MessageBox.Show("المستخدم غير مفعل يرجى التواصل مع مدير التطبيق لتفعيله");
                }
                else if (is_valid_user == 1)
                {
                    DataSet ds = DBFunctions.getUserData(user_name);
                    DBFunctions.is_valid_user = true;
                    DBFunctions.user_name = ds.Tables[0].Rows[0].ItemArray[1].ToString();
                    DBFunctions.user_role = (int)ds.Tables[0].Rows[0].ItemArray[5];
                    DBFunctions.user_id = (int)ds.Tables[0].Rows[0].ItemArray[0];
                    Close();
                }
                else
                {
                    DBFunctions.is_valid_user = false;
                    MessageBox.Show("بيانات المستخدم غير صحيحة");
                }
            }
        }

        private void Login_Load(object sender, EventArgs e)
        {
            DBFunctions.conn = new OleDbConnection();
            string database_path = Application.StartupPath + "\\db";
            DBFunctions.conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + database_path + ";Jet OLEDB:Database Password=" + DBFunctions.Base64Decode("RDA1XjczMiMzMEBCQzgqOTA1Q3g=");

            DBFunctions.createNewAdmin();
        }
    }
}
