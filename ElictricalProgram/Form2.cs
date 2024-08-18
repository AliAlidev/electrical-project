using System;
using System.Data;
using System.Windows.Forms;

namespace Elictrical_Program
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string qrt = "";
            if (!DBFunctions.isSectionExist(date))
            {
                qrt = "insert into Sections(working_date, station, section_name, morning_notes_1, evening_notes_1, morning_notes_2, evening_notes_2, total, supervisor_engineer, user_id) " + " VALUES('" + date + "','" + textBox2.Text + "','" + textBox3.Text + "','" + richTextBox2.Text + "','" + richTextBox1.Text + "','" + richTextBox4.Text + "','" + richTextBox3.Text + "','0','" + textBox1.Text + "'," + DBFunctions.user_id + ")";
            }
            else
            {
                qrt = "update Sections set station='" + textBox2.Text + "', section_name='" + textBox3.Text + "', morning_notes_1='" + richTextBox2.Text + "', evening_notes_1='" + richTextBox1.Text + "', morning_notes_2='" + richTextBox4.Text + "', evening_notes_2='" + richTextBox3.Text + "', supervisor_engineer='" + textBox1.Text + "' where working_date = '" + date + "', user_id=" + DBFunctions.user_id;
            }
            DBFunctions.executeCommand(qrt);
            MessageBox.Show("تم ادخال البيانات بنجاح");
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            string date = DBFunctions.now;
            string qrt = "select * from Sections where working_date='" + date + "'";
            DataSet ds = DBFunctions.fillDataSet(qrt);
            if (ds.Tables[0].Rows.Count > 0)
            {
                textBox2.Text = ds.Tables[0].Rows[0].ItemArray[2].ToString();
                textBox3.Text = ds.Tables[0].Rows[0].ItemArray[3].ToString();
                richTextBox2.Text = ds.Tables[0].Rows[0].ItemArray[4].ToString();
                richTextBox1.Text = ds.Tables[0].Rows[0].ItemArray[5].ToString();
                richTextBox4.Text = ds.Tables[0].Rows[0].ItemArray[6].ToString();
                richTextBox3.Text = ds.Tables[0].Rows[0].ItemArray[7].ToString();
                textBox1.Text = ds.Tables[0].Rows[0].ItemArray[9].ToString();
            }
        }
    }
}
