using System.Windows.Forms;

namespace Elictrical_Program
{
    public partial class UsersList : Form
    {
        public UsersList()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0)
            {
                string user_name = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                string status = DBFunctions.activateUser(user_name);
                dataGridView1.Rows[e.RowIndex].Cells[3].Value = status == "True"? "نعم" : "لا";
                dataGridView1.Rows[e.RowIndex].Cells[4].Value = status == "True" ? "الغاء تفعيل" : "تفعيل";
            }
        }
    }
}
