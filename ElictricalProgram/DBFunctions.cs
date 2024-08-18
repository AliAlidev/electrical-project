using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Elictrical_Program
{
    static class DBFunctions
    {
        public static OleDbConnection conn;
        public static OleDbDataAdapter adap;
        public static OleDbCommand comm;
        public static string now;
        public static string user_name;
        public static int user_role, user_id;
        public static bool is_valid_user;

        public static void openConnection()
        {
            if (conn.State != ConnectionState.Open)
                conn.Open();
        }
        public static void closeConnection()
        {
            if (conn.State == ConnectionState.Open)
                conn.Close();
        }

        public static DataSet fillDataSet(string qrt)
        {
            adap = new OleDbDataAdapter(qrt, conn);
            DataSet ds = new DataSet();
            adap.Fill(ds);
            return ds;
        }

        public static void executeCommand(string qrt)
        {
            comm = new OleDbCommand(qrt, conn);
            openConnection();
            comm.ExecuteNonQuery();
            closeConnection();
        }

        public static bool isHourExist(string hour, string date)
        {
            string qrt = "select count(*) from Readings where working_hour='" + hour + "' and working_date='" + date + "'";
            comm = new OleDbCommand(qrt, conn);
            openConnection();
            int result = (int)comm.ExecuteScalar();
            return result > 0 ? true : false;
        }

        public static bool isUserExist(string name)
        {
            string qrt = "select count(*) from Users where user_name='" + name + "'";
            comm = new OleDbCommand(qrt, conn);
            openConnection();
            int result = (int)comm.ExecuteScalar();
            return result > 0 ? true : false;
        }

        public static bool isSectionExist(string date)
        {
            string qrt = "select count(*) from Sections where working_date='" + date + "'";
            comm = new OleDbCommand(qrt, conn);
            openConnection();
            int result = (int)comm.ExecuteScalar();
            return result > 0 ? true : false;
        }

        public static DataSet convertDataridViewToDataSet(DataGridView dgv)
        {
            DataTable dt = new DataTable();
            dt = (DataTable)dgv.DataSource;

            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            return ds;
        }

        public static bool isEmptyDataridView(DataGridView dgv)
        {
            bool isEmpty = true;
            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ColumnIndex > 0 && cell.Value != null && cell.Value.ToString() != "")
                    {
                        isEmpty = false; break;
                    }
                }
            }
            return isEmpty;
        }

        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        public static int isValidLogin(string name, string password)
        {
            string qrt = "select count(*) from Users where user_name='" + name + "' and user_password='" + Base64Encode(password) + "'";
            comm = new OleDbCommand(qrt, conn);
            openConnection();
            int result = (int)comm.ExecuteScalar();
            if (result > 0)
            {
                if (getUserStatus(name) == "True")
                {
                    qrt = "update Users set last_login='" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "' where user_name='" + name + "'";
                    executeCommand(qrt);
                }
                else
                {
                    return 2;
                }
            }
            return result > 0 ? 1 : 0;
        }

        public static void createNewAdmin()
        {
            if (!isUserExist("aliali"))
            {
                string qrt = "insert into Users(user_name, user_password, is_active, last_login, user_role) values('aliali','" + Base64Encode("kimim519899")
                + "','True','" + DateTime.Now.ToString("yyyy-mm-dd hh:mm:ss") + "',0)";
                comm = new OleDbCommand(qrt, conn);
                executeCommand(qrt);
            }
        }

        public static DataSet getUserData(string name)
        {
            string qrt = "select * from Users where user_name='" + name + "'";
            DataSet userData = fillDataSet(qrt);
            return userData;
        }

        public static string getUserStatus(string name)
        {
            string qrt = "select is_active from Users where user_name='" + name + "'";
            DataSet ds = fillDataSet(qrt);
            return ds.Tables[0].Rows[0].ItemArray[0].ToString();
        }

        public static string activateUser(string name)
        {
            string qrt = "", status = "";
            if (getUserStatus(name) == "False")
            {
                qrt = "update Users set is_active = 'True' where user_name='" + name + "'";
                status = "True";
            }
            else
            {
                qrt = "update Users set is_active = 'False' where user_name='" + name + "'";
                status = "False";
            }
            executeCommand(qrt);
            return status;
        }

        public static bool isAppValid()
        {
            string qrt = "select count(*) from Activation";
            comm = new OleDbCommand(qrt, conn);
            openConnection();
            int result = (int)comm.ExecuteScalar();
            if (result > 0)
            {
                qrt = "SELECT count(*) FROM Activation WHERE DateDiff('d', start_date, Date()) > period;";
                DataSet ds = fillDataSet(qrt);
                int result1 = 0;
                int.TryParse(ds.Tables[0].Rows[0].ItemArray[0].ToString(), out result1);
                if (result1 > 0)
                    return false;

                qrt = "update Activation set last_login='" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "'";
                executeCommand(qrt);
            }
            else
            {
                qrt = "insert into Activation(period, start_date, last_login) values(360, '" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "','" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "')";
                executeCommand(qrt);
            }
            return true;
        }

    }
}
