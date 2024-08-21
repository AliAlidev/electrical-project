using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Elictrical_Program
{
    public partial class Form1 : Form
    {
        private TextBox[] textBoxs = new TextBox[100];
        private Label[] labels = new Label[100];
        private DataTable readingTable;
        private Rectangle originalFormSize;
        Control[] anotherComponents;
        Rectangle[] anotherComponentsRectangle;
        public Form1()
        {
            InitializeComponent();
            anotherComponents = new Control[100];
            anotherComponentsRectangle = new Rectangle[100];
        }

        private void take_database_backup()
        {
            string file_name = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
            string directory_path = Application.StartupPath + "\\backups\\" + file_name;
            if (!Directory.Exists(directory_path))
            {
                Directory.CreateDirectory(directory_path);
            }

            string fileToCopy = Application.StartupPath + "\\db";
            string destinationFile = directory_path + "\\" + file_name + ".back";
            string random_string = DateTime.Now.Ticks.ToString();
            destinationFile = destinationFile + "." + random_string;
            File.Copy(fileToCopy, destinationFile);
        }

        private void check_required_directories()
        {
            string destinationDirectory = Application.StartupPath + "\\backups";
            if (!Directory.Exists(destinationDirectory))
                Directory.CreateDirectory(destinationDirectory);

            destinationDirectory = Application.StartupPath + "\\xlsxs";
            if (!Directory.Exists(destinationDirectory))
                Directory.CreateDirectory(destinationDirectory);
        }

        private void Form1_Load(object sender, System.EventArgs e)
        {
            check_required_directories();
            take_database_backup();
            Login lg = new Login();
            lg.ShowDialog();
            if (!DBFunctions.is_valid_user)
            {
                Application.Exit();
                goto program_end;
            }

            DBFunctions.now = DateTime.Now.ToString("yyyy-MM-dd");
            linkLabel1.Visible = false;
            copyGroupBox(0, "بانياس1", flowLayoutPanel1, "old");
            copyGroupBox(4, "بانياس2", flowLayoutPanel1, "old");
            copyGroupBox(8, "سمريان1", flowLayoutPanel1, "old");
            copyGroupBox(12, "سمريان2", flowLayoutPanel1, "old");
            copyGroupBox(16, "وصول1", flowLayoutPanel1, "old");
            copyGroupBox(20, "وصول2", flowLayoutPanel1, "old");
            copyGroupBox(24, "وصول3", flowLayoutPanel1, "old");
            copyGroupBox(28, "عمريت", flowLayoutPanel1, "old");
            copyGroupBox(32, "اسمنت", flowLayoutPanel1, "old");
            copyGroupBox(36, "الشركة", flowLayoutPanel1, "old");
            copyGroupBox(40, "محولة1", flowLayoutPanel1, "old");
            copyGroupBox(44, "محولة2", flowLayoutPanel1, "old");
            copyGroupBox(48, "محولة3", flowLayoutPanel1, "old");
            copyGroupBox(52, "الشمال", flowLayoutPanel1, "old");

            copyGroupBox(0, "بانياس1", flowLayoutPanel2, "current");
            copyGroupBox(4, "بانياس2", flowLayoutPanel2, "current");
            copyGroupBox(8, "سمريان1", flowLayoutPanel2, "current");
            copyGroupBox(12, "سمريان2", flowLayoutPanel2, "current");
            copyGroupBox(16, "وصول1", flowLayoutPanel2, "current");
            copyGroupBox(20, "وصول2", flowLayoutPanel2, "current");
            copyGroupBox(24, "وصول3", flowLayoutPanel2, "current");
            copyGroupBox(28, "عمريت", flowLayoutPanel2, "current");
            copyGroupBox(32, "اسمنت", flowLayoutPanel2, "current");
            copyGroupBox(36, "الشركة", flowLayoutPanel2, "current");
            copyGroupBox(40, "محولة1", flowLayoutPanel2, "current");
            copyGroupBox(44, "محولة2", flowLayoutPanel2, "current");
            copyGroupBox(48, "محولة3", flowLayoutPanel2, "current");
            copyGroupBox(52, "الشمال", flowLayoutPanel2, "current");

            // create reading table 
            createReadingTable();

            // fill table with empty rows
            fillTableEmptyRows(readingTable, 25);

            // get last reading data
            string[] lastReadingData = getLastReadingData();

            // fill table with previous records
            fillTableWithCurrentDateData(lastReadingData);

            // fill from inputs with previous data
            setInitialValuesForStartInputs(lastReadingData);

            //////////////////
            originalFormSize = new Rectangle(this.Location.X, this.Location.Y, this.Size.Width, this.Size.Height);
            for (int i = 0; i < this.Controls.Count; i++)
            {
                anotherComponentsRectangle[i] = new Rectangle(this.Controls[i].Location.X, this.Controls[i].Location.Y, this.Controls[i].Size.Width, this.Controls[i].Size.Height);
                anotherComponents[i] = this.Controls[i];
            }

            if (DBFunctions.user_role == 0)
            {
                button6.Visible = true;
                button7.Visible = true;
                button8.Visible = true;
                dateTimePicker1.Enabled = true;
            }

            label4.Text = DBFunctions.user_name;

        program_end:;
        }

        string[] getLastReadingData()
        {
            string current_date = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
            string qrt = "select top 1 * from Readings where working_date='" + current_date + "' AND banias1_send <> '0' " +
                " AND banias1_send <> '0' AND banias1_receive <> '0'" +
                " AND banias2_send <> '0' AND banias2_receive <> '0'" +
                " AND semerian1_send <> '0' AND semerian1_receive <> '0'" +
                " AND semerian2_send <> '0' AND semerian2_receive <> '0'" +
                " AND arrival1_send <> '0'" +
                " AND arrival2_send <> '0'" +
                " AND arrival3_send <> '0'" +
                " AND amreet_send <> '0'" +
                " AND esmant_send <> '0'" +
                " AND company_send <> '0'" +
                " AND transformer1_send <> '0'" +
                " AND transformer2_send <> '0'" +
                " AND transformer3_send <> '0'" +
                " AND north_send <> '0'" +
                " order by id desc";
            DataSet ds = DBFunctions.fillDataSet(qrt);
            string[] readingData = new string[0];
            if (ds.Tables[0].Rows.Count > 0)
            {
                int items_count = ds.Tables[0].Rows[0].ItemArray.Length;
                if (items_count > 0)
                {
                    readingData = new string[items_count];
                    int counter = 0;
                    foreach (var item in ds.Tables[0].Rows[0].ItemArray)
                    {
                        readingData[counter] = item.ToString();
                        counter++;
                    }
                }
            }
            return readingData;
        }

        string[] getPreviousDayreadingValues()
        {
            string current_date = dateTimePicker1.Value.Date.AddDays(-1).ToString("yyyy-MM-dd");
            string qrt = "select top 1 * from Readings where working_date='" + current_date + "' AND banias1_send <> '0' " +
                " AND banias1_send <> '0' AND banias1_receive <> '0'" +
                " AND banias2_send <> '0' AND banias2_receive <> '0'" +
                " AND semerian1_send <> '0' AND semerian1_receive <> '0'" +
                " AND semerian2_send <> '0' AND semerian2_receive <> '0'" +
                " AND arrival1_send <> '0'" +
                " AND arrival2_send <> '0'" +
                " AND arrival3_send <> '0'" +
                " AND amreet_send <> '0'" +
                " AND esmant_send <> '0'" +
                " AND company_send <> '0'" +
                " AND transformer1_send <> '0'" +
                " AND transformer2_send <> '0'" +
                " AND transformer3_send <> '0'" +
                " AND north_send <> '0'" +
                " order by id desc";
            DataSet ds = DBFunctions.fillDataSet(qrt);
            string[] readingData = null;
            if (ds.Tables[0].Rows.Count > 0)
            {
                int items_count = ds.Tables[0].Rows[0].ItemArray.Length;
                if (items_count > 0)
                {
                    readingData = new string[items_count];
                    int counter = 0;
                    foreach (var item in ds.Tables[0].Rows[0].ItemArray)
                    {
                        readingData[counter] = item.ToString();
                        counter++;
                    }
                }
            }
            return readingData;
        }

        void fillTableWithCurrentDateData(string[] lastReadingData)
        {
            string current_date = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
            int except_hour = 0;
            if (lastReadingData.Count() > 0)
            {
                except_hour = int.Parse(lastReadingData[2]);
            }

            string qrt = "select * from Readings where working_date='" + current_date + "' order by val(working_hour)";

            DataSet ds = DBFunctions.fillDataSet(qrt);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                if (except_hour > 0 && i < except_hour)
                {
                    for (int j = 0; j < 61; j++)
                    {
                        if (j == 0)
                            j = 1;
                        int temp;
                        int.TryParse(ds.Tables[0].Rows[i].ItemArray[j + 2].ToString(), out temp);
                        dataGridView1[j, i].Value = temp;
                    }
                }
            }
        }

        void setInitialValuesForStartInputs(string[] lastReadingData)
        {
            if (lastReadingData != null)
            {
                if (lastReadingData.Count() > 0)
                {
                    numericUpDown1.Value = int.Parse(lastReadingData[2]);
                    TextBox line = (TextBox)this.Controls.Find("line01old", true).First();
                    line.Text = lastReadingData[3];
                    line = (TextBox)this.Controls.Find("line02old", true).First();
                    line.Text = lastReadingData[4];
                    line = (TextBox)this.Controls.Find("line03old", true).First();
                    line.Text = lastReadingData[5];
                    line = (TextBox)this.Controls.Find("line04old", true).First();
                    line.Text = lastReadingData[6];
                    line = (TextBox)this.Controls.Find("line11old", true).First();
                    line.Text = lastReadingData[7];
                    line = (TextBox)this.Controls.Find("line12old", true).First();
                    line.Text = lastReadingData[8];
                    line = (TextBox)this.Controls.Find("line13old", true).First();
                    line.Text = lastReadingData[9];
                    line = (TextBox)this.Controls.Find("line14old", true).First();
                    line.Text = lastReadingData[10];
                    line = (TextBox)this.Controls.Find("line21old", true).First();
                    line.Text = lastReadingData[11];
                    line = (TextBox)this.Controls.Find("line22old", true).First();
                    line.Text = lastReadingData[12];
                    line = (TextBox)this.Controls.Find("line23old", true).First();
                    line.Text = lastReadingData[13];
                    line = (TextBox)this.Controls.Find("line24old", true).First();
                    line.Text = lastReadingData[14];
                    line = (TextBox)this.Controls.Find("line31old", true).First();
                    line.Text = lastReadingData[15];
                    line = (TextBox)this.Controls.Find("line32old", true).First();
                    line.Text = lastReadingData[16];
                    line = (TextBox)this.Controls.Find("line33old", true).First();
                    line.Text = lastReadingData[17];
                    line = (TextBox)this.Controls.Find("line34old", true).First();
                    line.Text = lastReadingData[18];
                    line = (TextBox)this.Controls.Find("line41old", true).First();
                    line.Text = lastReadingData[19];
                    line = (TextBox)this.Controls.Find("line42old", true).First();
                    line.Text = lastReadingData[20];
                    line = (TextBox)this.Controls.Find("line43old", true).First();
                    line.Text = lastReadingData[21];
                    line = (TextBox)this.Controls.Find("line44old", true).First();
                    line.Text = lastReadingData[22];
                    line = (TextBox)this.Controls.Find("line51old", true).First();
                    line.Text = lastReadingData[23];
                    line = (TextBox)this.Controls.Find("line52old", true).First();
                    line.Text = lastReadingData[24];
                    line = (TextBox)this.Controls.Find("line53old", true).First();
                    line.Text = lastReadingData[25];
                    line = (TextBox)this.Controls.Find("line54old", true).First();
                    line.Text = lastReadingData[26];
                    line = (TextBox)this.Controls.Find("line61old", true).First();
                    line.Text = lastReadingData[27];
                    line = (TextBox)this.Controls.Find("line62old", true).First();
                    line.Text = lastReadingData[28];
                    line = (TextBox)this.Controls.Find("line63old", true).First();
                    line.Text = lastReadingData[29];
                    line = (TextBox)this.Controls.Find("line64old", true).First();
                    line.Text = lastReadingData[30];
                    line = (TextBox)this.Controls.Find("line71old", true).First();
                    line.Text = lastReadingData[31];
                    line = (TextBox)this.Controls.Find("line72old", true).First();
                    line.Text = lastReadingData[32];
                    line = (TextBox)this.Controls.Find("line73old", true).First();
                    line.Text = lastReadingData[33];
                    line = (TextBox)this.Controls.Find("line74old", true).First();
                    line.Text = lastReadingData[34];
                    line = (TextBox)this.Controls.Find("line81old", true).First();
                    line.Text = lastReadingData[35];
                    line = (TextBox)this.Controls.Find("line82old", true).First();
                    line.Text = lastReadingData[36];
                    line = (TextBox)this.Controls.Find("line83old", true).First();
                    line.Text = lastReadingData[37];
                    line = (TextBox)this.Controls.Find("line84old", true).First();
                    line.Text = lastReadingData[38];
                    line = (TextBox)this.Controls.Find("line91old", true).First();
                    line.Text = lastReadingData[39];
                    line = (TextBox)this.Controls.Find("line92old", true).First();
                    line.Text = lastReadingData[40];
                    line = (TextBox)this.Controls.Find("line93old", true).First();
                    line.Text = lastReadingData[41];
                    line = (TextBox)this.Controls.Find("line94old", true).First();
                    line.Text = lastReadingData[42];
                    line = (TextBox)this.Controls.Find("line101old", true).First();
                    line.Text = lastReadingData[43];
                    line = (TextBox)this.Controls.Find("line102old", true).First();
                    line.Text = lastReadingData[44];
                    line = (TextBox)this.Controls.Find("line103old", true).First();
                    line.Text = lastReadingData[45];
                    line = (TextBox)this.Controls.Find("line104old", true).First();
                    line.Text = lastReadingData[46];
                    line = (TextBox)this.Controls.Find("line111old", true).First();
                    line.Text = lastReadingData[47];
                    line = (TextBox)this.Controls.Find("line112old", true).First();
                    line.Text = lastReadingData[48];
                    line = (TextBox)this.Controls.Find("line113old", true).First();
                    line.Text = lastReadingData[49];
                    line = (TextBox)this.Controls.Find("line114old", true).First();
                    line.Text = lastReadingData[50];
                    line = (TextBox)this.Controls.Find("line121old", true).First();
                    line.Text = lastReadingData[51];
                    line = (TextBox)this.Controls.Find("line122old", true).First();
                    line.Text = lastReadingData[52];
                    line = (TextBox)this.Controls.Find("line123old", true).First();
                    line.Text = lastReadingData[53];
                    line = (TextBox)this.Controls.Find("line124old", true).First();
                    line.Text = lastReadingData[54];
                    line = (TextBox)this.Controls.Find("line131old", true).First();
                    line.Text = lastReadingData[55];
                    line = (TextBox)this.Controls.Find("line132old", true).First();
                    line.Text = lastReadingData[56];
                    line = (TextBox)this.Controls.Find("line133old", true).First();
                    line.Text = lastReadingData[57];
                    line = (TextBox)this.Controls.Find("line134old", true).First();
                    line.Text = lastReadingData[58];

                    numericUpDown2.Value = int.Parse(lastReadingData[2]);

                    // clear current values
                    line = (TextBox)this.Controls.Find("line01current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line02current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line03current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line04current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line11current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line12current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line13current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line14current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line21current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line22current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line23current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line24current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line31current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line32current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line33current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line34current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line41current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line42current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line43current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line44current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line51current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line52current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line53current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line54current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line61current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line62current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line63current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line64current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line71current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line72current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line73current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line74current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line81current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line82current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line83current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line84current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line91current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line92current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line93current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line94current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line101current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line102current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line103current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line104current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line111current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line112current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line113current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line114current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line121current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line122current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line123current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line124current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line131current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line132current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line133current", true).First();
                    line.Text = "";
                    line = (TextBox)this.Controls.Find("line134current", true).First();
                    line.Text = "";
                }
            }
            else
            {
                string[] previousDayReadingDate = getPreviousDayreadingValues();
                numericUpDown1.Value = 0;
                TextBox line;
                if (previousDayReadingDate != null)
                {
                    line = (TextBox)this.Controls.Find("line01old", true).First();
                    line.Text = previousDayReadingDate[3];
                    line = (TextBox)this.Controls.Find("line02old", true).First();
                    line.Text = previousDayReadingDate[4];
                    line = (TextBox)this.Controls.Find("line03old", true).First();
                    line.Text = previousDayReadingDate[5];
                    line = (TextBox)this.Controls.Find("line04old", true).First();
                    line.Text = previousDayReadingDate[6];
                    line = (TextBox)this.Controls.Find("line11old", true).First();
                    line.Text = previousDayReadingDate[7];
                    line = (TextBox)this.Controls.Find("line12old", true).First();
                    line.Text = previousDayReadingDate[8];
                    line = (TextBox)this.Controls.Find("line13old", true).First();
                    line.Text = previousDayReadingDate[9];
                    line = (TextBox)this.Controls.Find("line14old", true).First();
                    line.Text = previousDayReadingDate[10];
                    line = (TextBox)this.Controls.Find("line21old", true).First();
                    line.Text = previousDayReadingDate[11];
                    line = (TextBox)this.Controls.Find("line22old", true).First();
                    line.Text = previousDayReadingDate[12];
                    line = (TextBox)this.Controls.Find("line23old", true).First();
                    line.Text = previousDayReadingDate[13];
                    line = (TextBox)this.Controls.Find("line24old", true).First();
                    line.Text = previousDayReadingDate[14];
                    line = (TextBox)this.Controls.Find("line31old", true).First();
                    line.Text = previousDayReadingDate[15];
                    line = (TextBox)this.Controls.Find("line32old", true).First();
                    line.Text = previousDayReadingDate[16];
                    line = (TextBox)this.Controls.Find("line33old", true).First();
                    line.Text = previousDayReadingDate[17];
                    line = (TextBox)this.Controls.Find("line34old", true).First();
                    line.Text = previousDayReadingDate[18];
                    line = (TextBox)this.Controls.Find("line41old", true).First();
                    line.Text = previousDayReadingDate[19];
                    line = (TextBox)this.Controls.Find("line42old", true).First();
                    line.Text = previousDayReadingDate[20];
                    line = (TextBox)this.Controls.Find("line43old", true).First();
                    line.Text = previousDayReadingDate[21];
                    line = (TextBox)this.Controls.Find("line44old", true).First();
                    line.Text = previousDayReadingDate[22];
                    line = (TextBox)this.Controls.Find("line51old", true).First();
                    line.Text = previousDayReadingDate[23];
                    line = (TextBox)this.Controls.Find("line52old", true).First();
                    line.Text = previousDayReadingDate[24];
                    line = (TextBox)this.Controls.Find("line53old", true).First();
                    line.Text = previousDayReadingDate[25];
                    line = (TextBox)this.Controls.Find("line54old", true).First();
                    line.Text = previousDayReadingDate[26];
                    line = (TextBox)this.Controls.Find("line61old", true).First();
                    line.Text = previousDayReadingDate[27];
                    line = (TextBox)this.Controls.Find("line62old", true).First();
                    line.Text = previousDayReadingDate[28];
                    line = (TextBox)this.Controls.Find("line63old", true).First();
                    line.Text = previousDayReadingDate[29];
                    line = (TextBox)this.Controls.Find("line64old", true).First();
                    line.Text = previousDayReadingDate[30];
                    line = (TextBox)this.Controls.Find("line71old", true).First();
                    line.Text = previousDayReadingDate[31];
                    line = (TextBox)this.Controls.Find("line72old", true).First();
                    line.Text = previousDayReadingDate[32];
                    line = (TextBox)this.Controls.Find("line73old", true).First();
                    line.Text = previousDayReadingDate[33];
                    line = (TextBox)this.Controls.Find("line74old", true).First();
                    line.Text = previousDayReadingDate[34];
                    line = (TextBox)this.Controls.Find("line81old", true).First();
                    line.Text = previousDayReadingDate[35];
                    line = (TextBox)this.Controls.Find("line82old", true).First();
                    line.Text = previousDayReadingDate[36];
                    line = (TextBox)this.Controls.Find("line83old", true).First();
                    line.Text = previousDayReadingDate[37];
                    line = (TextBox)this.Controls.Find("line84old", true).First();
                    line.Text = previousDayReadingDate[38];
                    line = (TextBox)this.Controls.Find("line91old", true).First();
                    line.Text = previousDayReadingDate[39];
                    line = (TextBox)this.Controls.Find("line92old", true).First();
                    line.Text = previousDayReadingDate[40];
                    line = (TextBox)this.Controls.Find("line93old", true).First();
                    line.Text = previousDayReadingDate[41];
                    line = (TextBox)this.Controls.Find("line94old", true).First();
                    line.Text = previousDayReadingDate[42];
                    line = (TextBox)this.Controls.Find("line101old", true).First();
                    line.Text = previousDayReadingDate[43];
                    line = (TextBox)this.Controls.Find("line102old", true).First();
                    line.Text = previousDayReadingDate[44];
                    line = (TextBox)this.Controls.Find("line103old", true).First();
                    line.Text = previousDayReadingDate[45];
                    line = (TextBox)this.Controls.Find("line104old", true).First();
                    line.Text = previousDayReadingDate[46];
                    line = (TextBox)this.Controls.Find("line111old", true).First();
                    line.Text = previousDayReadingDate[47];
                    line = (TextBox)this.Controls.Find("line112old", true).First();
                    line.Text = previousDayReadingDate[48];
                    line = (TextBox)this.Controls.Find("line113old", true).First();
                    line.Text = previousDayReadingDate[49];
                    line = (TextBox)this.Controls.Find("line114old", true).First();
                    line.Text = previousDayReadingDate[50];
                    line = (TextBox)this.Controls.Find("line121old", true).First();
                    line.Text = previousDayReadingDate[51];
                    line = (TextBox)this.Controls.Find("line122old", true).First();
                    line.Text = previousDayReadingDate[52];
                    line = (TextBox)this.Controls.Find("line123old", true).First();
                    line.Text = previousDayReadingDate[53];
                    line = (TextBox)this.Controls.Find("line124old", true).First();
                    line.Text = previousDayReadingDate[54];
                    line = (TextBox)this.Controls.Find("line131old", true).First();
                    line.Text = previousDayReadingDate[55];
                    line = (TextBox)this.Controls.Find("line132old", true).First();
                    line.Text = previousDayReadingDate[56];
                    line = (TextBox)this.Controls.Find("line133old", true).First();
                    line.Text = previousDayReadingDate[57];
                    line = (TextBox)this.Controls.Find("line134old", true).First();

                }
                else
                {
                    line = (TextBox)this.Controls.Find("line01old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line02old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line03old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line04old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line11old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line12old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line13old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line14old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line21old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line22old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line23old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line24old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line31old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line32old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line33old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line34old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line41old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line42old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line43old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line44old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line51old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line52old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line53old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line54old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line61old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line62old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line63old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line64old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line71old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line72old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line73old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line74old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line81old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line82old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line83old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line84old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line91old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line92old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line93old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line94old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line101old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line102old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line103old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line104old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line111old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line112old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line113old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line114old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line121old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line122old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line123old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line124old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line131old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line132old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line133old", true).First();
                    line.Text = line.Text = "";
                    line = (TextBox)this.Controls.Find("line134old", true).First();
                }
                numericUpDown2.Value = 1;

                // clear current values
                line = (TextBox)this.Controls.Find("line01current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line02current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line03current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line04current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line11current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line12current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line13current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line14current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line21current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line22current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line23current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line24current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line31current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line32current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line33current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line34current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line41current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line42current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line43current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line44current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line51current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line52current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line53current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line54current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line61current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line62current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line63current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line64current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line71current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line72current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line73current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line74current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line81current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line82current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line83current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line84current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line91current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line92current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line93current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line94current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line101current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line102current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line103current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line104current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line111current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line112current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line113current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line114current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line121current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line122current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line123current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line124current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line131current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line132current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line133current", true).First();
                line.Text = "";
                line = (TextBox)this.Controls.Find("line134current", true).First();
                line.Text = "";
            }

        }

        void copyGroupBox(int index, string gb_p_name, FlowLayoutPanel panel1, string extension = "")
        {
            GroupBox globlaGroupBox = new GroupBox();

            // globlaGroupBox
            globlaGroupBox.Location = new System.Drawing.Point(523, 174);
            globlaGroupBox.Name = "gb" + index + extension;
            globlaGroupBox.Size = new System.Drawing.Size(282, 75);
            globlaGroupBox.TabIndex = 8;
            globlaGroupBox.TabStop = false;
            globlaGroupBox.Text = gb_p_name;

            // textbox1
            textBoxs[0 + index] = new TextBox();
            textBoxs[0 + index].Location = new System.Drawing.Point(212, 39);
            textBoxs[0 + index].Name = "line" + (index / 4) + 1 + extension;
            textBoxs[0 + index].AccessibleName = "line" + (index / 4) + 1 + extension;
            textBoxs[0 + index].Size = new System.Drawing.Size(61, 20);
            textBoxs[0 + index].TabIndex = 2;
            // label1
            labels[0 + index] = new Label();
            labels[0 + index].AutoSize = true;
            labels[0 + index].Location = new System.Drawing.Point(221, 23);
            labels[0 + index].Name = "lb" + index + extension;
            labels[0 + index].Size = new System.Drawing.Size(35, 13);
            labels[0 + index].TabIndex = 11;
            labels[0 + index].Text = "ارسال";

            // textbox2
            textBoxs[1 + index] = new TextBox();
            textBoxs[1 + index].Location = new System.Drawing.Point(145, 39);
            textBoxs[1 + index].Name = "line" + (index / 4) + 2 + extension;
            textBoxs[1 + index].AccessibleName = "line" + (index / 4) + 2 + extension;
            textBoxs[1 + index].Size = new System.Drawing.Size(61, 20);
            textBoxs[1 + index].TabIndex = 5;
            // label2
            labels[1 + index] = new Label();
            labels[1 + index].AutoSize = true;
            labels[1 + index].Location = new System.Drawing.Point(154, 23);
            labels[1 + index].Name = "lb" + (index + 1) + extension;
            labels[1 + index].Size = new System.Drawing.Size(43, 13);
            labels[1 + index].TabIndex = 9;
            labels[1 + index].Text = "استهلاك";

            // textbox3
            textBoxs[2 + index] = new TextBox();
            textBoxs[2 + index].Location = new System.Drawing.Point(78, 39);
            textBoxs[2 + index].Name = "line" + (index / 4) + 3 + extension;
            textBoxs[2 + index].AccessibleName = "line" + (index / 4) + 3 + extension;
            textBoxs[2 + index].Size = new System.Drawing.Size(61, 20);
            textBoxs[2 + index].TabIndex = 10;
            // label3
            labels[2 + index] = new Label();
            labels[2 + index].AutoSize = true;
            labels[2 + index].Location = new System.Drawing.Point(87, 23);
            labels[2 + index].Name = "lb" + (index + 2) + extension;
            labels[2 + index].Size = new System.Drawing.Size(35, 13);
            labels[2 + index].TabIndex = 6;
            labels[2 + index].Text = "استقبال";

            // textbox4
            textBoxs[3 + index] = new TextBox();
            textBoxs[3 + index].Location = new System.Drawing.Point(6, 39);
            textBoxs[3 + index].Name = "line" + (index / 4) + 4 + extension;
            textBoxs[1 + index].AccessibleName = "line" + (index / 4) + 4 + extension;
            textBoxs[3 + index].Size = new System.Drawing.Size(61, 20);
            textBoxs[3 + index].TabIndex = 7;
            // label4
            labels[3 + index] = new Label();
            labels[3 + index].AutoSize = true;
            labels[3 + index].Location = new System.Drawing.Point(20, 23);
            labels[3 + index].Name = "lb" + (index + 3) + extension;
            labels[3 + index].Size = new System.Drawing.Size(43, 13);
            labels[3 + index].TabIndex = 4;
            labels[3 + index].Text = "استهلاك";

            Controls.Add(globlaGroupBox);

            globlaGroupBox.Controls.Add(textBoxs[0 + index]);
            globlaGroupBox.Controls.Add(textBoxs[1 + index]);
            globlaGroupBox.Controls.Add(textBoxs[2 + index]);
            globlaGroupBox.Controls.Add(textBoxs[3 + index]);
            globlaGroupBox.Controls.Add(labels[0 + index]);
            globlaGroupBox.Controls.Add(labels[1 + index]);
            globlaGroupBox.Controls.Add(labels[2 + index]);
            globlaGroupBox.Controls.Add(labels[3 + index]);

            panel1.Controls.Add(globlaGroupBox);

        }

        private void button1_Click(object sender, System.EventArgs e)
        {

            //DataRow dr;
            //dr = readingTable.NewRow();
            //dr["banias1_send"] = this.Controls.Find("tb0", true).First().Text;
            //dr["banias1_send_subscription"] = this.Controls.Find("tb1", true).First().Text;
            //dr["banias1_receive"] = this.Controls.Find("tb2", true).First().Text;
            //dr["banias1_receive_subscription"] = this.Controls.Find("tb3", true).First().Text;
            ///////
            //dr["banias2_send"] = this.Controls.Find("tb4", true).First().Text;
            //dr["banias2_send_subscription"] = this.Controls.Find("tb5", true).First().Text;
            //dr["banias2_receive"] = this.Controls.Find("tb6", true).First().Text;
            //dr["banias2_receive_subscription"] = this.Controls.Find("tb7", true).First().Text;
            ///////
            //dr["semerian1_send"] = this.Controls.Find("tb8", true).First().Text;
            //dr["semerian1_send_subscription"] = this.Controls.Find("tb9", true).First().Text;
            //dr["semerian1_receive"] = this.Controls.Find("tb10", true).First().Text;
            //dr["semerian1_receive_subscription"] = this.Controls.Find("tb11", true).First().Text;
            ///////
            //dr["semerian2_send"] = this.Controls.Find("tb12", true).First().Text;
            //dr["semerian2_send_subscription"] = this.Controls.Find("tb13", true).First().Text;
            //dr["semerian2_receive"] = this.Controls.Find("tb14", true).First().Text;
            //dr["semerian2_receive_subscription"] = this.Controls.Find("tb15", true).First().Text;
            ///////
            //dr["company_send"] = this.Controls.Find("tb16", true).First().Text;
            //dr["company_send_subscription"] = this.Controls.Find("tb17", true).First().Text;
            //dr["company_receive"] = this.Controls.Find("tb18", true).First().Text;
            //dr["company_receive_subscription"] = this.Controls.Find("tb19", true).First().Text;
            ///////
            //dr["amreet_send"] = this.Controls.Find("tb20", true).First().Text;
            //dr["amreet_send_subscription"] = this.Controls.Find("tb21", true).First().Text;
            //dr["amreet_receive"] = this.Controls.Find("tb22", true).First().Text;
            //dr["amreet_receive_subscription"] = this.Controls.Find("tb23", true).First().Text;
            ///////
            //dr["north_send"] = this.Controls.Find("tb24", true).First().Text;
            //dr["north_send_subscription"] = this.Controls.Find("tb25", true).First().Text;
            //dr["north_receive"] = this.Controls.Find("tb26", true).First().Text;
            //dr["north_receive_subscription"] = this.Controls.Find("tb27", true).First().Text;
            ///////
            //dr["esmant_send"] = this.Controls.Find("tb28", true).First().Text;
            //dr["esmant_send_subscription"] = this.Controls.Find("tb29", true).First().Text;
            //dr["esmant_receive"] = this.Controls.Find("tb30", true).First().Text;
            //dr["esmant_receive_subscription"] = this.Controls.Find("tb31", true).First().Text;
            ///////
            //dr["arrive1_send"] = this.Controls.Find("tb32", true).First().Text;
            //dr["arrive1_send_subscription"] = this.Controls.Find("tb33", true).First().Text;
            //dr["arrive1_receive"] = this.Controls.Find("tb34", true).First().Text;
            //dr["arrive1_receive_subscription"] = this.Controls.Find("tb35", true).First().Text;
            ///////
            //dr["arrive2_send"] = this.Controls.Find("tb36", true).First().Text;
            //dr["arrive2_send_subscription"] = this.Controls.Find("tb37", true).First().Text;
            //dr["arrive2_receive"] = this.Controls.Find("tb38", true).First().Text;
            //dr["arrive2_receive_subscription"] = this.Controls.Find("tb39", true).First().Text;
            ///////
            //dr["arrive3_send"] = this.Controls.Find("tb40", true).First().Text;
            //dr["arrive3_send_subscription"] = this.Controls.Find("tb41", true).First().Text;
            //dr["arrive3_receive"] = this.Controls.Find("tb42", true).First().Text;
            //dr["arrive3_receive_subscription"] = this.Controls.Find("tb43", true).First().Text;
            ///////
            //dr["transformer1_send"] = this.Controls.Find("tb44", true).First().Text;
            //dr["transformer1_send_subscription"] = this.Controls.Find("tb45", true).First().Text;
            //dr["transformer1_receive"] = this.Controls.Find("tb46", true).First().Text;
            //dr["transformer1_receive_subscription"] = this.Controls.Find("tb47", true).First().Text;
            ///////
            //dr["transformer2_send"] = this.Controls.Find("tb48", true).First().Text;
            //dr["transformer2_send_subscription"] = this.Controls.Find("tb49", true).First().Text;
            //dr["transformer2_receive"] = this.Controls.Find("tb50", true).First().Text;
            //dr["transformer2_receive_subscription"] = this.Controls.Find("tb51", true).First().Text;
            ///////
            //dr["transformer3_send"] = this.Controls.Find("tb52", true).First().Text;
            //dr["transformer3_send_subscription"] = this.Controls.Find("tb53", true).First().Text;
            //dr["transformer3_receive"] = this.Controls.Find("tb54", true).First().Text;
            //dr["transformer3_receive_subscription"] = this.Controls.Find("tb55", true).First().Text;

            //readingTable.Rows.Add(dr);



        }

        void createReadingTable()
        {
            readingTable = new DataTable("readings");

            // add id column
            addColumnToTable(readingTable, "hour", "hour", "string");

            // Create Banias1 column.
            addColumnToTable(readingTable, "banias1_send", "Banias1 Send", "string");
            addColumnToTable(readingTable, "banias1_send_subscription", "Banias1 Send Subscription", "string");
            addColumnToTable(readingTable, "banias1_receive", "Banias1 Receive", "string");
            addColumnToTable(readingTable, "banias1_receive_subscription", "Banias1 Receive Subscription", "string");

            // Create Banias2 column.
            addColumnToTable(readingTable, "banias2_send", "Banias2 Send", "string");
            addColumnToTable(readingTable, "banias2_send_subscription", "Banias2 Send Subscription", "string");
            addColumnToTable(readingTable, "banias2_receive", "Banias2 Receive", "string");
            addColumnToTable(readingTable, "banias2_receive_subscription", "Banias2 Receive Subscription", "string");

            // Create Semerian1 column.
            addColumnToTable(readingTable, "semerian1_send", "Semerian1 Send", "string");
            addColumnToTable(readingTable, "semerian1_send_subscription", "Semerian1 Send Subscription", "string");
            addColumnToTable(readingTable, "semerian1_receive", "Semerian1 Receive", "string");
            addColumnToTable(readingTable, "semerian1_receive_subscription", "Semerian1 Receive Subscription", "string");

            // Create Semerian2 column.
            addColumnToTable(readingTable, "semerian2_send", "Semerian2 Send", "string");
            addColumnToTable(readingTable, "semerian2_send_subscription", "Semerian2 Send Subscription", "string");
            addColumnToTable(readingTable, "semerian2_receive", "Semerian2 Receive", "string");
            addColumnToTable(readingTable, "semerian2_receive_subscription", "Semerian2 Receive Subscription", "string");

            // Create Arrive1 column.
            addColumnToTable(readingTable, "arrive1_send", "Arrive1 Send", "string");
            addColumnToTable(readingTable, "arrive1_send_subscription", "Arrive1 Send Subscription", "string");
            addColumnToTable(readingTable, "arrive1_receive", "Arrive1 Receive", "string");
            addColumnToTable(readingTable, "arrive1_receive_subscription", "Arrive1 Receive Subscription", "string");

            // Create Arrive2 column.
            addColumnToTable(readingTable, "arrive2_send", "Arrive2 Send", "string");
            addColumnToTable(readingTable, "arrive2_send_subscription", "Arrive2 Send Subscription", "string");
            addColumnToTable(readingTable, "arrive2_receive", "Arrive2 Receive", "string");
            addColumnToTable(readingTable, "arrive2_receive_subscription", "Arrive2 Receive Subscription", "string");

            // Create Arrive3 column.
            addColumnToTable(readingTable, "arrive3_send", "Arrive3 Send", "string");
            addColumnToTable(readingTable, "arrive3_send_subscription", "Arrive3 Send Subscription", "string");
            addColumnToTable(readingTable, "arrive3_receive", "Arrive3 Receive", "string");
            addColumnToTable(readingTable, "arrive3_receive_subscription", "Arrive3 Receive Subscription", "string");

            // Create Amreet column.
            addColumnToTable(readingTable, "amreet_send", "Amreet Send", "string");
            addColumnToTable(readingTable, "amreet_send_subscription", "Amreet Send Subscription", "string");
            addColumnToTable(readingTable, "amreet_receive", "Amreet Receive", "string");
            addColumnToTable(readingTable, "amreet_receive_subscription", "Amreet Receive Subscription", "string");

            // Create Esmant column.
            addColumnToTable(readingTable, "esmant_send", "Esmant Send", "string");
            addColumnToTable(readingTable, "esmant_send_subscription", "Esmant Send Subscription", "string");
            addColumnToTable(readingTable, "esmant_receive", "Esmant Receive", "string");
            addColumnToTable(readingTable, "esmant_receive_subscription", "Esmant Receive Subscription", "string");

            // Create Company column.
            addColumnToTable(readingTable, "company_send", "Company Send", "string");
            addColumnToTable(readingTable, "company_send_subscription", "Company Send Subscription", "string");
            addColumnToTable(readingTable, "company_receive", "Company Receive", "string");
            addColumnToTable(readingTable, "company_receive_subscription", "Company Receive Subscription", "string");

            // Create Transformer1 column.
            addColumnToTable(readingTable, "transformer1_send", "Transformer1 Send", "string");
            addColumnToTable(readingTable, "transformer1_send_subscription", "Transformer1 Send Subscription", "string");
            addColumnToTable(readingTable, "transformer1_receive", "Transformer1 Receive", "string");
            addColumnToTable(readingTable, "transformer1_receive_subscription", "Transformer1 Receive Subscription", "string");

            // Create Transformer2 column.
            addColumnToTable(readingTable, "transformer2_send", "Transformer2 Send", "string");
            addColumnToTable(readingTable, "transformer2_send_subscription", "Transformer2 Send Subscription", "string");
            addColumnToTable(readingTable, "transformer2_receive", "Transformer2 Receive", "string");
            addColumnToTable(readingTable, "transformer2_receive_subscription", "Transformer2 Receive Subscription", "string");

            // Create Transformer3 column.
            addColumnToTable(readingTable, "transformer3_send", "Transformer3 Send", "string");
            addColumnToTable(readingTable, "transformer3_send_subscription", "Transformer3 Send Subscription", "string");
            addColumnToTable(readingTable, "transformer3_receive", "Transformer3 Receive", "string");
            addColumnToTable(readingTable, "transformer3_receive_subscription", "Transformer3 Receive Subscription", "string");

            // Create North column.
            addColumnToTable(readingTable, "north_send", "North Send", "string");
            addColumnToTable(readingTable, "north_send_subscription", "North Send Subscription", "string");
            addColumnToTable(readingTable, "north_receive", "North Receive", "string");
            addColumnToTable(readingTable, "north_receive_subscription", "North Receive Subscription", "string");

            // Create tartous column.
            addColumnToTable(readingTable, "tartous1", "tartous1", "string");
            addColumnToTable(readingTable, "tartous2", "tartous2", "string");
            addColumnToTable(readingTable, "tartous3", "tartous3", "string");
            addColumnToTable(readingTable, "tartous4", "tartous4", "string");

            // Create total column.
            addColumnToTable(readingTable, "total1", "total1", "string");
            addColumnToTable(readingTable, "total2", "total2", "string");
            addColumnToTable(readingTable, "total3", "total3", "string");
            addColumnToTable(readingTable, "total4", "total4", "string");
        }

        void addColumnToTable(DataTable tableName, string captionName, string columnName, string type)
        {
            DataColumn dtColumn;
            dtColumn = new DataColumn();
            if (type == "string")
            {
                dtColumn.DataType = typeof(String);
                dtColumn.ColumnName = columnName;
                dtColumn.Caption = captionName;
                dtColumn.AutoIncrement = false;
                dtColumn.ReadOnly = false;
                dtColumn.Unique = false;
            }
            else
            {
                dtColumn.DataType = System.Type.GetType("System.Int32");
                dtColumn.AutoIncrement = true;
                dtColumn.AutoIncrementSeed = 1;
                dtColumn.AutoIncrementStep = 1;
                dtColumn.Caption = captionName;
                dtColumn.ColumnName = columnName;
            }

            /// Add column to the DataColumnCollection.
            tableName.Columns.Add(dtColumn);
        }

        void fillTableEmptyRows(DataTable tableName, int rowsCount)
        {
            DataRow dr;
            for (int i = 0; i < rowsCount; i++)
            {
                dr = tableName.NewRow();
                if (i < rowsCount - 1)
                    dr[0] = (i + 1);
                tableName.Rows.Add(dr);
            }

            dataGridView1.DataSource = tableName;
            dataGridView1.Columns[0].Width = 70;
            dataGridView1.Columns[0].HeaderText = "الساعة";
            //////////////////////////////////////////// banias1 ////////////////////////////////////////////
            dataGridView1.Columns[1].Width = 80;
            dataGridView1.Columns[1].HeaderText = "بانياس1 \n\n ارسال";
            dataGridView1.Columns[2].HeaderText = "استهلاك";
            dataGridView1.Columns[2].Width = 70;
            dataGridView1.Columns[3].Width = 80;
            dataGridView1.Columns[3].HeaderText = "بانياس1 \n\n استقبال";
            dataGridView1.Columns[4].HeaderText = "استهلاك";
            dataGridView1.Columns[4].Width = 70;
            dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[2].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[3].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[4].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            //////////////////////////////////////////// banias2 ////////////////////////////////////////////
            dataGridView1.Columns[5].Width = 80;
            dataGridView1.Columns[5].HeaderText = "بانياس2 \n\n ارسال";
            dataGridView1.Columns[6].HeaderText = "استهلاك";
            dataGridView1.Columns[6].Width = 70;
            dataGridView1.Columns[7].Width = 80;
            dataGridView1.Columns[7].HeaderText = "بانياس2 \n\n استقبال";
            dataGridView1.Columns[8].HeaderText = "استهلاك";
            dataGridView1.Columns[8].Width = 70;
            dataGridView1.Columns[5].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[6].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[7].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[8].DefaultCellStyle.BackColor = Color.AliceBlue;
            //////////////////////////////////////////// semerian1 ////////////////////////////////////////////
            dataGridView1.Columns[9].Width = 80;
            dataGridView1.Columns[9].HeaderText = "سمريان1 \n\n ارسال";
            dataGridView1.Columns[10].HeaderText = "استهلاك";
            dataGridView1.Columns[10].Width = 70;
            dataGridView1.Columns[11].Width = 80;
            dataGridView1.Columns[11].HeaderText = "سمريان1 \n\n استقبال";
            dataGridView1.Columns[12].HeaderText = "استهلاك";
            dataGridView1.Columns[12].Width = 70;
            dataGridView1.Columns[9].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[10].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[11].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[12].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            //////////////////////////////////////////// semerian2 ////////////////////////////////////////////
            dataGridView1.Columns[13].Width = 80;
            dataGridView1.Columns[13].HeaderText = "سمريان2 \n\n ارسال";
            dataGridView1.Columns[14].HeaderText = "استهلاك";
            dataGridView1.Columns[14].Width = 70;
            dataGridView1.Columns[15].Width = 80;
            dataGridView1.Columns[15].HeaderText = "سمريان2 \n\n استقبال";
            dataGridView1.Columns[16].HeaderText = "استهلاك";
            dataGridView1.Columns[16].Width = 70;
            dataGridView1.Columns[13].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[14].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[15].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[16].DefaultCellStyle.BackColor = Color.AliceBlue;
            //////////////////////////////////////////// arrival1 ////////////////////////////////////////////
            dataGridView1.Columns[17].Width = 80;
            dataGridView1.Columns[17].HeaderText = "وصول1 \n\n ارسال";
            dataGridView1.Columns[18].HeaderText = "استهلاك";
            dataGridView1.Columns[18].Width = 70;
            dataGridView1.Columns[19].Width = 80;
            dataGridView1.Columns[19].HeaderText = "وصول1 \n\n استقبال";
            dataGridView1.Columns[20].HeaderText = "استهلاك";
            dataGridView1.Columns[20].Width = 70;
            dataGridView1.Columns[17].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[18].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[19].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[20].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            //////////////////////////////////////////// arrival2 ////////////////////////////////////////////
            dataGridView1.Columns[21].Width = 80;
            dataGridView1.Columns[21].HeaderText = "وصول2 \n\n ارسال";
            dataGridView1.Columns[22].HeaderText = "استهلاك";
            dataGridView1.Columns[22].Width = 70;
            dataGridView1.Columns[23].Width = 80;
            dataGridView1.Columns[23].HeaderText = "وصول2 \n\n استقبال";
            dataGridView1.Columns[24].HeaderText = "استهلاك";
            dataGridView1.Columns[24].Width = 70;
            dataGridView1.Columns[21].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[22].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[23].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[24].DefaultCellStyle.BackColor = Color.AliceBlue;
            //////////////////////////////////////////// arrival3 ////////////////////////////////////////////
            dataGridView1.Columns[25].Width = 80;
            dataGridView1.Columns[25].HeaderText = "وصول3 \n\n ارسال";
            dataGridView1.Columns[26].HeaderText = "استهلاك";
            dataGridView1.Columns[26].Width = 70;
            dataGridView1.Columns[27].Width = 80;
            dataGridView1.Columns[27].HeaderText = "وصول3 \n\n استقبال";
            dataGridView1.Columns[28].HeaderText = "استهلاك";
            dataGridView1.Columns[28].Width = 70;
            dataGridView1.Columns[25].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[26].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[27].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[28].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            //////////////////////////////////////////// amreet ////////////////////////////////////////////
            dataGridView1.Columns[29].Width = 80;
            dataGridView1.Columns[29].HeaderText = "عمريت \n\n ارسال";
            dataGridView1.Columns[30].HeaderText = "استهلاك";
            dataGridView1.Columns[30].Width = 70;
            dataGridView1.Columns[31].Width = 80;
            dataGridView1.Columns[31].HeaderText = "عمريت \n\n استقبال";
            dataGridView1.Columns[32].HeaderText = "استهلاك";
            dataGridView1.Columns[32].Width = 70;
            dataGridView1.Columns[29].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[30].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[31].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[32].DefaultCellStyle.BackColor = Color.AliceBlue;
            //////////////////////////////////////////// esmant ////////////////////////////////////////////
            dataGridView1.Columns[33].Width = 80;
            dataGridView1.Columns[33].HeaderText = "اسمنت \n\n ارسال";
            dataGridView1.Columns[34].HeaderText = "استهلاك";
            dataGridView1.Columns[34].Width = 70;
            dataGridView1.Columns[35].Width = 80;
            dataGridView1.Columns[35].HeaderText = "اسمنت \n\n استقبال";
            dataGridView1.Columns[36].HeaderText = "استهلاك";
            dataGridView1.Columns[36].Width = 70;
            dataGridView1.Columns[33].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[34].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[35].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[36].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            //////////////////////////////////////////// company ////////////////////////////////////////////
            dataGridView1.Columns[37].Width = 80;
            dataGridView1.Columns[37].HeaderText = "الشركة \n\n ارسال";
            dataGridView1.Columns[38].HeaderText = "استهلاك";
            dataGridView1.Columns[38].Width = 70;
            dataGridView1.Columns[39].Width = 80;
            dataGridView1.Columns[39].HeaderText = "الشركة \n\n استقبال";
            dataGridView1.Columns[40].HeaderText = "استهلاك";
            dataGridView1.Columns[40].Width = 70;
            dataGridView1.Columns[37].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[38].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[39].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[40].DefaultCellStyle.BackColor = Color.AliceBlue;
            //////////////////////////////////////////// transformer1 ////////////////////////////////////////////
            dataGridView1.Columns[41].Width = 80;
            dataGridView1.Columns[41].HeaderText = "محولة1 \n\n ارسال";
            dataGridView1.Columns[42].HeaderText = "استهلاك";
            dataGridView1.Columns[42].Width = 70;
            dataGridView1.Columns[43].Width = 80;
            dataGridView1.Columns[43].HeaderText = "محولة1 \n\n استقبال";
            dataGridView1.Columns[44].HeaderText = "استهلاك";
            dataGridView1.Columns[44].Width = 70;
            dataGridView1.Columns[41].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[42].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[43].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[44].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            //////////////////////////////////////////// transformer2 ////////////////////////////////////////////
            dataGridView1.Columns[45].Width = 80;
            dataGridView1.Columns[45].HeaderText = "محولة2 \n\n ارسال";
            dataGridView1.Columns[46].HeaderText = "استهلاك";
            dataGridView1.Columns[46].Width = 70;
            dataGridView1.Columns[47].Width = 80;
            dataGridView1.Columns[47].HeaderText = "محولة2 \n\n استقبال";
            dataGridView1.Columns[48].HeaderText = "استهلاك";
            dataGridView1.Columns[48].Width = 70;
            dataGridView1.Columns[45].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[46].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[47].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[48].DefaultCellStyle.BackColor = Color.AliceBlue;
            //////////////////////////////////////////// transformer3 ////////////////////////////////////////////
            dataGridView1.Columns[49].Width = 80;
            dataGridView1.Columns[49].HeaderText = "محولة3 \n\n ارسال";
            dataGridView1.Columns[50].HeaderText = "استهلاك";
            dataGridView1.Columns[50].Width = 70;
            dataGridView1.Columns[51].Width = 80;
            dataGridView1.Columns[51].HeaderText = "محولة3 \n\n استقبال";
            dataGridView1.Columns[52].HeaderText = "استهلاك";
            dataGridView1.Columns[52].Width = 70;
            dataGridView1.Columns[49].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[50].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[51].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[52].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            //////////////////////////////////////////// north ////////////////////////////////////////////
            dataGridView1.Columns[53].Width = 80;
            dataGridView1.Columns[53].HeaderText = "الشمال \n\n ارسال";
            dataGridView1.Columns[54].HeaderText = "استهلاك";
            dataGridView1.Columns[54].Width = 70;
            dataGridView1.Columns[55].Width = 80;
            dataGridView1.Columns[55].HeaderText = "الشمال \n\n استقبال";
            dataGridView1.Columns[56].HeaderText = "استهلاك";
            dataGridView1.Columns[56].Width = 70;
            dataGridView1.Columns[53].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[54].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[55].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[56].DefaultCellStyle.BackColor = Color.AliceBlue;
            //////////////////////////////////////////// tartous ////////////////////////////////////////////
            dataGridView1.Columns[57].Width = 80;
            dataGridView1.Columns[57].HeaderText = "بانياس";
            dataGridView1.Columns[58].HeaderText = "سمريان";
            dataGridView1.Columns[58].Width = 70;
            dataGridView1.Columns[59].Width = 80;
            dataGridView1.Columns[59].HeaderText = "طرطوس";
            dataGridView1.Columns[60].HeaderText = "";
            dataGridView1.Columns[60].Width = 70;
            dataGridView1.Columns[57].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[58].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[59].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            dataGridView1.Columns[60].DefaultCellStyle.BackColor = Color.LightSkyBlue;
            //////////////////////////////////////////// total ////////////////////////////////////////////
            dataGridView1.Columns[61].Width = 80;
            dataGridView1.Columns[61].HeaderText = "المجموع";
            dataGridView1.Columns[62].HeaderText = "";
            dataGridView1.Columns[62].Width = 70;
            dataGridView1.Columns[63].Width = 80;
            dataGridView1.Columns[63].HeaderText = "";
            dataGridView1.Columns[64].HeaderText = "";
            dataGridView1.Columns[64].Width = 70;
            dataGridView1.Columns[61].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[62].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[63].DefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.Columns[64].DefaultCellStyle.BackColor = Color.AliceBlue;
        }

        private void flowLayoutPanel1_Scroll(object sender, ScrollEventArgs e)
        {
            flowLayoutPanel2.HorizontalScroll.Value = flowLayoutPanel1.HorizontalScroll.Value;
        }

        private void flowLayoutPanel2_Scroll(object sender, ScrollEventArgs e)
        {
            flowLayoutPanel1.HorizontalScroll.Value = flowLayoutPanel2.HorizontalScroll.Value;
        }

        private void resizeControl(Rectangle r, Control c)
        {
            float xRatio = (float)(this.Width) / (float)(originalFormSize.Width);
            float yRatio = (float)(this.Height) / (float)(originalFormSize.Height);

            int newX = (int)(r.Location.X * xRatio);
            int newY = (int)(r.Location.Y * yRatio);

            int newWidth = (int)(r.Width * xRatio);
            int newHeight = (int)(r.Height * yRatio);

            c.Location = new Point(newX, newY);
            c.Size = new Size(newWidth, newHeight);
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            for (int i = 0; i < this.Controls.Count; i++)
            {
                resizeControl(anotherComponentsRectangle[i], anotherComponents[i]);
            }

        }

        private List<List<decimal>> calculate(int old, int current, int consumption, int hourDiff)
        {
            decimal increased_value = old;
            decimal tmp1 = current - old;
            if (hourDiff > 0)
            {
                decimal maxValueCount = tmp1 % hourDiff;
                decimal minValue = Math.Floor(tmp1 / hourDiff);
                decimal maxValue = Math.Ceiling(tmp1 / hourDiff);

                var isAscOrder = false;
                if (consumption != null && consumption <= minValue)
                    isAscOrder = true;

                List<List<decimal>> newValues = new List<List<decimal>>();

                List<decimal> tmp = new List<decimal>();
                tmp.Add(consumption);
                tmp.Add(old);
                newValues.Add(tmp);

                if (!isAscOrder)
                {
                    for (int index = 0; index < maxValueCount; index++)
                    {
                        tmp = new List<decimal>();
                        tmp.Add(maxValue);
                        increased_value += maxValue;
                        tmp.Add(increased_value);
                        newValues.Add(tmp);
                    }
                    for (int index = 0; index < (hourDiff - maxValueCount); index++)
                    {
                        tmp = new List<decimal>();
                        tmp.Add(minValue);
                        increased_value += minValue;
                        tmp.Add(increased_value);
                        newValues.Add(tmp);
                    }
                }
                else
                {
                    for (int index = 0; index < (hourDiff - maxValueCount); index++)
                    {
                        tmp = new List<decimal>();
                        tmp.Add(minValue);
                        increased_value += minValue;
                        tmp.Add(increased_value);
                        newValues.Add(tmp);
                    }
                    for (int index = 0; index < maxValueCount; index++)
                    {
                        tmp = new List<decimal>();
                        tmp.Add(maxValue);
                        increased_value += maxValue;
                        tmp.Add(increased_value);
                        newValues.Add(tmp);
                    }
                }
                return newValues;
            }
            else
            {
                List<List<decimal>> newValues = new List<List<decimal>>();
                List<decimal> tmp = new List<decimal>();
                tmp.Add(consumption);
                tmp.Add(old);
                newValues.Add(tmp);
                return newValues;
            }
        }

        private List<List<decimal>> calculateTransform(int old, int current, int consumption, int hourDiff)
        {
            decimal increased_value = old;
            decimal tmp1 = current - old;
            if (hourDiff > 0)
            {
                decimal maxValueCount = tmp1 % hourDiff;
                decimal minValue = Math.Floor(tmp1 / hourDiff);
                decimal maxValue = Math.Ceiling(tmp1 / hourDiff);

                minValue = minValue / 2;
                maxValue = maxValue / 2;
                var isAscOrder = false;
                if (consumption != 0 && consumption <= minValue)
                    isAscOrder = true;

                List<List<decimal>> newValues = new List<List<decimal>>();

                List<decimal> tmp = new List<decimal>();
                tmp.Add(consumption);
                tmp.Add(consumption);
                tmp.Add(old);
                newValues.Add(tmp);

                if (!isAscOrder)
                {
                    for (int index = 0; index < maxValueCount; index++)
                    {
                        tmp = new List<decimal>();
                        tmp.Add(Math.Ceiling(maxValue));
                        tmp.Add(Math.Ceiling(tmp1 / hourDiff));
                        increased_value += Math.Ceiling(tmp1 / hourDiff);
                        tmp.Add(increased_value);
                        newValues.Add(tmp);
                    }
                    for (int index = 0; index < (hourDiff - maxValueCount); index++)
                    {
                        tmp = new List<decimal>();
                        tmp.Add(Math.Floor(minValue));
                        tmp.Add(Math.Floor(tmp1 / hourDiff));
                        increased_value += Math.Floor(tmp1 / hourDiff);
                        tmp.Add(increased_value);
                        newValues.Add(tmp);
                    }
                }
                else
                {
                    for (int index = 0; index < (hourDiff - maxValueCount); index++)
                    {
                        tmp = new List<decimal>();
                        tmp.Add(Math.Floor(minValue));
                        tmp.Add(Math.Floor(tmp1 / hourDiff));
                        increased_value += Math.Floor(tmp1 / hourDiff);
                        tmp.Add(increased_value);
                        newValues.Add(tmp);
                    }
                    for (int index = 0; index < maxValueCount; index++)
                    {
                        tmp = new List<decimal>();
                        tmp.Add(Math.Ceiling(maxValue));
                        tmp.Add(Math.Ceiling(tmp1 / hourDiff));
                        increased_value += Math.Ceiling(tmp1 / hourDiff);
                        tmp.Add(increased_value);
                        newValues.Add(tmp);
                    }
                }
                return newValues;
            }
            else
            {
                List<List<decimal>> newValues = new List<List<decimal>>();
                List<decimal> tmp = new List<decimal>();
                tmp.Add(consumption);
                tmp.Add(consumption);
                tmp.Add(current);
                newValues.Add(tmp);
                return newValues;
            }

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            var result_msg = MessageBox.Show(" هل انت متأكد من القيام بحساب الاستهلاك من الساعة " + numericUpDown1.Value + " إلى الساعة " + numericUpDown2.Value, "تنبيه", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
            if (result_msg == DialogResult.Yes)
            {
                // name convention is line1old, line1current
                // transformers start from line 12 
                int hourDiff = (int)(numericUpDown2.Value - numericUpDown1.Value);

                if (hourDiff < 0)
                {
                    MessageBox.Show("ساعة البدء يجب أن تكون أكبر أو تساوي ساعة الانتهاء");
                    goto end0;
                }

                if(numericUpDown2.Value >24 || numericUpDown1.Value <0)
                {
                    MessageBox.Show("يجب اختيار قيم صحيحة لساعة البدء والانتهاء");
                    goto end0;
                }

                // normal lines
                for (int col = 0; col < 40; col += 4)
                {
                    int old = 0, current = 0, consumption = 0;
                    TextBox lineold = (TextBox)this.Controls.Find("line" + (col / 4) + 1 + "old", true).First();
                    int.TryParse(lineold.Text, out old);
                    TextBox linecurrent = (TextBox)this.Controls.Find("line" + (col / 4) + 1 + "current", true).First();
                    int.TryParse(linecurrent.Text, out current);
                    TextBox lineconsumption = (TextBox)this.Controls.Find("line" + (col / 4) + 2 + "old", true).First();
                    int.TryParse(lineconsumption.Text, out consumption);
                    List<List<decimal>> result = calculate(old, current, consumption, hourDiff);

                    // line part1

                        if ((int)numericUpDown1.Value == 0)
                        {
                            int counter2 = 1;
                            for (int i = (int)numericUpDown1.Value; i < (int)numericUpDown2.Value; i++)
                            {
                                dataGridView1[col + 1, i].Value = result[counter2][1];
                                dataGridView1[col + 2, i].Value = result[counter2][0];
                                counter2++;
                            }
                        }
                        else
                        {
                            int counter2 = 0;
                            for (int i = (int)numericUpDown1.Value - 1; i < (int)numericUpDown2.Value; i++)
                            {
                                dataGridView1[col + 1, i].Value = result[counter2][1];
                                dataGridView1[col + 2, i].Value = result[counter2][0];
                                counter2++;
                            }
                        }

                    // line part2
                    lineold = (TextBox)this.Controls.Find("line" + (col / 4) + 3 + "old", true).First();
                    int.TryParse(lineold.Text, out old);
                    linecurrent = (TextBox)this.Controls.Find("line" + (col / 4) + 3 + "current", true).First();
                    int.TryParse(linecurrent.Text, out current);
                    lineconsumption = (TextBox)this.Controls.Find("line" + (col / 4) + 4 + "old", true).First();
                    int.TryParse(lineconsumption.Text, out consumption);
                    result = calculate(old, current, consumption, hourDiff);

                        if ((int)numericUpDown1.Value == 0)
                        {
                            int counter4 = 1;
                            for (int i = (int)numericUpDown1.Value; i < (int)numericUpDown2.Value; i++)
                            {
                                dataGridView1[col + 3, i].Value = result[counter4][1];
                                dataGridView1[col + 4, i].Value = result[counter4][0];
                                counter4++;
                            }
                        }
                        else
                        {
                            int counter4 = 0;
                            for (int i = (int)numericUpDown1.Value - 1; i < (int)numericUpDown2.Value; i++)
                            {
                                dataGridView1[col + 3, i].Value = result[counter4][1];
                                dataGridView1[col + 4, i].Value = result[counter4][0];
                                counter4++;
                            }
                        }
                }

                // transformer lines
                for (int col = 40; col < 52; col += 4)
                {
                    int old = 0, current = 0, consumption = 0;
                    TextBox lineold = (TextBox)this.Controls.Find("line" + (col / 4) + 1 + "old", true).First();
                    int.TryParse(lineold.Text, out old);
                    TextBox linecurrent = (TextBox)this.Controls.Find("line" + (col / 4) + 1 + "current", true).First();
                    int.TryParse(linecurrent.Text, out current);
                    TextBox lineconsumption = (TextBox)this.Controls.Find("line" + (col / 4) + 2 + "old", true).First();
                    int.TryParse(lineconsumption.Text, out consumption);
                    List<List<decimal>> resultTransformer = calculateTransform(old, current, consumption, hourDiff);

                    // line part1
                        if ((int)numericUpDown1.Value == 0)
                        {
                            int counter6 = 1;
                            for (int i = (int)numericUpDown1.Value; i < (int)numericUpDown2.Value; i++)
                            {
                                dataGridView1[col + 1, i].Value = resultTransformer[counter6][2];
                                dataGridView1[col + 2, i].Value = resultTransformer[counter6][0];
                                counter6++;
                            }
                        }
                        else
                        {
                            int counter6 = 0;
                            for (int i = (int)numericUpDown1.Value - 1; i < (int)numericUpDown2.Value; i++)
                            {
                                dataGridView1[col + 1, i].Value = resultTransformer[counter6][2];
                                dataGridView1[col + 2, i].Value = resultTransformer[counter6][0];
                                counter6++;
                            }
                        }


                    // line part2
                    lineold = (TextBox)this.Controls.Find("line" + (col / 4) + 3 + "old", true).First();
                    int.TryParse(lineold.Text, out old);
                    linecurrent = (TextBox)this.Controls.Find("line" + (col / 4) + 3 + "current", true).First();
                    int.TryParse(linecurrent.Text, out current);
                    lineconsumption = (TextBox)this.Controls.Find("line" + (col / 4) + 4 + "old", true).First();
                    int.TryParse(lineconsumption.Text, out consumption);
                    resultTransformer = calculateTransform(old, current, consumption, hourDiff);

                        if ((int)numericUpDown1.Value == 0)
                        {
                            int counter9 = 1;
                            for (int i = (int)numericUpDown1.Value; i < (int)numericUpDown2.Value; i++)
                            {
                                dataGridView1[col + 3, i].Value = resultTransformer[counter9][2];
                                dataGridView1[col + 4, i].Value = resultTransformer[counter9][0];
                                counter9++;
                            }
                        }
                        else
                        {
                            int counter9 = 0;
                            for (int i = (int)numericUpDown1.Value - 1; i < (int)numericUpDown2.Value; i++)
                            {
                                dataGridView1[col + 3, i].Value = resultTransformer[counter9][2];
                                dataGridView1[col + 4, i].Value = resultTransformer[counter9][0];
                                counter9++;
                            }
                        }
                }

                // north line
                int north_old = 0, north_current = 0, north_consumption = 0;
                TextBox north_lineold = (TextBox)this.Controls.Find("line" + (52 / 4) + 1 + "old", true).First();
                int.TryParse(north_lineold.Text, out north_old);
                TextBox north_linecurrent = (TextBox)this.Controls.Find("line" + (52 / 4) + 1 + "current", true).First();
                int.TryParse(north_linecurrent.Text, out north_current);
                TextBox north_lineconsumption = (TextBox)this.Controls.Find("line" + (52 / 4) + 2 + "old", true).First();
                int.TryParse(north_lineconsumption.Text, out north_consumption);
                List<List<decimal>> north_result = calculate(north_old, north_current, north_consumption, hourDiff);

                // line part1

                    if ((int)numericUpDown1.Value == 0)
                    {
                        int counter2 = 1;
                        for (int i = (int)numericUpDown1.Value; i < (int)numericUpDown2.Value; i++)
                        {
                            dataGridView1[52 + 1, i].Value = north_result[counter2][1];
                            dataGridView1[52 + 2, i].Value = north_result[counter2][0];
                            counter2++;
                        }
                    }
                    else
                    {
                        int counter2 = 0;
                        for (int i = (int)numericUpDown1.Value - 1; i < (int)numericUpDown2.Value; i++)
                        {
                            dataGridView1[52 + 1, i].Value = north_result[counter2][1];
                            dataGridView1[52 + 2, i].Value = north_result[counter2][0];
                            counter2++;
                        }
                    }

                // line part2
                north_lineold = (TextBox)this.Controls.Find("line" + (52 / 4) + 3 + "old", true).First();
                int.TryParse(north_lineold.Text, out north_old);
                north_linecurrent = (TextBox)this.Controls.Find("line" + (52 / 4) + 3 + "current", true).First();
                int.TryParse(north_linecurrent.Text, out north_current);
                north_lineconsumption = (TextBox)this.Controls.Find("line" + (52 / 4) + 4 + "old", true).First();
                int.TryParse(north_lineconsumption.Text, out north_consumption);
                north_result = calculate(north_old, north_current, north_consumption, hourDiff);


                    if ((int)numericUpDown1.Value == 0)
                    {
                        int counter4 = 1;
                        for (int i = (int)numericUpDown1.Value; i < (int)numericUpDown2.Value; i++)
                        {
                            dataGridView1[52 + 3, i].Value = north_result[counter4][1];
                            dataGridView1[52 + 4, i].Value = north_result[counter4][0];
                            counter4++;
                        }
                    }
                    else
                    {
                        int counter4 = 0;
                        for (int i = (int)numericUpDown1.Value - 1; i < (int)numericUpDown2.Value; i++)
                        {
                            dataGridView1[52 + 3, i].Value = north_result[counter4][1];
                            dataGridView1[52 + 4, i].Value = north_result[counter4][0];
                            counter4++;
                        }
                    }

                // tartous lines
                // arrival1 start at 17
                // esmant start at 29
                if ((int)numericUpDown1.Value == 0)
                {
                    for (int i = (int)numericUpDown1.Value; i < (int)numericUpDown2.Value; i++)
                    {
                        int arrival1, arrival2, arrival3, esmant, tartous, total;
                        int.TryParse(dataGridView1[18, i].Value.ToString(), out arrival1);
                        int.TryParse(dataGridView1[22, i].Value.ToString(), out arrival2);
                        int.TryParse(dataGridView1[26, i].Value.ToString(), out arrival3);
                        int.TryParse(dataGridView1[34, i].Value.ToString(), out esmant);
                        tartous = arrival1 + arrival2 + arrival3 - esmant;

                        // طرطوس
                        dataGridView1[59, i].Value = tartous;
                    }
                }
                else
                {
                    for (int i = (int)numericUpDown1.Value - 1; i < (int)numericUpDown2.Value; i++)
                    {
                        int arrival1, arrival2, arrival3, esmant, tartous, total;
                        int.TryParse(dataGridView1[18, i].Value.ToString(), out arrival1);
                        int.TryParse(dataGridView1[22, i].Value.ToString(), out arrival2);
                        int.TryParse(dataGridView1[26, i].Value.ToString(), out arrival3);
                        int.TryParse(dataGridView1[34, i].Value.ToString(), out esmant);
                        tartous = arrival1 + arrival2 + arrival3 - esmant;

                        // طرطوس
                        dataGridView1[59, i].Value = tartous;
                    }
                }

            }
            end0:;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        void export()
        {
            string date = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");

            string station = "", section = "", morning_notes_1 = "", evening_notes_1 = "", morning_notes_2 = "", evening_notes_2 = "", superviser_engineer = "";
            string qrt = "select * from Sections where working_date='" + date + "'";
            DataSet ds = DBFunctions.fillDataSet(qrt);
            if (ds.Tables[0].Rows.Count > 0)
            {
                station = ds.Tables[0].Rows[0].ItemArray[2].ToString();
                section = ds.Tables[0].Rows[0].ItemArray[3].ToString();
                morning_notes_1 = ds.Tables[0].Rows[0].ItemArray[4].ToString();
                evening_notes_1 = ds.Tables[0].Rows[0].ItemArray[5].ToString();
                morning_notes_2 = ds.Tables[0].Rows[0].ItemArray[6].ToString();
                evening_notes_2 = ds.Tables[0].Rows[0].ItemArray[7].ToString();
                superviser_engineer = ds.Tables[0].Rows[0].ItemArray[9].ToString();
            }

            int startRow = 6;
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            var workbook = new ExcelFile();
            var worksheet = workbook.Worksheets.Add("Sheet1");

            worksheet.ViewOptions.ShowColumnsFromRightToLeft = true;

            // date of file
            var range = worksheet.Cells.GetSubrange("A1:U1");
            range.Merged = true;
            range.Value = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            //range.Style.Font.Color = Color.Red;
            range.Style.Font.Size = 16 * 16;
            range.Style.Font.Weight = ExcelFont.BoldWeight;

            // header
            range = worksheet.Cells.GetSubrange("A2:BM2");
            range.Merged = true;
            range.Value = "مديرية نقل الطاقة";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;
            range.Style.Font.Weight = ExcelFont.BoldWeight;

            range = worksheet.Cells.GetSubrange("A3:BM3");
            range.Merged = true;
            range.Value = "دائرة نقل الطاقة بطرطوس";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;
            range.Style.Font.Weight = ExcelFont.BoldWeight;

            range = worksheet.Cells.GetSubrange("A4:BM4");
            range.Merged = true;
            range.Value = "محطة تحويل:   " + " " + station;
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;
            range.Style.Font.Weight = ExcelFont.BoldWeight;

            range = worksheet.Cells.GetSubrange("A5:BM5");
            range.Merged = true;
            range.Value = "قســـــــم:   " + " " + section;
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;
            range.Style.Font.Weight = ExcelFont.BoldWeight;



            // hours
            range = worksheet.Cells.GetSubrange("A" + startRow + ":A" + (startRow + 1));
            range.Merged = true;
            range.Value = "التوقيت بالساعة";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            range.Style.Rotation = 90;

            // banias1
            range = worksheet.Cells.GetSubrange("B" + startRow + ":E" + startRow);
            range.Merged = true;
            range.Value = "بانياس 1";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // banias2
            range = worksheet.Cells.GetSubrange("F" + startRow + ":I" + startRow);
            range.Merged = true;
            range.Value = "بانياس 2";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // semerian1
            range = worksheet.Cells.GetSubrange("J" + startRow + ":M" + startRow);
            range.Merged = true;
            range.Value = "سمريان 1";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // semerian2
            range = worksheet.Cells.GetSubrange("N" + startRow + ":Q" + startRow);
            range.Merged = true;
            range.Value = "سمريان 2";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // arrival1
            range = worksheet.Cells.GetSubrange("R" + startRow + ":U" + startRow);
            range.Merged = true;
            range.Value = "وصول 1";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // arrival2
            range = worksheet.Cells.GetSubrange("V" + startRow + ":Y" + startRow);
            range.Merged = true;
            range.Value = "وصول 2";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // arrival3
            range = worksheet.Cells.GetSubrange("Z" + startRow + ":AC" + startRow);
            range.Merged = true;
            range.Value = "وصول 3";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // amreet
            range = worksheet.Cells.GetSubrange("AD" + startRow + ":AG" + startRow);
            range.Merged = true;
            range.Value = "عمريت";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // esmant
            range = worksheet.Cells.GetSubrange("AH" + startRow + ":AK" + startRow);
            range.Merged = true;
            range.Value = "اسمنت";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // company
            range = worksheet.Cells.GetSubrange("AL" + startRow + ":AO" + startRow);
            range.Merged = true;
            range.Value = "شركة";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // transformer1
            range = worksheet.Cells.GetSubrange("AP" + startRow + ":AS" + startRow);
            range.Merged = true;
            range.Value = "محولة1";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // transformer2
            range = worksheet.Cells.GetSubrange("AT" + startRow + ":AW" + startRow);
            range.Merged = true;
            range.Value = "محولة2";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // transformer3
            range = worksheet.Cells.GetSubrange("AX" + startRow + ":BA" + startRow);
            range.Merged = true;
            range.Value = "محولة3";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // north
            range = worksheet.Cells.GetSubrange("BB" + startRow + ":BE" + startRow);
            range.Merged = true;
            range.Value = "الشمال";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            for (int i = 1; i <= 64; i += 4)
            {
                worksheet.Cells[startRow, i].Value = "تأشيرة العداد الفعلي";
                worksheet.Cells[startRow, i].Style.WrapText = true;
                worksheet.Cells[startRow, i].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                worksheet.Cells[startRow, i].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                worksheet.Cells[startRow, i + 1].Value = "الاستهلاك ك.و.س";
                worksheet.Cells[startRow, i + 1].Style.WrapText = true;
                worksheet.Cells[startRow, i + 1].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                worksheet.Cells[startRow, i + 1].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                worksheet.Cells[startRow, i + 2].Value = "تأشيرة العداد الردي";
                worksheet.Cells[startRow, i + 2].Style.WrapText = true;
                worksheet.Cells[startRow, i + 2].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                worksheet.Cells[startRow, i + 2].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                worksheet.Cells[startRow, i + 3].Value = "الاستهلاك ك.ف.س";
                worksheet.Cells[startRow, i + 3].Style.WrapText = true;
                worksheet.Cells[startRow, i + 3].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                worksheet.Cells[startRow, i + 3].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            }

            for (int k = 0; k < 24; k++)
            {
                for (int j = 0; j < 64; j++)
                {
                    try
                    {
                        worksheet.Cells[startRow + k + 1, j].Value = int.Parse(dataGridView1.Rows[k].Cells[j].Value.ToString());
                        worksheet.Cells[startRow + k + 1, j].Style.VerticalAlignment = VerticalAlignmentStyle.Center;
                        worksheet.Cells[startRow + k + 1, j].Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
                    }
                    catch (Exception)
                    {
                    }
                }
                worksheet.Cells[startRow + k + 1, 0].Value = k + 1;
            }

            ////// last 2 lines

            /// المجموع
            range = worksheet.Cells.GetSubrange("BI" + (startRow + 26) + ":BJ" + (startRow + 26));
            range.Merged = true;
            range.Value = "المجموع";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            range = worksheet.Cells.GetSubrange("BI" + (startRow + (26 + 1)) + ":BJ" + (startRow + (26 + 1 + 4)));
            range.Merged = true;
            range.Value = dataGridView1[61, 24].Value;
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;


            /// ملاحظات رئيس الوردية الصباحية
            range = worksheet.Cells.GetSubrange("A" + (startRow + 26) + ":Q" + (startRow + 26));
            range.Merged = true;
            range.Value = "ملاحظات رئيس الوردية الصباحية";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            range = worksheet.Cells.GetSubrange("A" + (startRow + 26 + 1) + ":Q" + (startRow + 26 + 5));
            range.Merged = true;
            range.Value = morning_notes_1;
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;

            /// ملاحظات رئيس الوردية المسائية
            range = worksheet.Cells.GetSubrange("R" + (startRow + 26) + ":AG" + (startRow + 26));
            range.Merged = true;
            range.Value = "ملاحظات رئيس الوردية المسائية";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            range = worksheet.Cells.GetSubrange("R" + (startRow + 26 + 1) + ":AG" + (startRow + 26 + 5));
            range.Merged = true;
            range.Value = evening_notes_1;
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;

            /// ملاحظات رئيس الوردية الصباحية
            range = worksheet.Cells.GetSubrange("AH" + (startRow + 26) + ":AW" + (startRow + 26));
            range.Merged = true;
            range.Value = "ملاحظات رئيس الوردية الصباحية";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            range = worksheet.Cells.GetSubrange("AH" + (startRow + 26 + 1) + ":AW" + (startRow + 26 + 5));
            range.Merged = true;
            range.Value = morning_notes_2;
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;

            /// 
            range = worksheet.Cells.GetSubrange("AX" + (startRow + 26) + ":BE" + (startRow + 26));
            range.Merged = true;
            range.Value = "ملاحظات رئيس الوردية المسائية";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            range = worksheet.Cells.GetSubrange("AX" + (startRow + 26 + 1) + ":BH" + (startRow + 26 + 5));
            range.Merged = true;
            range.Value = evening_notes_2;
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Right;

            /// شوهد من قبل المهندس المشرف
            range = worksheet.Cells.GetSubrange("A" + (startRow + 26 + 6) + ":AW" + (startRow + 26 + 7));
            range.Merged = true;
            range.Value = "";
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            range = worksheet.Cells.GetSubrange("AX" + (startRow + 26 + 6) + ":BJ" + (startRow + 26 + 7));
            range.Merged = true;
            range.Value = "شوهد من قبل المهندس المشرف" + " \n" + superviser_engineer;
            range.Style.VerticalAlignment = VerticalAlignmentStyle.Top;
            range.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;

            // Set the style of the merged range using a cell within.
            //worksheet.Cells["C3"].Style.Borders
            //    .SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Red), LineStyle.Double);

            workbook.Save("xlsxs\\" + date + ".xlsx");
            label3.Visible = false;
            button2.Enabled = true;
            linkLabel1.Text = date + ".xlsx";
            linkLabel1.Visible = true;
            MessageBox.Show("تم الانتهاء من تصدير ملف الاكسل");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!DBFunctions.isEmptyDataridView(dataGridView1))
            {
                TextBox.CheckForIllegalCrossThreadCalls = false;
                Button.CheckForIllegalCrossThreadCalls = false;
                LinkLabel.CheckForIllegalCrossThreadCalls = false;
                label3.Visible = true;
                button2.Enabled = false;
                Thread th1 = new Thread(new ThreadStart(export));
                th1.Start();
            }
            else
            {
                MessageBox.Show("لايوجد بيانات لتصديرها");
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == 57 || e.ColumnIndex == 58 || e.ColumnIndex == 59)
                {
                    int banias, semerian, tartous;
                    int.TryParse(dataGridView1[57, e.RowIndex].Value.ToString(), out banias);
                    int.TryParse(dataGridView1[58, e.RowIndex].Value.ToString(), out semerian);
                    int.TryParse(dataGridView1[59, e.RowIndex].Value.ToString(), out tartous);

                    // بانياس
                    dataGridView1[61, e.RowIndex].Value = banias + semerian + tartous;

                    int sum = 0, curr;
                    for (int i = 0; i < 24; i++)
                    {
                        int.TryParse(dataGridView1[61, i].Value.ToString(), out curr);
                        sum += curr;
                    }

                    dataGridView1[61, 24].Value = sum;
                    dataGridView1[62, 24].Value = "المجموع";
                }
            }
            catch (Exception)
            {

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
            string[] lastReadingData = getLastReadingData();
            int except_hour = 0;
            if (lastReadingData.Count() > 0)
            {
                except_hour = int.Parse(lastReadingData[2]);
            }

            var result = MessageBox.Show(" هل انت متأكد من القيام بجلب البيانات من قاعدة البيانات بتاريخ " + date + "للتأكيد اضغط نعم", "تنبيه", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
            if (result == DialogResult.Yes)
            {
                loadDataFromSelectedDateToGridView(except_hour, lastReadingData);
                MessageBox.Show("تم تحميل البيانات بنجاح");
            }

        }

        void loadDataFromSelectedDateToGridView(int except_hour = 25, string[] lastReadingData = null)
        {
            if (except_hour == 24)
                except_hour = 25;
            string date = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
            string qrt = "";
            if (DBFunctions.user_role == 0)
            {
                qrt = "select * from Readings where working_date='" + date + "' order by val(working_hour)";
            }
            else
            {
                qrt = "select * from Readings where working_date='" + date + "' and user_id= " + DBFunctions.user_id + " order by val(working_hour)";
            }
            DataSet ds = DBFunctions.fillDataSet(qrt);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                if (except_hour > 0 && i < except_hour)
                {
                    for (int j = 0; j < 61; j++)
                    {
                        int temp;
                        int.TryParse(ds.Tables[0].Rows[i].ItemArray[j + 2].ToString(), out temp);
                        dataGridView1[j, i].Value = temp;
                    }
                }
            }
            setInitialValuesForStartInputs(lastReadingData);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
            var result = MessageBox.Show("هل انت متأكد من القيام بالحفظ في قاعدة البيانات بتاريخ " + date + " للتاكيد اضغط نعم", "تنبيه", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
            if (result == DialogResult.Yes)
            {
                int hourDiff = (int)(numericUpDown2.Value - numericUpDown1.Value);
                if (hourDiff < 0)
                {
                    MessageBox.Show("ساعة البدء يجب أن تكون أكبر أو تساوي ساعة الانتهاء");
                    goto end0;
                }

                if (numericUpDown2.Value > 24 || numericUpDown1.Value < 0)
                {
                    MessageBox.Show("يجب اختيار قيم صحيحة لساعة البدء والانتهاء");
                    goto end0;
                }

                bool error_happen = false;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    int hour, banias1_send, banias1_send_consumption, banias1_receive, banias1_receive_consumption, banias2_send, banias2_send_consumption,
                        banias2_receive, banias2_receive_consumption, semerian1_send, semerian1_send_consumption, semerian1_receive, semerian1_receive_consumption,
                        semerian2_send, semerian2_send_consumption, semerian2_receive, semerian2_receive_consumption, arrival1_send, arrival1_send_consumption,
                        arrival2_send, arrival2_send_consumption, arrival3_send, arrival3_send_consumption, amreet_send, amreet_send_consumption,
                        esmant_send, esmant_send_consumption, company_send, company_send_consumption, transformer1_send, transformer1_send_consumption,
                        transformer2_send, transformer2_send_consumption, transformer3_send, transformer3_send_consumption, north_send, north_send_consumption,
                        banias, semerian, tartous, total;

                    if (row.Index < 24)
                    {
                        try
                        {
                            int.TryParse(row.Cells[0].Value.ToString(), out hour);
                            int.TryParse(row.Cells[1].Value.ToString(), out banias1_send);
                            int.TryParse(row.Cells[2].Value.ToString(), out banias1_send_consumption);
                            int.TryParse(row.Cells[3].Value.ToString(), out banias1_receive);
                            int.TryParse(row.Cells[4].Value.ToString(), out banias1_receive_consumption);
                            int.TryParse(row.Cells[5].Value.ToString(), out banias2_send);
                            int.TryParse(row.Cells[6].Value.ToString(), out banias2_send_consumption);
                            int.TryParse(row.Cells[7].Value.ToString(), out banias2_receive);
                            int.TryParse(row.Cells[8].Value.ToString(), out banias2_receive_consumption);
                            int.TryParse(row.Cells[9].Value.ToString(), out semerian1_send);
                            int.TryParse(row.Cells[10].Value.ToString(), out semerian1_send_consumption);
                            int.TryParse(row.Cells[11].Value.ToString(), out semerian1_receive);
                            int.TryParse(row.Cells[12].Value.ToString(), out semerian1_receive_consumption);
                            int.TryParse(row.Cells[13].Value.ToString(), out semerian2_send);
                            int.TryParse(row.Cells[14].Value.ToString(), out semerian2_send_consumption);
                            int.TryParse(row.Cells[15].Value.ToString(), out semerian2_receive);
                            int.TryParse(row.Cells[16].Value.ToString(), out semerian2_receive_consumption);
                            int.TryParse(row.Cells[17].Value.ToString(), out arrival1_send);
                            int.TryParse(row.Cells[18].Value.ToString(), out arrival1_send_consumption);
                            int.TryParse(row.Cells[21].Value.ToString(), out arrival2_send);
                            int.TryParse(row.Cells[22].Value.ToString(), out arrival2_send_consumption);
                            int.TryParse(row.Cells[25].Value.ToString(), out arrival3_send);
                            int.TryParse(row.Cells[26].Value.ToString(), out arrival3_send_consumption);
                            int.TryParse(row.Cells[29].Value.ToString(), out amreet_send);
                            int.TryParse(row.Cells[30].Value.ToString(), out amreet_send_consumption);
                            int.TryParse(row.Cells[33].Value.ToString(), out esmant_send);
                            int.TryParse(row.Cells[34].Value.ToString(), out esmant_send_consumption);
                            int.TryParse(row.Cells[37].Value.ToString(), out company_send);
                            int.TryParse(row.Cells[38].Value.ToString(), out company_send_consumption);
                            int.TryParse(row.Cells[41].Value.ToString(), out transformer1_send);
                            int.TryParse(row.Cells[42].Value.ToString(), out transformer1_send_consumption);
                            int.TryParse(row.Cells[45].Value.ToString(), out transformer2_send);
                            int.TryParse(row.Cells[46].Value.ToString(), out transformer2_send_consumption);
                            int.TryParse(row.Cells[49].Value.ToString(), out transformer3_send);
                            int.TryParse(row.Cells[50].Value.ToString(), out transformer3_send_consumption);
                            int.TryParse(row.Cells[53].Value.ToString(), out north_send);
                            int.TryParse(row.Cells[54].Value.ToString(), out north_send_consumption);
                            int.TryParse(row.Cells[57].Value.ToString(), out banias);
                            int.TryParse(row.Cells[58].Value.ToString(), out semerian);
                            int.TryParse(row.Cells[59].Value.ToString(), out tartous);
                            int.TryParse(row.Cells[61].Value.ToString(), out total);

                            string qrt = "";
                            if (DBFunctions.isHourExist(hour.ToString(), date))
                            {
                                qrt = "update Readings set banias1_send='" + banias1_send + "', banias1_send_consumption='" + banias1_send_consumption
                                    + "', banias1_receive='" + banias1_receive + "', banias1_receive_consumption='" + banias1_receive_consumption
                                    + "', banias2_send='" + banias2_send + "', banias2_send_consumption='" + banias2_send_consumption + "', banias2_receive='" + banias2_receive +
                                    "', banias2_receive_consumption='" + banias2_receive_consumption + "', semerian1_send='" + semerian1_send
                                    + "', semerian1_send_consumption='" + semerian1_send_consumption + "', semerian1_receive='" + semerian1_receive +
                                    "', semerian1_receive_consumption='" + semerian1_receive_consumption + "', semerian2_send='" + semerian2_send
                                    + "', semerian2_send_consumption='" + semerian2_send_consumption + "', semerian2_receive='" + semerian2_receive +
                                    "', semerian2_receive_consumption='" + semerian2_receive_consumption + "', arrival1_send='" + arrival1_send
                                    + "', arrival1_send_consumption='" + arrival1_send_consumption + "', arrival2_send='" + arrival2_send +
                                    "', arrival2_send_consumption='" + arrival2_send_consumption + "', arrival3_send='" + arrival3_send
                                    + "', arrival3_send_consumption='" + arrival3_send_consumption + "', amreet_send='" + amreet_send +
                                    "', amreet_send_consumption='" + amreet_send_consumption + "', esmant_send='" + esmant_send +
                                    "', esmant_send_consumption='" + esmant_send_consumption + "', company_send='" + company_send +
                                    "', company_send_consumption='" + company_send_consumption + "', transformer1_send='" + transformer1_send +
                                    "', transformer1_send_consumption='" + transformer1_send_consumption + "', transformer2_send='" + transformer2_send +
                                    "', transformer2_send_consumption='" + transformer2_send_consumption + "', transformer3_send='" + transformer3_send +
                                    "', transformer3_send_consumption='" + transformer3_send_consumption + "',north_send='" + north_send +
                                    "', north_send_consumption='" + north_send_consumption + "', banias='" + banias + "', semerian='" + semerian +
                                    "', tartous='" + tartous + "', total='" + total + "' where working_hour='" + hour + "' and working_date='" + date + "'";
                            }
                            else if (hour < 25)
                            {
                                qrt = "INSERT INTO Readings (working_date, working_hour, banias1_send, banias1_send_consumption, banias1_receive, banias1_receive_consumption, banias2_send, banias2_send_consumption, banias2_receive, banias2_receive_consumption, semerian1_send, semerian1_send_consumption, semerian1_receive, semerian1_receive_consumption, semerian2_send, semerian2_send_consumption, semerian2_receive, semerian2_receive_consumption, arrival1_send, arrival1_send_consumption, arrival2_send, arrival2_send_consumption, arrival3_send, arrival3_send_consumption, amreet_send, amreet_send_consumption, esmant_send, esmant_send_consumption, company_send, company_send_consumption, transformer1_send, transformer1_send_consumption, transformer2_send, transformer2_send_consumption, transformer3_send, transformer3_send_consumption, north_send, north_send_consumption,banias, semerian, tartous, total, user_id) " +
                                    " VALUES ('" + date + "','" + hour + "','" + banias1_send + "','" + banias1_send_consumption + "','" + banias1_receive + "','" + banias1_receive_consumption + "','" + banias2_send + "','" + banias2_send_consumption + "','" + banias2_receive + "','" + banias2_receive_consumption + "','" + semerian1_send + "','" + semerian1_send_consumption + "','" + semerian1_receive + "','" + semerian1_receive_consumption + "','" + semerian2_send + "','" + semerian2_send_consumption + "','" + semerian2_receive + "','" + semerian2_receive_consumption + "','" + arrival1_send + "','" + arrival1_send_consumption + "','" + arrival2_send + "','" + arrival2_send_consumption + "','" + arrival3_send + "','" + arrival3_send_consumption + "','" + amreet_send + "','" + amreet_send_consumption + "','" + esmant_send + "','" + esmant_send_consumption + "','" + company_send + "','" + company_send_consumption + "','" + transformer1_send + "','" + transformer1_send_consumption + "','" + transformer2_send + "','" + transformer2_send_consumption + "','" + transformer3_send + "','" + transformer3_send_consumption + "','" + north_send + "','" + north_send_consumption + "','" + banias + "','" + semerian + "','" + tartous + "','" + total + "'," + DBFunctions.user_id + ")";
                            }

                            if (qrt != "")
                            {
                                DBFunctions.executeCommand(qrt);
                            }
                        }
                        catch (Exception ex)
                        {
                            error_happen = true;
                        }

                    }

                }

                if (error_happen)
                    MessageBox.Show("الرجاء التحقق من المدخلات");
                else
                {
                    string[] lastReadingData = getLastReadingData();
                    setInitialValuesForStartInputs(lastReadingData);
                    MessageBox.Show("تم تخزين البيانات بنجاح");
                }
            end0:;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string fileName = fileName = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
            string file_path = linkLabel1.Text = Application.StartupPath + "/xlsxs/" + fileName + ".xlsx";
            System.Diagnostics.Process.Start(file_path);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            AddUser addUser = new AddUser();
            addUser.ShowDialog();
        }

        private void label4_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            UsersList ul = new UsersList();
            DataSet ds = DBFunctions.fillDataSet("select * from Users where user_role <> 0");

            foreach (DataRow row in ds.Tables[0].Rows)
            {
                string role = row.ItemArray[5].ToString() == "0" ? "مدير" : row.ItemArray[5].ToString() == "1" ? "رئيس وردية" : "مناوب";
                string status = row.ItemArray[3].ToString() == "True" ? "نعم" : "لا";

                string[] str = new string[] { row.ItemArray[1].ToString(), role, row.ItemArray[4].ToString(), status };
                ul.dataGridView1.Rows.Add(str);
                ul.dataGridView1.Rows[ul.dataGridView1.Rows.Count - 1].Cells[4].Value = row.ItemArray[3].ToString() == "True" ? "الغاء تفعيل" : "تفعيل";
            }
            ul.ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string date = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
            var result = MessageBox.Show("هل انت متأكد من القيام بتفريغ محتوى الجدول " + " للتاكيد اضغط نعم", "تنبيه", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
            if (result == DialogResult.Yes)
            {
                // create reading table 
                createReadingTable();

                // fill table with empty rows
                fillTableEmptyRows(readingTable, 25);
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            createReadingTable();
            fillTableEmptyRows(readingTable, 25);
            string[] lastReadingData = getLastReadingData();
            int except_hour = 25;
            if (lastReadingData.Count() > 0)
            {
                except_hour = int.Parse(lastReadingData[2]);
            }
            loadDataFromSelectedDateToGridView(except_hour);
            setInitialValuesForStartInputs(lastReadingData);
        }
    }
}
