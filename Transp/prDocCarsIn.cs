using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Transp
{
    public partial class prDocCarsIn : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        BindingSource BSob = new BindingSource();
        DataTable DTob = new DataTable();
        public bool EditMode = false;
        public int doc_id = -1;

        public prDocCarsIn()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private bool HasErrors()
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните номер приказа !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите ТС !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            bool fl_er1 = false;
            bool fl_er2 = false;
            for (int i = 0; i < DTob.Rows.Count; i++)
            {
                DataRow dr = DTob.Rows[i];
                if (dr[0].ToString() != "" || dr[1].ToString() != "" || dr[2].ToString() != "")
                {
                    if (!fn.ChkIntVal(dr[1].ToString())) fl_er1 = true;
                    if (!fn.ChkDoubleVal(dr[2].ToString())) fl_er2 = true;
                }
            }
            if (fl_er1)
            {
                MessageBox.Show("Неверно заполнен пробег в расходах по месяцам!", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            if (fl_er2)
            {
                MessageBox.Show("Неверно заполнено топливо в расходах по месяцам !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            return false;
        }

        private void SaveOborot()
        {
            int m;
            string monthes = "январь,февраль,март,апрель,май,июнь,июль,август,сентябрь,октябрь,ноябрь,декабрь";
            string[] mon_arr = monthes.Split(',');
            string strSQL = "delete from pr_cars_in_t where pr_id=" + doc_id.ToString();
            ClSQL.ExecuteSQL(strSQL);
            for (int i = 0; i < DTob.Rows.Count; i++)
            {
                DataRow dr = DTob.Rows[i];
                if (dr[0].ToString() != "" || dr[1].ToString() != "" || dr[2].ToString() != "")
                {
                    strSQL = "insert into pr_cars_in_t values (";
                    strSQL += doc_id.ToString() + ",";
                    for (m = 0; m < 12; m++) if (mon_arr[m] == dr[0].ToString()) break;
                    strSQL += "'" + (m + 1).ToString() + "/1/2013',";
                    strSQL += fn.NumStr(dr[1].ToString()) + ",";
                    strSQL += fn.NumStr(dr[2].ToString()) + ")";
                    ClSQL.ExecuteSQL(strSQL);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!HasErrors())
            {
                string strSQL = "";
                if (EditMode)
                {
                    strSQL = "update pr_cars_in set ";
                    strSQL += "pr_nom='" + textBox1.Text + "',";
                    strSQL += "pr_date='" + fn.DateToStr(dateTimePicker1.Value) + "',";
                    strSQL += "car_id=" + fn.GetTranspSr(comboBox1) + ",";
                    strSQL += "beg_date='" + fn.DateToStr(dateTimePicker2.Value) + "',";
                    strSQL += "in_mileage=" + fn.NumStr(textBox3) + ",";
                    strSQL += "in_fuel=" + fn.NumStr(textBox4) + " ";
                    strSQL += "where pr_id=" + doc_id.ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
                else
                {
                    strSQL = "insert into pr_cars_in (pr_nom, pr_date, car_id, beg_date, in_mileage, in_fuel) values (";
                    strSQL += "'" + textBox1.Text + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker1.Value.Date) + "',";
                    strSQL += fn.GetTranspSr(comboBox1) + ",";
                    strSQL += "'" + fn.DateToStr(dateTimePicker2.Value.Date) + "',";
                    strSQL += fn.NumStr(textBox3) + ",";
                    strSQL += fn.NumStr(textBox4) + ")";
                    ClSQL.ExecuteSQL(strSQL);
                    doc_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                }
                SaveOborot();
                Close();
            }
        }

        private void OborotTableCreate()
        {
            DataGridViewComboBoxColumn monthes = new DataGridViewComboBoxColumn();
            monthes.Items.AddRange("январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь");
            monthes.DataPropertyName = "beg_month";
            monthes.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
            monthes.FlatStyle = FlatStyle.Flat;
            DTob.Columns.Clear();
            DTob.Columns.Add("beg_month");
            DTob.Columns.Add("total_mileage");
            DTob.Columns.Add("total_fuel");
            BSob.DataSource = DTob;
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = BSob;
            dataGridView1.Columns.Add(monthes);
            DataGridViewColumn column1 = new DataGridViewTextBoxColumn();
            column1.DataPropertyName = "total_mileage";
            dataGridView1.Columns.Add(column1);
            DataGridViewColumn column2 = new DataGridViewTextBoxColumn();
            column2.DataPropertyName = "total_fuel";
            dataGridView1.Columns.Add(column2);
            dataGridView1.Columns[0].HeaderText = "Месяц";
            dataGridView1.Columns[0].DisplayIndex = 0;
            dataGridView1.Columns[0].Width = 120;
            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.Columns[1].HeaderText = "Пробег";
            dataGridView1.Columns[1].DisplayIndex = 1;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].HeaderText = "Топливо";
            dataGridView1.Columns[2].DisplayIndex = 2;
            dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void prDocCarsIn_Load(object sender, EventArgs e)
        {
            string monthes = "январь,февраль,март,апрель,май,июнь,июль,август,сентябрь,октябрь,ноябрь,декабрь";
            string[] mon_arr = monthes.Split(',');
            OborotTableCreate();
            if (EditMode)
            {
                DataRow dr = ClSQL.SelectRow("select * from pr_cars_in where pr_id=" + doc_id);
                textBox1.Text = dr["pr_nom"].ToString();
                dateTimePicker1.Value = DateTime.Parse(dr["pr_date"].ToString());
                dateTimePicker2.Value = DateTime.Parse(dr["beg_date"].ToString());
                fn.UpdateTranspSr(comboBox1, (int)dr["car_id"]);
                textBox3.Text = dr["in_mileage"].ToString();
                textBox4.Text = dr["in_fuel"].ToString();
                object[] rowArray = new object[3];
                string strSQL = "select * from pr_cars_in_t where pr_id=" + doc_id.ToString() + " order by beg_date";
                DataTable dt = ClSQL.SelectSQL(strSQL);
                if (dt.Rows.Count < 1)
                {
                    OborotAddNewRow();
                }
                else
                {
                    foreach (DataRow r in dt.Rows)
                    {
                        DataRow row = DTob.NewRow();
                        int m = Int32.Parse(r["beg_date"].ToString().Substring(3, 2));
                        rowArray[0] = mon_arr[m - 1];
                        rowArray[1] = r["total_mileage"].ToString() == "0" ? "" : r["total_mileage"].ToString();
                        rowArray[2] = r["total_fuel"].ToString() == "0" ? "" : r["total_fuel"].ToString();
                        row.ItemArray = rowArray;
                        DTob.Rows.Add(row);
                    }
                }
            }
            else
            {
                textBox1.Text = (fn.GetMaxNom("pr_nom", "pr_cars_in") + 1).ToString();
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;
                fn.UpdateTranspSr(comboBox1, 0);
                OborotAddNewRow();
            }
        }

        private void OborotAddNewRow()
        {
            DataRow row = DTob.NewRow();
            DTob.Rows.Add(row);
            dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0];
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            OborotAddNewRow();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            dataGridView1.BeginEdit(false);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 0)
            {
                DTob.Rows[dataGridView1.CurrentCell.RowIndex].Delete();
            }
        }

    }
}
