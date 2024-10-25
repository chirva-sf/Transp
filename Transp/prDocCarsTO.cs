using System;
using System.Collections.Generic;
using System.Globalization;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Transp
{
    public partial class prDocCarsTO : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public bool EditMode = false;
        public int doc_id = -1;
        BindingSource BSpl = new BindingSource();
        DataTable DTpl = new DataTable();

        public prDocCarsTO()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void UpdateHeader()
        {
            DataRow dr = ClSQL.SelectRow("select nom_do, gosnomer, garnomer from cars where car_id=" + fn.GetTranspSr(comboBox1));
            textBox2.Text = dr["gosnomer"].ToString();
            textBox3.Text = dr["garnomer"].ToString();
            textBox4.Text = dr["nom_do"].ToString();
        }

        private void TableCreate()
        {
            DataGridViewComboBoxColumn wtps = new DataGridViewComboBoxColumn();
            wtps.Items.AddRange("ТО", "ремонт");
            wtps.DataPropertyName = "work_type";
            wtps.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
            wtps.FlatStyle = FlatStyle.Flat;
            DTpl.Columns.Clear();
            DTpl.Columns.Add("rem_date");
            DTpl.Columns.Add("mileage");
            DTpl.Columns.Add("work_type");
            DTpl.Columns.Add("work_done");
            DTpl.Columns.Add("order_num");
            DTpl.Columns.Add("organization");
            DTpl.Columns.Add("sum_to");
            DTpl.Columns.Add("sum_tr");
            DTpl.Columns.Add("dop_to");
            DTpl.Columns.Add("dop_tr");
            BSpl.DataSource = DTpl;
            dataGridView1.AutoGenerateColumns = false;
            foreach (DataColumn cl in DTpl.Columns) {
                if (cl.ColumnName == "rem_date") {
                    DataGridViewColumn column = new CalendarColumn();
                    column.DataPropertyName = "rem_date";
                    dataGridView1.Columns.Add(column);
                }
                else if (cl.ColumnName == "work_type")
                {
                    dataGridView1.Columns.Add(wtps);
                }
                else
                {
                    DataGridViewColumn column = new DataGridViewTextBoxColumn();
                    column.DataPropertyName = cl.ColumnName;
                    dataGridView1.Columns.Add(column);
                }
            }
            dataGridView1.DataSource = BSpl;
            dataGridView1.Columns[0].HeaderText = "Дата ТО/ТР";
            dataGridView1.Columns[0].DisplayIndex = 0;
            dataGridView1.Columns[0].Width = 80;
            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[1].HeaderText = "Спидометр";
            dataGridView1.Columns[1].DisplayIndex = 1;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[2].HeaderText = "Вид работ";
            dataGridView1.Columns[2].DisplayIndex = 2;
            dataGridView1.Columns[2].Width = 75;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[3].HeaderText = "Описание работ";
            dataGridView1.Columns[3].DisplayIndex = 3;
            dataGridView1.Columns[3].Width = 205;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[4].HeaderText = "Заказ-наряд";
            dataGridView1.Columns[4].DisplayIndex = 4;
            dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[5].HeaderText = "Организация";
            dataGridView1.Columns[5].DisplayIndex = 5;
            dataGridView1.Columns[5].Width = 205;
            dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[6].HeaderText = "Стоимость ТО";
            dataGridView1.Columns[6].DisplayIndex = 6;
            dataGridView1.Columns[6].Width = 75;
            dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[7].HeaderText = "Стоимость ТР";
            dataGridView1.Columns[7].DisplayIndex = 7;
            dataGridView1.Columns[7].Width = 75;
            dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[8].HeaderText = "Запас.части ТО";
            dataGridView1.Columns[8].DisplayIndex = 8;
            dataGridView1.Columns[8].Width = 75;
            dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[8].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[9].HeaderText = "Запас.части ТР";
            dataGridView1.Columns[9].DisplayIndex = 9;
            dataGridView1.Columns[9].Width = 75;
            dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridView1.Columns[9].SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        private void prDocCarsTO_Load(object sender, EventArgs e)
        {
            TableCreate();
            if (EditMode)
            {
                object[] rowArray = new object[10];
                DataRow dr = ClSQL.SelectRow("select * from cars_rem where rem_id=" + doc_id.ToString());
                textBox1.Text = dr["rem_nom"].ToString();
                fn.UpdateTranspSr(comboBox1, (int)dr["car_id"]); UpdateHeader();
                textBox4.Text = dr["nom_do"].ToString();
                DataTable dt = ClSQL.SelectSQL("select rem_date,mileage,work_type,work_done,order_num,organization,sum_to,sum_tr,dop_to,dop_tr from cars_rem_t where rem_id=" + doc_id.ToString() + " order by npp");
                foreach (DataRow r in dt.Rows)
                {
                    for (int i = 0; i <= 9; i++)
                    {
                        string s = r[i].ToString();
                        if (i == 0)
                        {
                            s = fn.DateToStrR((DateTime)r[0]);
                        }
                        if (i == 2)
                        {
                            s = "";
                            if (r[2].ToString() == "1") s = "ТО";
                            if (r[2].ToString() == "2") s = "ремонт";
                        }
                        if (i > 5)
                        {
                            int p = s.IndexOf(",");
                            if (p == -1) p = s.IndexOf(".");
                            if (p == -1) s += ",00"; else if (s.Length - p < 3) s += "0";
                        }
                        rowArray[i] = s;

                    }
                    DataRow row = DTpl.NewRow();
                    row.ItemArray = rowArray;
                    DTpl.Rows.Add(row);
                }
            }
            else
            {
                textBox1.Text = (fn.GetMaxNom("rem_nom", "cars_rem") + 1).ToString();
                fn.UpdateTranspSr(comboBox1, 0);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateHeader();
        }

        private void AddNewRow()
        {
            DataRow row = DTpl.NewRow();
            DTpl.Rows.Add(row);
            dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0];
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            AddNewRow();
            dataGridView1.Focus();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            dataGridView1.BeginEdit(false);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 0)
            {
                DTpl.Rows[dataGridView1.CurrentCell.RowIndex].Delete();
            }
        }

        private bool HasErrors()
        {
            if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите автомобиль !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); return true;
            }
            else if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните номер документа !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); return true;
            }
            bool fl1 = false, fl2 = false, fl3 = false, fl4 = false, fl5 = false, fl6 = false, fl7 = false;
            for (int i = 0; i < DTpl.Rows.Count; i++)
            {
                DataRow dr = DTpl.Rows[i];
                if (!fn.ChkIntVal(dr["mileage"].ToString()) || dr["mileage"].ToString()=="") fl1 = true;
                if (!fn.ChkDoubleVal(dr["sum_to"].ToString())) fl2 = true;
                if (!fn.ChkDoubleVal(dr["sum_tr"].ToString())) fl3 = true;
                if (!fn.ChkDoubleVal(dr["dop_to"].ToString())) fl4 = true;
                if (!fn.ChkDoubleVal(dr["dop_tr"].ToString())) fl5 = true;
                if (dr["work_type"].ToString() == "") fl6 = true;
                if (dr["rem_date"].ToString() == "") fl7 = true;
            }
            string s = "";
            if (fl7) s += (s == "" ? "" : Environment.NewLine) + "Укажите дату ТО/Ремонта !";
            if (fl6) s += (s == "" ? "" : Environment.NewLine) + "Укажите вид работ !";
            if (fl1) s += (s == "" ? "" : Environment.NewLine) + "Не верно заполнен пробег !";
            if (fl2) s += (s == "" ? "" : Environment.NewLine) + "Не верно заполнена Стоимость ТО !";
            if (fl3) s += (s == "" ? "" : Environment.NewLine) + "Не верно заполнена Стоимость ТР !";
            if (fl4) s += (s == "" ? "" : Environment.NewLine) + "Не верно заполнена Запас.части ТО !";
            if (fl5) s += (s == "" ? "" : Environment.NewLine) + "Не верно заполнена Запас.части ТР !";
            if (s != "")
            {
                MessageBox.Show(s, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); return true;
            }
            return false;
        }

        private void SaveCarsTO()
        {
            int npp = 0;
            string strSQL = "";
            if (EditMode)
            {
                strSQL = "update cars_rem set ";
                strSQL += "rem_nom='" + textBox1.Text + "',";
                strSQL += "nom_do='" + textBox4.Text + "',";
                strSQL += "car_id=" + fn.GetTranspSr(comboBox1) + " ";
                strSQL += "where rem_id=" + doc_id.ToString();
                ClSQL.ExecuteSQL(strSQL);
            }
            else
            {
                strSQL = "insert into cars_rem (rem_nom,nom_do,car_id) values (";
                strSQL += "'" + textBox1.Text + "',";
                strSQL += "'" + textBox4.Text + "',";
                strSQL += fn.GetTranspSr(comboBox1) + ")";
                ClSQL.ExecuteSQL(strSQL);
                doc_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
            }
            strSQL = "delete from cars_rem_t where rem_id=" + doc_id.ToString();
            ClSQL.ExecuteSQL(strSQL);
            for (int i = 0; i < DTpl.Rows.Count; i++)
            {
                DataRow dr = DTpl.Rows[i];
                if (dr[0].ToString() != "" || dr[1].ToString() != "" || dr[2].ToString() != "" || dr[3].ToString() != "" ||
                    dr[4].ToString() != "" || dr[5].ToString() != "" || dr[6].ToString() != "" || dr[7].ToString() != "" || dr[8].ToString() != "")
                {
                    npp++;
                    strSQL = "insert into cars_rem_t values (";
                    strSQL += doc_id.ToString() + ",";
                    strSQL += npp.ToString() + ",";
                    strSQL += "'" + fn.StrDateToSQL(dr["rem_date"].ToString()) + "',";
                    strSQL += dr["mileage"].ToString() + ",";
                    strSQL += (dr["work_type"].ToString()=="ТО"?"1":"2") + ",";
                    strSQL += "'" + dr["work_done"].ToString() + "',";
                    strSQL += "'" + dr["order_num"].ToString() + "',";
                    strSQL += "'" + dr["organization"].ToString() + "',";
                    strSQL += fn.NumStr(dr["sum_to"].ToString()) + ",";
                    strSQL += fn.NumStr(dr["sum_tr"].ToString()) + ",";
                    strSQL += fn.NumStr(dr["dop_to"].ToString()) + ",";
                    strSQL += fn.NumStr(dr["dop_tr"].ToString()) + ")";
                    ClSQL.ExecuteSQL(strSQL);
                }
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!HasErrors())
            {
                SaveCarsTO();
                Close();
            }
        }
    }
}
