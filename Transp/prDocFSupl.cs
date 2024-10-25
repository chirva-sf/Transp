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
    public partial class prDocFSupl : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public bool EditMode = false;
        public int doc_id = -1;
        private int[] ftarr = new int[30];
        BindingSource BSpl1 = new BindingSource();
        DataTable DTpl1 = new DataTable();
        BindingSource BSpl2 = new BindingSource();
        DataTable DTpl2 = new DataTable();

        public prDocFSupl()
        {
            InitializeComponent();
        }

        private bool HasErrors()
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните номер договора !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox2.Text == "")
            {
                MessageBox.Show("Введите наименование поставщика топлива !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (!fn.ChkDoubleVal(textBox3))
            {
                MessageBox.Show("Неверно введена сумма договора !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            bool flerr = false;
            for (int i = 0; i < DTpl1.Rows.Count; i++)
            {
                DataRow dr = DTpl1.Rows[i];
                if (dr[0].ToString() == "" || dr[1].ToString() == "") { flerr = true; break; }
                if (!fn.ChkDoubleVal(dr[1].ToString())) { flerr = true; break; }
            }
            if (flerr)
            {
                MessageBox.Show("Неверно указанны данные в таблице цен на топливо !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            return false;
        }

        private void SaveTables()
        {
            string strSQL = "delete from pr_fsupl_t where pr_id=" + doc_id.ToString();
            ClSQL.ExecuteSQL(strSQL);
            for (int i = 0; i < DTpl1.Rows.Count; i++)
            {
                DataRow dr = DTpl1.Rows[i];
                if (dr[0].ToString() != "" && dr[1].ToString() != "")
                {
                    DataGridViewComboBoxCell dcc = (DataGridViewComboBoxCell)dataGridView1[0, i];
                    int ftidx = dcc.Items.IndexOf(dcc.Value);
                    strSQL = "insert into pr_fsupl_t values (";
                    strSQL += doc_id.ToString() + ",";
                    strSQL += ftarr[ftidx].ToString() + ",";
                    strSQL += fn.NumStr(dr[1].ToString()) + ")";
                    ClSQL.ExecuteSQL(strSQL);
                }
            }
            strSQL = "update fuel_cards set fs_id=0 where fs_id=" + doc_id.ToString();
            ClSQL.ExecuteSQL(strSQL);
            for (int i = 0; i < DTpl2.Rows.Count; i++)
            {
                DataRow dr = DTpl2.Rows[i];
                strSQL = "update fuel_cards set fs_id=" + doc_id.ToString() + " where fc_id=" + dr["fc_id"];
                ClSQL.ExecuteSQL(strSQL);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!HasErrors())
            {
                string strSQL = "";
                if (EditMode)
                {
                    strSQL = "update pr_fsupl set ";
                    strSQL += "pr_nom='" + textBox1.Text + "',";
                    strSQL += "pr_date='" + fn.DateToStr(dateTimePicker1.Value) + "',";
                    strSQL += "supplier='" + textBox2.Text + "',";
                    strSQL += "beg_date='" + fn.DateToStr(dateTimePicker2.Value) + "',";
                    strSQL += "end_date='" + fn.DateToStr(dateTimePicker3.Value) + "',";
                    strSQL += "max_fuel=0,";
                    strSQL += "max_sum=" + fn.NumStr(textBox3) + " ";
                    strSQL += "where pr_id=" + doc_id.ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
                else
                {
                    strSQL = "insert into pr_fsupl (pr_nom, pr_date, supplier, beg_date, end_date, max_fuel, max_sum) values (";
                    strSQL += "'" + textBox1.Text + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker1.Value.Date) + "',";
                    strSQL += "'" + textBox2.Text + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker2.Value.Date) + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker3.Value.Date) + "',";
                    strSQL += "0,";
                    strSQL += fn.NumStr(textBox3) + ")";
                    ClSQL.ExecuteSQL(strSQL);
                    doc_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                }
                SaveTables();
                Close();
            }
        }

        /* ***** Fuel Price ***** */

        private void FuelPriceTableCreate()
        {
            DataGridViewComboBoxColumn ftps = new DataGridViewComboBoxColumn();
            DataTable dt = ClSQL.SelectSQL("select * from fuel_types order by ft_id");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ftarr[i] = (int)dt.Rows[i]["ft_id"];
                ftps.Items.Add(dt.Rows[i]["ft_name"].ToString());
            }
            ftps.DataPropertyName = "ft_name";
            ftps.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
            ftps.FlatStyle = FlatStyle.Flat;
            DTpl1.Columns.Clear();
            DTpl1.Columns.Add("ft_name");
            DTpl1.Columns.Add("price");
            BSpl1.DataSource = DTpl1;
            dataGridView1.AutoGenerateColumns = false;
            foreach (DataColumn cl in DTpl1.Columns)
            {
                if (cl.ColumnName == "ft_name")
                {
                    dataGridView1.Columns.Add(ftps);
                }
                else
                {
                    DataGridViewColumn column = new DataGridViewTextBoxColumn();
                    column.DataPropertyName = cl.ColumnName;
                    dataGridView1.Columns.Add(column);
                }
            }
            dataGridView1.DataSource = BSpl1;
            dataGridView1.Columns[0].HeaderText = "Вид топлива";
            dataGridView1.Columns[0].DisplayIndex = 0;
            dataGridView1.Columns[0].Width = 80;
            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[1].HeaderText = "Цена за литр";
            dataGridView1.Columns[1].DisplayIndex = 1;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        private void FuelPriceAddNewRow()
        {
            DataRow row = DTpl1.NewRow();
            DTpl1.Rows.Add(row);
            dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0];
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            FuelPriceAddNewRow();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            dataGridView1.BeginEdit(false);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 0)
            {
                DTpl1.Rows[dataGridView1.CurrentCell.RowIndex].Delete();
            }
        }


        /* ***** Fuel Cards ***** */

        private void FuelCardsTableCreate()
        {
            DTpl2.Columns.Clear();
            DTpl2.Columns.Add("fc_id");
            DTpl2.Columns.Add("nom_do");
            DTpl2.Columns.Add("fc_nomer");
            DTpl2.Columns.Add("ft_name");
            DTpl2.Columns.Add("fc_limit");
            BSpl2.DataSource = DTpl2;
            dataGridView2.AutoGenerateColumns = false;
            foreach (DataColumn cl in DTpl2.Columns)
            {
                DataGridViewColumn column = new DataGridViewTextBoxColumn();
                column.DataPropertyName = cl.ColumnName;
                dataGridView2.Columns.Add(column);
            }
            dataGridView2.DataSource = BSpl2;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[1].HeaderText = "ДО";
            dataGridView2.Columns[1].DisplayIndex = 1;
            dataGridView2.Columns[1].Width = 50;
            dataGridView2.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView2.Columns[2].HeaderText = "Номер карты";
            dataGridView2.Columns[2].DisplayIndex = 2;
            dataGridView2.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView2.Columns[3].HeaderText = "Вит топлива";
            dataGridView2.Columns[3].DisplayIndex = 3;
            dataGridView2.Columns[3].Width = 80;
            dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView2.Columns[4].HeaderText = "Лимит";
            dataGridView2.Columns[4].DisplayIndex = 4;
            dataGridView2.Columns[4].Width = 80;
            dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView2.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        private void FuelCardsAddNewRow()
        {
            FuelCards_List frm = new FuelCards_List();
            frm.StartMode = 1;
            frm.ShowDialog();
            if (frm.selid > 0)
            {
                int fc_id = frm.selid;
                object[] rowArray = new object[5];
                DataRow row = ClSQL.SelectRow("select fc.nom_do, fc.fc_nomer, ft.ft_name, fc.fc_limit from fuel_cards fc, fuel_types ft where fc.ft_id=ft.ft_id and fc.fc_id=" + fc_id.ToString());
                rowArray[0] = fc_id.ToString();
                rowArray[1] = row["nom_do"].ToString();
                rowArray[2] = row["fc_nomer"].ToString();
                rowArray[3] = row["ft_name"].ToString();
                rowArray[4] = row["fc_limit"].ToString();
                DTpl2.Rows.Add(rowArray);
                dataGridView2.CurrentCell = dataGridView2.Rows[dataGridView2.RowCount - 1].Cells[1];
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            FuelCardsAddNewRow();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            int crow = dataGridView2.CurrentCell.RowIndex;
            FuelCards_List frm = new FuelCards_List();
            frm.StartMode = 1;
            frm.selid = fn.StrToInt(dataGridView2[0, crow].ToString());
            frm.ShowDialog();
            if (frm.selid > 0)
            {
                int fc_id = frm.selid;
                DataRow row = ClSQL.SelectRow("select fc.nom_do, fc.fc_nomer, ft.ft_name, fc.fc_limit from fuel_cards fc, fuel_types ft where fc.ft_id=ft.ft_id and fc.fc_id=" + fc_id.ToString());
                DTpl2.Rows[crow][0] = fc_id.ToString();
                DTpl2.Rows[crow][1] = row["nom_do"].ToString();
                DTpl2.Rows[crow][2] = row["fc_nomer"].ToString();
                DTpl2.Rows[crow][3] = row["ft_name"].ToString();
                DTpl2.Rows[crow][4] = row["fc_limit"].ToString();
            }
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            if (dataGridView2.RowCount > 0)
            {
                DTpl2.Rows[dataGridView2.CurrentCell.RowIndex].Delete();
            }
        }

        /* ***** Load ***** */

        private void prDocFSupl_Load(object sender, EventArgs e)
        {
            FuelPriceTableCreate();
            FuelCardsTableCreate();
            if (EditMode)
            {
                DataRow dr = ClSQL.SelectRow("select * from pr_fsupl where pr_id=" + doc_id.ToString());
                textBox1.Text = dr["pr_nom"].ToString();
                dateTimePicker1.Value = DateTime.Parse(dr["pr_date"].ToString());
                textBox2.Text = dr["supplier"].ToString();
                dateTimePicker2.Value = DateTime.Parse(dr["beg_date"].ToString());
                dateTimePicker3.Value = DateTime.Parse(dr["end_date"].ToString());
                textBox3.Text = dr["max_sum"].ToString();
                object[] rowarr1 = new object[2];
                DataTable dt = ClSQL.SelectSQL("select ft.ft_name, pt.price from pr_fsupl_t pt, fuel_types ft where pt.ft_id=ft.ft_id and pt.pr_id=" + doc_id.ToString());
                foreach (DataRow r in dt.Rows)
                {
                    DataRow row = DTpl1.NewRow();
                    rowarr1[0] = r["ft_name"];
                    rowarr1[1] = r["price"];
                    row.ItemArray = rowarr1;
                    DTpl1.Rows.Add(row);
                }
                object[] rowarr2 = new object[5];
                dt = ClSQL.SelectSQL("select fc.fc_id, fc.nom_do, fc.fc_nomer, ft.ft_name, fc.fc_limit from fuel_cards fc, fuel_types ft where fc.ft_id=ft.ft_id and fc.fs_id=" + doc_id.ToString());
                foreach (DataRow r in dt.Rows)
                {
                    DataRow row = DTpl2.NewRow();
                    rowarr2[0] = r["fc_id"];
                    rowarr2[1] = r["nom_do"];
                    rowarr2[2] = r["fc_nomer"];
                    rowarr2[3] = r["ft_name"];
                    rowarr2[4] = r["fc_limit"];
                    row.ItemArray = rowarr2;
                    DTpl2.Rows.Add(row);
                }
            }
            else
            {
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

    }
}
