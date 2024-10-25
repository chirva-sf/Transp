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
    public partial class PutList_pr : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        BindingSource BSpr = new BindingSource();
        DataTable DTpr = new DataTable();
        public int doc_id = -1;
        DateTime doc_date = DateTime.MinValue;
        int doc_car_id = -1;
        int doc_mileage = -1;
        decimal doc_fuel = -1;

        public PutList_pr()
        {
            InitializeComponent();
        }

        private void PutList_pr_Load(object sender, EventArgs e)
        {
            DTpr.Columns.Clear();
            DTpr.Columns.Add("pl_id");
            DTpr.Columns.Add("pl_nom");
            DTpr.Columns.Add("pl_date");
            DTpr.Columns.Add("beg_mileage");
            DTpr.Columns.Add("cor_mileage");
            DTpr.Columns.Add("beg_fuel");
            DTpr.Columns.Add("bcor_fuel");
            DTpr.Columns.Add("end_fuel");
            DTpr.Columns.Add("ecor_fuel");
            BSpr.DataSource = DTpr;
            dataGridView1.DataSource = BSpr;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[0].DisplayIndex = 0;
            dataGridView1.Columns[1].HeaderText = "Номер документа";
            dataGridView1.Columns[1].DisplayIndex = 1;
            dataGridView1.Columns[1].Width = 80;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[1].DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.Columns[2].HeaderText = "Дата документа";
            dataGridView1.Columns[2].DisplayIndex = 2;
            dataGridView1.Columns[2].Width = 80;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.Columns[3].HeaderText = "Нач.пробег в док-те";
            dataGridView1.Columns[3].DisplayIndex = 3;
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.Columns[4].HeaderText = "Должет быть пробег";
            dataGridView1.Columns[4].DisplayIndex = 4;
            dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.Columns[5].HeaderText = "Нач.остаток топлива";
            dataGridView1.Columns[5].DisplayIndex = 5;
            dataGridView1.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[5].DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.Columns[6].HeaderText = "Должет быть остаток";
            dataGridView1.Columns[6].DisplayIndex = 6;
            dataGridView1.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[6].DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.Columns[7].HeaderText = "Кон.остаток топлива";
            dataGridView1.Columns[7].DisplayIndex = 7;
            dataGridView1.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[7].DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.Columns[8].HeaderText = "Должет быть остаток";
            dataGridView1.Columns[8].DisplayIndex = 8;
            dataGridView1.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[8].DefaultCellStyle.SelectionForeColor = Color.Black;
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            /* *************************************** */
            DataRow rr = ClSQL.SelectRow("select pl_date, car_id, end_mileage, end_fuel from put_lists where pl_id=" + doc_id.ToString());
            doc_date = (DateTime)rr["pl_date"];
            doc_car_id = (int)rr["car_id"];
            doc_mileage = (int)rr["end_mileage"];
            doc_fuel = (decimal)(double)rr["end_fuel"];
            DataRow cr = ClSQL.SelectRow("select marka, gosnomer from cars where car_id=" + doc_car_id.ToString());
            textBox1.Text = cr["marka"].ToString() + " " + cr["gosnomer"].ToString();
            /* *************************************** */
            bool fldt = true; string ss = ClSQL.SelectCell("select curvalue from settings where setkod='enable_downtime'");
            if (ss != "Да" && ss != "да" && ss != "ДА") fldt = false;
            bool fl1 = false, fl2 = false, fl3 = false;
            int totkm = 0;
            int tek_mileage = doc_mileage;
            decimal tek_fuel = doc_fuel;
            decimal trasx_fuel;
            int cor_mileage;
            decimal bcor_fuel;
            decimal ecor_fuel;
            int rk = -1;
            string uslov = "where car_id=" + doc_car_id.ToString() + " and pl_date>='" + fn.DateToStr(doc_date) + "' ";
            DataTable dt = ClSQL.SelectSQL("select * from put_lists " + uslov + "order by pl_date, pl_id");
            foreach (DataRow dr in dt.Rows)
            {
                if (DateTime.Parse(dr["pl_date"].ToString()) == doc_date && (int)dr["pl_id"] <= doc_id) continue;
                fl1 = false; fl2 = false; fl3 = false;
                cor_mileage = tek_mileage; bcor_fuel = tek_fuel;
                if ((int)dr["beg_mileage"] != tek_mileage) fl1 = true;
                if ((decimal)(double)dr["beg_fuel"] != tek_fuel) fl2 = true;
                /* Расчет расхода и остатка топлива */
                totkm = (int)dr["end_mileage"] - tek_mileage;
                if (totkm < 0)
                {
                    trasx_fuel = 0;
                    tek_fuel = 0;
                }
                else
                {
                    if ((DateTime)dr["pl_date"] < new DateTime(2013, 09, 16))
                    {
                        trasx_fuel = Decimal.Round((totkm * (decimal)(double)dr["rasx_gorod"]) / 100, 1, MidpointRounding.AwayFromZero);
                        tek_fuel = tek_fuel + Decimal.Round((decimal)(double)dr["fuel_in"], 1, MidpointRounding.AwayFromZero) - trasx_fuel;
                    }
                    else if ((DateTime)dr["pl_date"] < new DateTime(2017, 09, 04))
                    {
                        trasx_fuel = (totkm * (decimal)(double)dr["rasx_gorod"]) / 100;
                        if (fldt) trasx_fuel = trasx_fuel + (decimal)(double)dr["rasx_base"] * (decimal)(double)dr["downtime"];
                        tek_fuel = tek_fuel + (decimal)(double)dr["fuel_in"] - trasx_fuel;
                    }
                    else
                    {
                        DataRow tdr = ClSQL.SelectRow("select * from (select sum(mileage) as mileage_gorod from put_lists_t where pl_id=" + dr["pl_id"] + " and mtype=1) t1, (select sum(mileage) as mileage_trassa from put_lists_t where pl_id=" + dr["pl_id"] + " and mtype=2) t2");
                        decimal mileage_gorod = tdr["mileage_gorod"].ToString() == "" ? 0 : (decimal)(double)tdr["mileage_gorod"];
                        decimal mileage_trassa = tdr["mileage_trassa"].ToString() == "" ? 0 : (decimal)(double)tdr["mileage_trassa"];
                        if (mileage_gorod > 0 || mileage_trassa > 0)
                        {
                            trasx_fuel = ((mileage_gorod * (decimal)(double)dr["rasx_gorod"]) / 100) + ((mileage_trassa * (decimal)(double)dr["rasx_trassa"]) / 100);
                            if (fldt) trasx_fuel = trasx_fuel + (decimal)(double)dr["rasx_base"] * (decimal)(double)dr["downtime"];
                            tek_fuel = tek_fuel + (decimal)(double)dr["fuel_in"] - trasx_fuel;
                        }
                        else
                        {
                            trasx_fuel = (totkm * (decimal)(double)dr["rasx_gorod"]) / 100;
                            if (fldt) trasx_fuel = trasx_fuel + (decimal)(double)dr["rasx_base"] * (decimal)(double)dr["downtime"];
                            tek_fuel = tek_fuel + (decimal)(double)dr["fuel_in"] - trasx_fuel;
                        }
                    }
                }
                tek_mileage = (int)dr["end_mileage"];
                /* ***** ***** ***** */
                ecor_fuel = tek_fuel;
                if ((decimal)(double)dr["end_fuel"] != tek_fuel) fl3 = true;
                if (fl1 || fl2 || fl3)
                {
                    rk++;
                    DataRow row = DTpr.NewRow();
                    row["pl_id"] = dr["pl_id"];
                    row["pl_nom"] = dr["pl_nom"];
                    row["pl_date"] = fn.DateFromDateTime(dr["pl_date"].ToString());
                    row["beg_mileage"] = dr["beg_mileage"];
                    row["beg_fuel"] = dr["beg_fuel"];
                    row["end_fuel"] = dr["end_fuel"];
                    row["cor_mileage"] = cor_mileage;
                    row["bcor_fuel"] = bcor_fuel;
                    row["ecor_fuel"] = ecor_fuel;
                    DTpr.Rows.Add(row);
                    if (fl1)
                    {
                        dataGridView1.Rows[rk].Cells[3].Style.ForeColor = Color.DarkRed;
                        dataGridView1.Rows[rk].Cells[3].Style.SelectionForeColor = Color.DarkRed;
                        dataGridView1.Rows[rk].Cells[3].Style.Font = new Font(dataGridView1.Font, FontStyle.Bold);
                        dataGridView1.Rows[rk].Cells[4].Style.ForeColor = Color.DarkGreen;
                        dataGridView1.Rows[rk].Cells[4].Style.SelectionForeColor = Color.DarkGreen;
                        dataGridView1.Rows[rk].Cells[4].Style.Font = new Font(dataGridView1.Font, FontStyle.Bold);
                    }
                    if (fl2)
                    {
                        dataGridView1.Rows[rk].Cells[5].Style.ForeColor = Color.DarkRed;
                        dataGridView1.Rows[rk].Cells[5].Style.SelectionForeColor = Color.DarkRed;
                        dataGridView1.Rows[rk].Cells[5].Style.Font = new Font(dataGridView1.Font, FontStyle.Bold);
                        dataGridView1.Rows[rk].Cells[6].Style.ForeColor = Color.DarkGreen;
                        dataGridView1.Rows[rk].Cells[6].Style.SelectionForeColor = Color.DarkGreen;
                        dataGridView1.Rows[rk].Cells[6].Style.Font = new Font(dataGridView1.Font, FontStyle.Bold);
                    }
                    if (fl3)
                    {
                        dataGridView1.Rows[rk].Cells[7].Style.ForeColor = Color.DarkRed;
                        dataGridView1.Rows[rk].Cells[7].Style.SelectionForeColor = Color.DarkRed;
                        dataGridView1.Rows[rk].Cells[7].Style.Font = new Font(dataGridView1.Font, FontStyle.Bold);
                        dataGridView1.Rows[rk].Cells[8].Style.ForeColor = Color.DarkGreen;
                        dataGridView1.Rows[rk].Cells[8].Style.SelectionForeColor = Color.DarkGreen;
                        dataGridView1.Rows[rk].Cells[8].Style.Font = new Font(dataGridView1.Font, FontStyle.Bold);
                    }
                }
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            PutList clDoc = new PutList();
            clDoc.EditMode = true;
            clDoc.ReadOnly = true;
            string doc_id = dataGridView1.CurrentRow == null ? dataGridView1[0, 0].Value.ToString() : dataGridView1.CurrentRow.Cells[0].Value.ToString();
            clDoc.doc_id = Int32.Parse(doc_id);
            clDoc.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool fl1 = false, fl2 = false, fl3 = false;
            int totkm;
            int tek_mileage = doc_mileage;
            decimal tek_fuel = doc_fuel;
            decimal trasx_fuel;
            int cor_mileage;
            decimal bcor_fuel;
            decimal ecor_fuel;
            string strSQL = "";
            bool fldt = true; string ss = ClSQL.SelectCell("select curvalue from settings where setkod='enable_downtime'");
            if (ss != "Да" && ss != "да" && ss != "ДА") fldt = false;
            string uslov = "where car_id=" + doc_car_id.ToString() + " and pl_date>='" + fn.DateToStr(doc_date) + "' ";
            DataTable dt = ClSQL.SelectSQL("select * from put_lists " + uslov + "order by pl_date, pl_id");
            foreach (DataRow dr in dt.Rows)
            {
                if (DateTime.Parse(dr["pl_date"].ToString()) == doc_date && (int)dr["pl_id"] <= doc_id) continue;
                fl1 = false; fl2 = false; fl3 = false;
                cor_mileage = tek_mileage; bcor_fuel = tek_fuel;
                if ((int)dr["beg_mileage"] != tek_mileage) fl1 = true;
                if ((decimal)(double)dr["beg_fuel"] != tek_fuel) fl2 = true;
                /* Расчет расхода и остатка топлива */
                totkm = (int)dr["end_mileage"] - tek_mileage;
                if (totkm < 0)
                {
                    trasx_fuel = 0;
                    tek_fuel = 0;
                }
                else
                {
                    if ((DateTime)dr["pl_date"] < new DateTime(2013, 09, 16))
                    {
                        trasx_fuel = Decimal.Round((totkm * (decimal)(double)dr["rasx_gorod"]) / 100, 1, MidpointRounding.AwayFromZero);
                        tek_fuel = tek_fuel + Decimal.Round((decimal)(double)dr["fuel_in"], 1, MidpointRounding.AwayFromZero) - trasx_fuel;
                    }
                    else if ((DateTime)dr["pl_date"] < new DateTime(2017, 09, 04))
                    {
                        trasx_fuel = (totkm * (decimal)(double)dr["rasx_gorod"]) / 100;
                        if (fldt) trasx_fuel = trasx_fuel + (decimal)(double)dr["rasx_base"] * (decimal)(double)dr["downtime"];
                        tek_fuel = tek_fuel + (decimal)(double)dr["fuel_in"] - trasx_fuel;
                    }
                    else
                    {
                        DataRow tdr = ClSQL.SelectRow("select * from (select sum(mileage) as mileage_gorod from put_lists_t where pl_id=" + dr["pl_id"] + " and mtype=1) t1, (select sum(mileage) as mileage_trassa from put_lists_t where pl_id=" + dr["pl_id"] + " and mtype=2) t2");
                        decimal mileage_gorod = tdr["mileage_gorod"].ToString() == "" ? 0 : (decimal)(double)tdr["mileage_gorod"];
                        decimal mileage_trassa = tdr["mileage_trassa"].ToString() == "" ? 0 : (decimal)(double)tdr["mileage_trassa"];
                        if (mileage_gorod > 0 || mileage_trassa > 0)
                        {
                            trasx_fuel = ((mileage_gorod * (decimal)(double)dr["rasx_gorod"]) / 100) + ((mileage_trassa * (decimal)(double)dr["rasx_trassa"]) / 100);
                            if (fldt) trasx_fuel = trasx_fuel + (decimal)(double)dr["rasx_base"] * (decimal)(double)dr["downtime"];
                            tek_fuel = tek_fuel + (decimal)(double)dr["fuel_in"] - trasx_fuel;
                        }
                        else
                        {
                            trasx_fuel = (totkm * (decimal)(double)dr["rasx_gorod"]) / 100;
                            if (fldt) trasx_fuel = trasx_fuel + (decimal)(double)dr["rasx_base"] * (decimal)(double)dr["downtime"];
                            tek_fuel = tek_fuel + (decimal)(double)dr["fuel_in"] - trasx_fuel;
                        }
                    }
                }
                tek_mileage = (int)dr["end_mileage"];
                /* ***** ***** ***** */
                ecor_fuel = tek_fuel;
                if ((decimal)(double)dr["end_fuel"] != tek_fuel) fl3 = true;
                if (fl1 || fl2 || fl3)
                {
                    strSQL = "update put_lists set ";
                    strSQL += "beg_mileage=" + cor_mileage.ToString();
                    strSQL += ", beg_fuel=" + fn.NumStr(bcor_fuel.ToString());
                    strSQL += ", end_fuel=" + fn.NumStr(ecor_fuel.ToString());
                    strSQL += " where pl_id=" + dr["pl_id"].ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
            }
            MessageBox.Show("Перерасчет путевых листов выполнен.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("В путевых листах могут быть неверные данные. Вы уверены, что хотите отменить их пересчет ?", "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) Close();
        }
    }
}
