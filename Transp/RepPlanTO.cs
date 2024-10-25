using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace Transp
{
    public partial class RepPlanTO : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;

        public RepPlanTO()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            DataTable dt = ClSQL.SelectSQL("EXEC dbo.chkCarsTO");
            if (dt.Rows.Count > 0)
            {
                ClExcel Excel = new ClExcel();
                string FileName = Program.TmpPath + @"\~chkcarsto_" + DateTime.Now.ToString("ddMMHH_mmss") + ".xls";
                File.Copy(Program.ProgPath + @"\Templates\chkcarsto.xls", FileName);
                Excel.Open(FileName);
                Excel.SelectSheet("Лист1");
                for (int y = 1; y <= dt.Rows.Count; y++)
                {
                    for (int x = 1; x <= dt.Columns.Count; x++)
                    {
                        Excel.SetCellValue((char)(64 + x) + (y + 3).ToString(), dt.Rows[y - 1][x - 1].ToString());
                    }
                }
                Excel.Show();
                fn.SetActiveExcel();
            }
            else
            {
                DateTime d1 = new DateTime(comboBox2.SelectedIndex + 2010, comboBox1.SelectedIndex + 1, 1);
                DateTime d2 = new DateTime(comboBox2.SelectedIndex + 2010, comboBox1.SelectedIndex + 1, DateTime.DaysInMonth(comboBox2.SelectedIndex + 2010, comboBox1.SelectedIndex + 1));
                string FileName = Program.TmpPath + @"\~planto_" + DateTime.Now.ToString("ddMMHH_mmss") + ".doc";
                File.Copy(Program.ProgPath + @"\Templates\planto.doc", FileName);
                ClWord ClWord = new ClWord();
                ClWord.Open(FileName);
                ClWord.SetVar("rep_month", comboBox1.Text);
                ClWord.SetVar("rep_year", comboBox2.Text);
                dt = ClSQL.SelectSQL("EXEC dbo.repPlanTO '" + fn.DateToStr(d1) + "','" + fn.DateToStr(d2) + "'");
                for (int i = 1; i < dt.Rows.Count; i++) ClWord.AddRow(2);
                int npp = 0, to10 = 0, to15 = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    npp++;
                    ClWord.SetCellValue(2, 1, 2 + npp, npp.ToString());
                    ClWord.SetCellValue(2, 2, 2 + npp, dr["marka"].ToString());
                    ClWord.SetCellValue(2, 3, 2 + npp, dr["gosnomer"].ToString());
                    if (dr["to_date"].ToString() != "") ClWord.SetCellValue(2, 3 + Int32.Parse(dr["to_date"].ToString().Substring(0, 2)), 2 + npp, "ТО");
                    if (dr["mileage_to"].ToString() != "")
                    {
                        if ((int)dr["mileage_to"] == 10000) to10++;
                        if ((int)dr["mileage_to"] == 15000) to15++;
                    }
                }
                ClWord.SetVar("to10", to10.ToString());
                ClWord.SetVar("to15", to15.ToString());
                ClWord.Complete();
                fn.SetActiveWord();
            }
            button1.Enabled = true;
            Close();
        }

        private void RepPlanTO_Load(object sender, EventArgs e)
        {
            int i;
            string[] m = { "январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь" };
            DateTime d = DateTime.Now;
            comboBox1.Items.Clear();
            for (i = 1; i <= 12; i++)
            {
                comboBox1.Items.Add(m[i - 1]);
            }
            comboBox1.SelectedIndex = d.Month - 1;
            comboBox2.Items.Clear();
            for (i = 2010; i <= d.Year; i++)
            {
                comboBox2.Items.Add(i.ToString());
            }
            comboBox2.SelectedIndex = d.Year - 2010;

            if (!File.Exists(Program.ProgPath + @"\Templates\planto.doc"))
            {
                MessageBox.Show("Отсутствует шаблон planto.doc !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Enabled = false;
            }
        }
    }
}
