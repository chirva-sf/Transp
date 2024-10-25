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
    public partial class RepTranspNalog : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;

        public RepTranspNalog()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            DateTime d1 = new DateTime(comboBox2.SelectedIndex + 2010, comboBox1.SelectedIndex * 3 + 1, 1);
            DateTime d2 = new DateTime(comboBox2.SelectedIndex + 2010, comboBox1.SelectedIndex * 3 + 3, DateTime.DaysInMonth(comboBox2.SelectedIndex + 2010, comboBox1.SelectedIndex * 3 + 3));
            ClExcel Excel = new ClExcel();
            string FileName = Program.TmpPath + @"\~transpnal_" + DateTime.Now.ToString("ddMMHH_mmss") + ".xls";
            File.Copy(Program.ProgPath + @"\Templates\transpnal.xls", FileName);
            Excel.Open(FileName);
            Excel.SelectSheet("Лист1");
            Excel.SetCellValue("B2", "Данные для расчета транспорртного налога за " + comboBox1.Text + " " + comboBox2.Text + " (авансовый платеж)");
            DataTable dt = ClSQL.SelectSQL("SELECT * FROM viewAllCars WHERE beg_date<='" + fn.DateToStr(d2) + "' and (status=0 or end_date>='" + fn.DateToStr(d1) + "') ORDER BY nom_do, marka, vin");
            progressBar1.Maximum = dt.Rows.Count;
            progressBar1.Value = 0;
            int yy = 5;
            string tnom_do = "";
            for (int i = 1; i <= dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i - 1];
                progressBar1.Value = i;
                if (dr[1].ToString() != tnom_do)
                {
                    if (i > 2) if (dt.Rows[i - 2][0].ToString() == dt.Rows[i - 3][0].ToString()) yy++;
                    tnom_do = dr[1].ToString();
                    if (tnom_do == "5600")
                    {
                        Excel.SetCellValue("C" + yy.ToString(), "Кемеровский РФ");
                    }
                    else
                    {
                        Excel.SetCellValue("C" + yy.ToString(), ClSQL.SelectCell("SELECT name FROM departments WHERE nom_do='" + dr[1].ToString() + "'"));
                    }
                    yy++;
                }
                Excel.SetCellValue("A" + yy.ToString(), i.ToString());
                Excel.SetCellValue("B" + yy.ToString(), "5104");
                Excel.SetCellValue("C" + yy.ToString(), dr[2].ToString());
                Excel.SetCellValue("D" + yy.ToString(), dr[3].ToString());
                Excel.SetCellValue("E" + yy.ToString(), dr[4].ToString());
                Excel.SetCellValue("F" + yy.ToString(), dr[5].ToString().Substring(1,10));
                Excel.SetCellValue("G" + yy.ToString(), dr[6].ToString().Substring(1, 10));
                Excel.SetCellValue("H" + yy.ToString(), "2");
                Excel.SetCellValue("I" + yy.ToString(), ((comboBox1.SelectedIndex + 1) * 0.25).ToString());
                Excel.SetCellValue("J" + yy.ToString(), dr[7].ToString());
                Excel.SetCellValue("K" + yy.ToString(), "5,5");
                yy++;
            }
            Excel.RunVBA("StartProgramm");
            Excel.Show();
            progressBar1.Value = 0;
            button1.Enabled = true;
            fn.SetActiveExcel();
            Close();
        }

        private void RepTranspNalog_Load(object sender, EventArgs e)
        {
            DateTime d = DateTime.Now.AddDays(3);
            int kv = (int)Math.Floor((decimal)((d.Month - 1) / 3));
            int y = d.Year;
            if (kv < 1) { y--; kv = 4; }
            comboBox1.SelectedIndex = kv - 1;
            comboBox2.Items.Clear();
            for (int i = 2010; i <= d.Year; i++)
            {
                comboBox2.Items.Add(i.ToString());
            }
            comboBox2.SelectedIndex = y - 2010;
            if (!File.Exists(Program.ProgPath + @"\Templates\transpnal.xls"))
            {
                MessageBox.Show("Отсутствует шаблон transpnal.xls !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Enabled = false;
            }
        }
    }
}
