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
    public partial class RepDriversTime : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;

        public RepDriversTime()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            DateTime d1 = dateTimePicker1.Value.Date;
            DateTime d2 = dateTimePicker2.Value.Date;
            string FileName, s;
            ClExcel Excel = new ClExcel();
            FileName = Program.TmpPath + @"\~drvtime_" + DateTime.Now.ToString("ddMMHH_mmss") + ".xls";
            File.Copy(Program.ProgPath + @"\Templates\drvtime" + (checkBox1.Checked ? "1" : "2") + ".xls", FileName);
            Excel.Open(FileName);
            Excel.SelectSheet("Лист1");
            Excel.SetCellValue("A1", "Время работы водителей за период с " + fn.DateToStrR(d1) + " по " + fn.DateToStrR(d2));
            DataTable dt = ClSQL.SelectSQL("EXEC dbo.repDriversTime " + (checkBox1.Checked ? "1" : "2") + ",'" + fn.DateToStr(d1) + "','" + fn.DateToStr(d2) + "'");
            progressBar1.Maximum = dt.Rows.Count;
            progressBar1.Value = 0;
            for (int y = 1; y <= dt.Rows.Count; y++)
            {
                progressBar1.Value = y;
                for (int x = 1; x <= dt.Columns.Count; x++)
                {
                    if (checkBox1.Checked) s = dt.Rows[y - 1][x - 1].ToString();
                    else
                    {
                        if (x < 3 || x > 5) s = dt.Rows[y - 1][x - 1].ToString();
                        else if (x == 3) s = fn.DateFromDateTime(dt.Rows[y - 1][x - 1].ToString());
                        else s = fn.TimeFromDateTime(dt.Rows[y - 1][x - 1].ToString());
                    }
                    Excel.SetCellValue((char)(64 + x) + (y + 3).ToString(), s);
                }
            }
            Excel.RunVBA("StartProgramm");
            Excel.Show();
            progressBar1.Value = 0;
            button1.Enabled = true;
            fn.SetActiveExcel();
            Close();
        }

        private void RepDriversTime_Load(object sender, EventArgs e)
        {
            DateTime d = DateTime.Now.AddMonths(-1);
            dateTimePicker1.Value = new DateTime(d.Year, d.Month, 1);
            dateTimePicker2.Value = new DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month));
            checkBox1.Checked = true;
            if (!File.Exists(Program.ProgPath + @"\Templates\drvtime1.xls"))
            {
                MessageBox.Show("Отсутствует шаблон drvtime1.xls !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Enabled = false;
            }
            if (!File.Exists(Program.ProgPath + @"\Templates\drvtime2.xls"))
            {
                MessageBox.Show("Отсутствует шаблон drvtime2.xls !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Enabled = false;
            }
        }
    }
}
