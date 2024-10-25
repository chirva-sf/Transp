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
    public partial class RepVedomost : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;

        public RepVedomost()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            DateTime d1 = dateTimePicker1.Value.Date;
            DateTime d2 = dateTimePicker2.Value.Date;
            ClExcel Excel = new ClExcel();
            string FileName = Program.TmpPath + @"\~vedomost_" + DateTime.Now.ToString("ddMMHH_mmss") + ".xls";
            File.Copy(Program.ProgPath + @"\Templates\vedomost.xls", FileName);
            Excel.Open(FileName);
            Excel.SelectSheet("Данные");
            Excel.SetCellValue("A1", fn.DateToStrR(d1));
            Excel.SetCellValue("B1", fn.DateToStrR(d2));
            Excel.SetCellValue("C1", ClSQL.SelectCell("SELECT COUNT(*) FROM cars WHERE beg_date<='" + fn.DateToStr(d2) + "' and (status=0 or end_date>='" + fn.DateToStr(d1) + "')"));
            Excel.SetCellValue("D1", ClSQL.SelectCell("select name from settings where setkod='sign1'"));
            Excel.SetCellValue("E1", ClSQL.SelectCell("select curvalue from settings where setkod='sign1'"));
            Excel.SetCellValue("F1", ClSQL.SelectCell("select name from settings where setkod='sign2'"));
            Excel.SetCellValue("G1", ClSQL.SelectCell("select curvalue from settings where setkod='sign2'"));
            DataTable dt = ClSQL.SelectSQL("EXEC dbo.repVedomost '" + fn.DateToStr(d1) + "','" + fn.DateToStr(d2) + "'");
            progressBar1.Maximum = dt.Rows.Count;
            progressBar1.Value = 0;
            for (int y = 1; y <= dt.Rows.Count; y++)
            {
                progressBar1.Value = y;
                for (int x = 1; x <= dt.Columns.Count; x++)
                {
                    Excel.SetCellValue((char)(64 + x) + (y + 1).ToString(), dt.Rows[y - 1][x - 1].ToString());
                }
            }
            Excel.RunVBA("StartProgramm");
            Excel.Show();
            progressBar1.Value = 0;
            button1.Enabled = true;
            fn.SetActiveExcel();
            Close();
        }

        private void RepVedomost_Load(object sender, EventArgs e)
        {
            DateTime d = DateTime.Now.AddMonths(-1);
            dateTimePicker1.Value = new DateTime(d.Year, d.Month, 1);
            dateTimePicker2.Value = new DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month));
            if (!File.Exists(Program.ProgPath + @"\Templates\vedomost.xls"))
            {
                MessageBox.Show("Отсутствует шаблон vedomost.xls !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Enabled = false;
            }
        }
    }
}
