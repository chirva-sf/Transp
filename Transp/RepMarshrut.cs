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
    public partial class RepMarshrut : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;

        public RepMarshrut()
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
            File.Copy(Program.ProgPath + @"\Templates\drvmarshrut.xls", FileName);
            Excel.Open(FileName);
            Excel.SelectSheet("Лист1");
            Excel.SetCellValue("A1", "Маршруты водителей за период с " + fn.DateToStrR(d1) + " по " + fn.DateToStrR(d2));
            string sql = "select pl.nom_do, c.marka, c.gosnomer, d.fio, pl.pl_date, pt.place_out, pt.place_in, pt.mileage ";
            sql += "from put_lists pl, put_lists_t pt, cars c, drivers d ";
            sql += "where pl.pl_id = pt.pl_id and pl.car_id = c.car_id and pl.drv_id = d.drv_id and pl.pl_date >= '" + fn.DateToStr(d1) + "' and pl.pl_date <= '" + fn.DateToStr(d2) + "' ";
            sql += "order by pl.nom_do, pl.pl_date, c.gosnomer, d.fio, pt.npp";
            DataTable dt = ClSQL.SelectSQL(sql);
            progressBar1.Maximum = dt.Rows.Count;
            progressBar1.Value = 0;
            for (int y = 1; y <= dt.Rows.Count; y++)
            {
                progressBar1.Value = y;
                for (int x = 1; x <= dt.Columns.Count; x++)
                {
                    if (x == 5) s = fn.DateFromDateTime(dt.Rows[y - 1][x - 1].ToString());
                    else s = dt.Rows[y - 1][x - 1].ToString();
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

        private void RepMarshrut_Load(object sender, EventArgs e)
        {
            DateTime d = DateTime.Now.AddMonths(-1);
            dateTimePicker1.Value = new DateTime(d.Year, d.Month, 1);
            dateTimePicker2.Value = new DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month));
            if (!File.Exists(Program.ProgPath + @"\Templates\drvmarshrut.xls"))
            {
                MessageBox.Show("Отсутствует шаблон drvmarshrut.xls !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Enabled = false;
            }
        }
    }
}
