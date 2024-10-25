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
    public partial class RepFuelCards : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;

        public RepFuelCards()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            DateTime d1 = dateTimePicker1.Value.Date;
            DateTime d2 = dateTimePicker2.Value.Date;
            ClExcel Excel = new ClExcel();
            string FileName = Program.TmpPath + @"\~fuelcards_" + DateTime.Now.ToString("ddMMHH_mmss") + ".xls";
            File.Copy(Program.ProgPath + @"\Templates\fuelcards.xls", FileName);
            Excel.Open(FileName);
            Excel.SelectSheet("Лист1");
            Excel.SetCellValue("A1", "Расход по топливным картам за период с " + fn.DateToStrR(d1) + " по " + fn.DateToStrR(d2));
            DataTable dt = ClSQL.SelectSQL("EXEC dbo.repFuelCards '" + fn.DateToStr(d1) + "','" + fn.DateToStr(d2) + "'");
            progressBar1.Maximum = dt.Rows.Count;
            progressBar1.Value = 0;
            for (int y = 1; y <= dt.Rows.Count; y++)
            {
                int sm = 0; progressBar1.Value = y; 
                for (int x = 1; x <= dt.Columns.Count; x++)
                {
                    if (x > 4) sm++;
                    Excel.SetCellValue((char)(64 + x + sm) + (y + 4).ToString(), dt.Rows[y - 1][x - 1].ToString());
                }
            }
            Excel.RunVBA("StartProgramm");
            Excel.Show();
            progressBar1.Value = 0;
            button1.Enabled = true;
            fn.SetActiveExcel();
            Close();
        }

        private void RepFuelCards_Load(object sender, EventArgs e)
        {
            DateTime d = DateTime.Now.AddMonths(-1);
            dateTimePicker1.Value = new DateTime(d.Year, d.Month, 1);
            dateTimePicker2.Value = new DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month));
            if (!File.Exists(Program.ProgPath + @"\Templates\fuelcards.xls"))
            {
                MessageBox.Show("Отсутствует шаблон fuelcards.xls !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Enabled = false;
            }
        }
    }
}
