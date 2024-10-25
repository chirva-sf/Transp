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
    public partial class RepPlanEkspl : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;

        public RepPlanEkspl()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            DateTime d1 = new DateTime(comboBox1.SelectedIndex + 2010, 1, 1);
            DateTime d2 = new DateTime(comboBox1.SelectedIndex + 2010, 12, 31);
            string FileName = Program.TmpPath + @"\~planeks_" + DateTime.Now.ToString("ddMMHH_mmss") + ".doc";
            File.Copy(Program.ProgPath + @"\Templates\planeks.doc", FileName);
            ClWord ClWord = new ClWord();
            ClWord.Open(FileName);
            ClWord.SetVar("rep_year", comboBox1.Text);
            DataTable dt = ClSQL.SelectSQL("EXEC dbo.repPlanExpl '" + fn.DateToStr(d1) + "','" + fn.DateToStr(d2) + "'");
            for (int i = 1; i < dt.Rows.Count; i++) ClWord.AddRow(2);
            int npp = 0;
            foreach (DataRow dr in dt.Rows)
            {
                npp++;
                ClWord.SetCellValue(2, 1, 2 + npp, npp.ToString());
                ClWord.SetCellValue(2, 2, 2 + npp, dr["marka"].ToString());
                ClWord.SetCellValue(2, 3, 2 + npp, dr["gosnomer"].ToString());
                if (dr["mileage"].ToString() != "")
                {
                    ClWord.SetCellValue(2, 4, 2 + npp, ((int)dr["mileage"] * ((int)dr["end_month"] - (int)dr["beg_month"] + 1)).ToString());
                    for (int i = (int)dr["beg_month"]; i <= (int)dr["end_month"]; i++) ClWord.SetCellValue(2, 4 + i, 2 + npp, dr["mileage"].ToString());
                    double kto = Math.Truncate((double)((int)dr["mileage"] * ((int)dr["end_month"] - (int)dr["beg_month"] + 1)) / (int)dr["mileage_to"]);
                    if ((int)dr["mileage_to"] == 10000 && kto > 0) ClWord.SetCellValue(2, 17, 2 + npp, kto.ToString());
                    if ((int)dr["mileage_to"] == 15000 && kto > 0) ClWord.SetCellValue(2, 18, 2 + npp, kto.ToString());
                }
            }
            ClWord.Complete();
            fn.SetActiveWord();
            button1.Enabled = true;
            Close();
        }

        private void RepPlanEkspl_Load(object sender, EventArgs e)
        {
            int i;
            DateTime d = DateTime.Now;
            comboBox1.Items.Clear();
            for (i = 2010; i <= d.Year; i++)
            {
                comboBox1.Items.Add(i.ToString());
            }
            comboBox1.SelectedIndex = d.Year - 2010;
            if (!File.Exists(Program.ProgPath + @"\Templates\planeks.doc"))
            {
                MessageBox.Show("Отсутствует шаблон planeks.doc !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Enabled = false;
            }
        }
    }
}
