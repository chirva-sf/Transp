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
    public partial class RepPasportTS : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;

        public RepPasportTS()
        {
            InitializeComponent();
        }

        private void RepPasportTS_Load(object sender, EventArgs e)
        {
            DateTime d = DateTime.Now.AddMonths(-1);
            dateTimePicker1.Value = new DateTime(d.Year, d.Month, 1);
            dateTimePicker2.Value = new DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month));
            if (!File.Exists(Program.ProgPath + @"\Templates\pasport.doc"))
            {
                MessageBox.Show("Отсутствует шаблон pasport.doc !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.Enabled = false;
            }
            fn.UpdateTranspSr(comboBox1, 0);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите а/м !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            button1.Enabled = false;
            DateTime d1 = dateTimePicker1.Value.Date;
            DateTime d2 = dateTimePicker2.Value.Date;
            string car_id = fn.GetTranspSr(comboBox1);
            string FileName = Program.TmpPath + @"\~pasport_" + DateTime.Now.ToString("ddMMHH_mmss") + ".doc";
            File.Copy(Program.ProgPath + @"\Templates\pasport.doc", FileName);
            ClWord ClWord = new ClWord();
            ClWord.Open(FileName);
            DataRow car = ClSQL.SelectRow("select * from cars where car_id=" + car_id);
            DataRow car_in = ClSQL.SelectRow("select * from pr_cars_in where car_id=" + car_id);
            string drv_id = ClSQL.SelectCell("select drv_id from pr_drvcar where car_id=" + car_id);
            ClWord.SetVar("МаркаМодель", car["marka"].ToString());
            ClWord.SetVar("Госномер", car["gosnomer"].ToString());
            ClWord.SetVar("ДатаЭкспл", car["beg_date"].ToString().Substring(0,10));
            ClWord.SetVar("НачПробег", car_in["in_mileage"].ToString());
            ClWord.SetVar("Гарномер", car["garnomer"].ToString());
            ClWord.SetVar("Страна", car["country"].ToString());
            ClWord.SetVar("Цвет", car["color"].ToString());
            ClWord.SetVar("ГодИзготовл", car["pr_god"].ToString());
            ClWord.SetVar("ТипТС", car["tip_ts"].ToString());
            ClWord.SetVar("Двигатель", car["engine"].ToString());
            ClWord.SetVar("МощнЛс", car["power"].ToString());
            ClWord.SetVar("МощнКВт", ((double)car["power"] * 0.7355).ToString());
            ClWord.SetVar("ОбъемДвиг", car["capacity"].ToString());
            ClWord.SetVar("ТипДвиг", car["engine_type"].ToString());
            ClWord.SetVar("Топливо", ClSQL.SelectCell("select ft_name from fuel_types where ft_id=" + car["ft_id"].ToString()));
            ClWord.SetVar("VIN", car["vin"].ToString());
            ClWord.SetVar("Кузов", car["kuzov_nom"].ToString());
            ClWord.SetVar("Шасси", car["chassis_num"].ToString());
            ClWord.SetVar("Закреплен", ClSQL.SelectCell("select fio from drivers where drv_id=" + drv_id));
            ClWord.Complete();
            fn.SetActiveWord();
            Close();
        }
    }
}
