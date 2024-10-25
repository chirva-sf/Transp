using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Transp
{
    public partial class prDocRasxod : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public bool EditMode = false;
        public int doc_id = -1;

        public prDocRasxod()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private bool HasErrors()
        {
            DateTime dt;
            if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните номер приказа !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите транспортное средство !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox2.Text == "" && textBox3.Text == "" && textBox7.Text == "" && textBox8.Text == "")
            {
                MessageBox.Show("Введите расход по норме !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox4.Text != "" && !fn.ChkIntVal(textBox4))
            {
                MessageBox.Show("Не верно введен летний пробег !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox2.Text != "" && !fn.ChkDoubleVal(textBox2))
            {
                MessageBox.Show("Не верно введен летний расход по норме !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox7.Text != "" && !fn.ChkDoubleVal(textBox7))
            {
                MessageBox.Show("Не верно введен летний расход по норме !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox5.Text != "" && !fn.ChkIntVal(textBox5))
            {
                MessageBox.Show("Не верно введен зимний пробег !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox3.Text != "" && !fn.ChkDoubleVal(textBox3))
            {
                MessageBox.Show("Не верно введен зимний расход по норме !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox8.Text != "" && !fn.ChkDoubleVal(textBox8))
            {
                MessageBox.Show("Не верно введен зимний расход по норме !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (maskedTextBox1.Text.Trim() != "." && !DateTime.TryParse(maskedTextBox1.Text + "." + dateTimePicker1.Value.ToString("yyyy"), out dt))
            {
                MessageBox.Show("Не верно введено указана дата начала летнего периода !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (maskedTextBox2.Text.Trim() != "." && !DateTime.TryParse(maskedTextBox2.Text + "." + dateTimePicker1.Value.ToString("yyyy"), out dt))
            {
                MessageBox.Show("Не верно введено указана дата начала зимнего периода !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox4.Text != "" && maskedTextBox1.Text.Trim() == ".")
            {
                MessageBox.Show("Укажите дату начала летнего расхода !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox5.Text != "" && maskedTextBox2.Text.Trim() == ".")
            {
                MessageBox.Show("Укажите дату начала зимнего расхода !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            return false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!HasErrors())
            {
                DateTime dt1 = maskedTextBox1.Text.Trim() == "." ? DateTime.MinValue : DateTime.Parse(maskedTextBox1.Text + "." + dateTimePicker1.Value.ToString("yyyy"));
                DateTime dt2 = maskedTextBox2.Text.Trim() == "." ? DateTime.MinValue : DateTime.Parse(maskedTextBox2.Text + "." + dateTimePicker1.Value.ToString("yyyy"));
                if (dt1 != DateTime.MinValue && dt2 != DateTime.MinValue && dt1 >= dt2)
                {
                    MessageBox.Show("Летняя дата должна быть меньше зимней !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;
                }
                string strSQL = "";
                if (EditMode)
                {
                    strSQL = "update pr_rasxod set ";
                    strSQL += "pr_nom='" + textBox1.Text + "',";
                    strSQL += "pr_date='" + fn.DateToStr(dateTimePicker1.Value) + "',";
                    strSQL += "beg_date='" + fn.DateToStr(dateTimePicker2.Value) + "',";
                    strSQL += "car_id=" + fn.GetTranspSr(comboBox1) + ",";
                    strSQL += "leto_date=" + (dt1 == DateTime.MinValue ? "null" : "'" + fn.DateToStr(dt1) + "'") + ",";
                    strSQL += "leto_mileage=" + fn.NumStr(textBox4) + ",";
                    strSQL += "leto_gorod=" + fn.NumStr(textBox2) + ",";
                    strSQL += "leto_trassa=" + fn.NumStr(textBox7) + ",";
                    strSQL += "zima_date=" + (dt2 == DateTime.MinValue ? "null" : "'" + fn.DateToStr(dt2) + "'") + ",";
                    strSQL += "zima_mileage=" + fn.NumStr(textBox5) + ",";
                    strSQL += "zima_gorod=" + fn.NumStr(textBox3) + ",";
                    strSQL += "zima_trassa=" + fn.NumStr(textBox8) + ",";
                    strSQL += "base_rasxod=" + fn.NumStr(textBox6) + " ";
                    strSQL += "where pr_id=" + doc_id.ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
                else
                {
                    strSQL = "insert into pr_rasxod (pr_nom, pr_date, beg_date, car_id, leto_date, leto_mileage, leto_gorod, leto_trassa, zima_date, zima_mileage, zima_gorod, zima_trassa, base_rasxod) values (";
                    strSQL += "'" + textBox1.Text + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker1.Value.Date) + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker2.Value.Date) + "',";
                    strSQL += fn.GetTranspSr(comboBox1) + ",";
                    strSQL += (dt1 == DateTime.MinValue ? "null" : "'" + fn.DateToStr(dt1) + "'") + ",";
                    strSQL += fn.NumStr(textBox4) + ",";
                    strSQL += fn.NumStr(textBox2) + ",";
                    strSQL += fn.NumStr(textBox7) + ",";
                    strSQL += (dt2 == DateTime.MinValue ? "null" : "'" + fn.DateToStr(dt2) + "'") + ",";
                    strSQL += fn.NumStr(textBox5) + ",";
                    strSQL += fn.NumStr(textBox3) + ",";
                    strSQL += fn.NumStr(textBox8) + ",";
                    strSQL += fn.NumStr(textBox6) + ")";
                    ClSQL.ExecuteSQL(strSQL);
                    doc_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                }
                //if (MessageBox.Show("Выполнить пересчет путевых листов ?", "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) RepairLists();
                Close();
            }
        }

        private void RepairLists()
        {
            // Нужно переделывать reCalcRasxod, чтобы правильно пересчитывал, с учетом деления на Город/Трасса и с учета настройки "Разрешить часы простоя"
            
            if (Program.UserType > 1)
            {
                MessageBox.Show("Недостаточно прав для пересчета путевых листов !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;
            }
            string beg_date = "";
            string car_id = fn.GetTranspSr(comboBox1);
            decimal leto_rasx = fn.StrToDecimal(textBox2);
            decimal zima_rasx = fn.StrToDecimal(textBox3);
            DateTime dt1 = maskedTextBox1.Text.Trim() == "." ? DateTime.MinValue : DateTime.Parse(maskedTextBox1.Text + "." + dateTimePicker1.Value.ToString("yyyy"));
            DateTime dt2 = maskedTextBox2.Text.Trim() == "." ? DateTime.MinValue : DateTime.Parse(maskedTextBox2.Text + "." + dateTimePicker1.Value.ToString("yyyy"));
            beg_date = fn.DateToStr(dateTimePicker2.Value);
            if (dateTimePicker2.Value < new DateTime(2013, 09, 16))
            {
                MessageBox.Show("Для даты меньше 16.09.2013 пересчет пут.листов не предусмотрен !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return;
            } 
            ClSQL.ExecuteSQL("EXEC dbo.reCalcRasxod " + car_id + ",'" + beg_date + "',1");
            
        }

        private void prDocRasxod_Load(object sender, EventArgs e)
        {
            if (EditMode)
            {
                DataRow dr = ClSQL.SelectRow("select * from pr_rasxod where pr_id=" + doc_id);
                textBox1.Text = dr["pr_nom"].ToString();
                dateTimePicker1.Value = DateTime.Parse(dr["pr_date"].ToString());
                dateTimePicker2.Value = DateTime.Parse(dr["beg_date"].ToString());
                fn.UpdateTranspSr(comboBox1, (int)dr["car_id"]);
                maskedTextBox1.Text = dr["leto_date"].ToString() == "" ? "" : dr["leto_date"].ToString().Substring(0, 5);
                textBox4.Text = dr["leto_mileage"].ToString() == "0" ? "" : dr["leto_mileage"].ToString();
                textBox2.Text = dr["leto_gorod"].ToString() == "0" ? "" : dr["leto_gorod"].ToString();
                textBox7.Text = dr["leto_trassa"].ToString() == "0" ? "" : dr["leto_trassa"].ToString();
                maskedTextBox2.Text = dr["zima_date"].ToString() == "" ? "" : dr["zima_date"].ToString().Substring(0, 5);
                textBox5.Text = dr["zima_mileage"].ToString() == "0" ? "" : dr["zima_mileage"].ToString();
                textBox3.Text = dr["zima_gorod"].ToString() == "0" ? "" : dr["zima_gorod"].ToString();
                textBox8.Text = dr["zima_trassa"].ToString() == "0" ? "" : dr["zima_trassa"].ToString();
                textBox6.Text = dr["base_rasxod"].ToString() == "0" ? "" : dr["base_rasxod"].ToString();
            }
            else
            {
                textBox1.Text = (fn.GetMaxNom("pr_nom", "pr_rasxod") + 1).ToString();
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;
                fn.UpdateTranspSr(comboBox1, 0);
            }
        }
    }
}
