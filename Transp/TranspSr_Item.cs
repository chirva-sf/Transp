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
    public partial class TranspSr_Item : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public bool EditMode = false;
        public int item_id = -1;

        public TranspSr_Item()
        {
            InitializeComponent();
        }

        private bool HasErrors()
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните марку, модель !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox2.Text == "")
            {
                MessageBox.Show("Заполните гос.номер !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите вид топлива !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (textBox3.Text == "")
            {
                MessageBox.Show("Заполните межсервисный пробег !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (comboBox2.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите филиал/доп.офис !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (!fn.ChkIntVal(textBox3))
            {
                MessageBox.Show("Не верно введен межсервисный пробег !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (!fn.ChkDoubleVal(textBox9))
            {
                MessageBox.Show("Не верно введена мощность двигателя !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (!fn.ChkDoubleVal(textBox10))
            {
                MessageBox.Show("Не верно введен объем двигателя !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (!fn.ChkIntVal(textBox7))
            {
                MessageBox.Show("Не верно введен год изготовления !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (checkBox1.Checked && dateTimePicker3.Value == null)
            {
                MessageBox.Show("Укажите дату окончания эксплуатации ТС !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            return false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!HasErrors()) 
            {
                string strSQL = "";
                if (EditMode)
                {
                    strSQL = "update cars set ";
                    strSQL += "nom_do='" + fn.GetNomDO(comboBox2) + "',";
                    strSQL += "marka='" + textBox1.Text + "',";
                    strSQL += "ft_id=" + fn.GetFuelType(comboBox1) + ",";
                    strSQL += "beg_date='" + fn.DateToStr(dateTimePicker1.Value) + "',";
                    strSQL += "gai_date='" + fn.DateToStr(dateTimePicker2.Value) + "',";
                    strSQL += "reg_nom='" + textBox16.Text + "',";
                    strSQL += "mileage_to=" + fn.NumStr(textBox3) + ",";
                    strSQL += "gosnomer='" + textBox2.Text + "',";
                    strSQL += "garnomer='" + textBox4.Text + "',";
                    strSQL += "country='" + textBox5.Text + "',";
                    strSQL += "color='" + textBox6.Text + "',";
                    strSQL += "pr_god=" + fn.NumStr(textBox7) + ",";
                    strSQL += "tip_ts='" + textBox8.Text + "',";
                    strSQL += "power=" + fn.NumStr(textBox9) + ",";
                    strSQL += "capacity=" + fn.NumStr(textBox10) + ",";
                    strSQL += "engine='" + textBox12.Text + "',";
                    strSQL += "engine_type='" + textBox11.Text + "',";
                    strSQL += "vin='" + textBox13.Text + "',";
                    strSQL += "kuzov_nom='" + textBox14.Text + "',";
                    strSQL += "chassis_num='" + textBox15.Text + "', ";
                    strSQL += "status=" + (checkBox1.Checked ? 1 : 0) + ", ";
                    strSQL += "end_date=" + (checkBox1.Checked ? "'" + fn.DateToStr(dateTimePicker3.Value.Date) + "'" : "Null") + " ";
                    strSQL += "where car_id=" + item_id.ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
                else
                {
                    strSQL = "insert into cars (nom_do, marka, ft_id, beg_date, gai_date, reg_nom, mileage_to, gosnomer, garnomer, country, color, pr_god, tip_ts,";
                    strSQL += "power, capacity, engine, engine_type, vin, kuzov_nom, chassis_num, status, end_date) values (";
                    strSQL += "'" + fn.GetNomDO(comboBox2) + "',";
                    strSQL += "'" + textBox1.Text + "',";
                    strSQL += fn.GetFuelType(comboBox1) + ",";
                    strSQL += "'" + fn.DateToStr(dateTimePicker1.Value.Date) + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker2.Value.Date) + "',";
                    strSQL += "'" + textBox16.Text + "',";
                    strSQL += fn.NumStr(textBox3) + ",";
                    strSQL += "'" + textBox2.Text + "',";
                    strSQL += "'" + textBox4.Text + "',";
                    strSQL += "'" + textBox5.Text + "',";
                    strSQL += "'" + textBox6.Text + "',";
                    strSQL += fn.NumStr(textBox7) + ",";
                    strSQL += "'" + textBox8.Text + "',";
                    strSQL += fn.NumStr(textBox9) + ",";
                    strSQL += fn.NumStr(textBox10) + ",";
                    strSQL += "'" + textBox12.Text + "',";
                    strSQL += "'" + textBox11.Text + "',";
                    strSQL += "'" + textBox13.Text + "',";
                    strSQL += "'" + textBox14.Text + "',";
                    strSQL += "'" + textBox15.Text + "',";
                    strSQL += (checkBox1.Checked ? 1 : 0) + ",";
                    strSQL += (checkBox1.Checked ? "'" + fn.DateToStr(dateTimePicker3.Value.Date) + "'" : "Null") + ")";
                    ClSQL.ExecuteSQL(strSQL);
                    item_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                }
                Close();
            }
        }

        private void TranspSr_Item_Load(object sender, EventArgs e)
        {
            if (EditMode)
            {
                DataRow dr = ClSQL.SelectRow("select * from cars where car_id=" + item_id);
                fn.UpdateNomDO(comboBox2, dr["nom_do"].ToString());
                textBox1.Text = dr["marka"].ToString();
                fn.UpdateFuelType(comboBox1, (int)dr["ft_id"]);
                dateTimePicker1.Value = DateTime.Parse(dr["beg_date"].ToString());
                dateTimePicker2.Value = DateTime.Parse(dr["gai_date"].ToString());
                textBox16.Text = dr["reg_nom"].ToString();
                textBox3.Text = dr["mileage_to"].ToString();
                textBox2.Text = dr["gosnomer"].ToString();
                textBox4.Text = dr["garnomer"].ToString();
                textBox5.Text = dr["country"].ToString();
                textBox6.Text = dr["color"].ToString();
                textBox7.Text = dr["pr_god"].ToString();
                textBox8.Text = dr["tip_ts"].ToString();
                textBox9.Text = dr["power"].ToString();
                textBox10.Text = dr["capacity"].ToString();
                textBox12.Text = dr["engine"].ToString();
                textBox11.Text = dr["engine_type"].ToString();
                textBox13.Text = dr["vin"].ToString();
                textBox14.Text = dr["kuzov_nom"].ToString();
                textBox15.Text = dr["chassis_num"].ToString();
                checkBox1.Checked = false; dateTimePicker3.Visible = false; label21.Visible = false;
                if (dr["status"].ToString() == "1")
                {
                    checkBox1.Checked = true;
                    if (dr["end_date"].ToString() == "")
                    {
                        dateTimePicker3.Value = DateTime.Now;
                    }
                    else
                    {
                        dateTimePicker3.Value = DateTime.Parse(dr["end_date"].ToString());
                    }
                    dateTimePicker3.Visible = true;
                    label21.Visible = true;
                }
            }
            else
            {
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;
                dateTimePicker3.Value = DateTime.Now;
                dateTimePicker3.Visible = false;
                label21.Visible = false;
                fn.UpdateNomDO(comboBox2, "");
                fn.UpdateFuelType(comboBox1, 0);
                checkBox1.Checked = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            bool b = false;
            if (checkBox1.Checked) b = true;
            dateTimePicker3.Visible = b;
            label21.Visible = b;
        }
    }
}
