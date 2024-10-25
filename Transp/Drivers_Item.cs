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
    public partial class Drivers_Item : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public bool EditMode = false;
        public int item_id = -1;

        public Drivers_Item()
        {
            InitializeComponent();
        }

        private bool HasErrors()
        {
            if (textBox2.Text == "")
            {
                MessageBox.Show("Заполните ФИО !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите подразделение !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (comboBox2.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите филиал/доп.офис !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            return false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!HasErrors())
            {
                string strSQL = "";
                if (EditMode)
                {
                    strSQL = "update drivers set ";
                    strSQL += "nom_do='" + fn.GetNomDO(comboBox2) + "',";
                    strSQL += "tab_no='" + textBox1.Text + "',";
                    strSQL += "dep_id=" + fn.GetDepartment(comboBox1) + ",";
                    strSQL += "fio='" + textBox2.Text + "',";
                    strSQL += "udostov='" + textBox3.Text + "', ";
                    strSQL += "klass='" + textBox4.Text + "', ";
                    strSQL += "status=" + (checkBox1.Checked ? 1 : 0) + ", ";
                    strSQL += "end_date=" + (checkBox1.Checked ? "'" + fn.DateToStr(dateTimePicker1.Value.Date) + "'" : "Null") + " ";
                    strSQL += "where drv_id=" + item_id.ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
                else
                {
                    strSQL = "insert into drivers (nom_do,tab_no,dep_id,fio,udostov,klass,status,end_date) values (";
                    strSQL += fn.GetNomDO(comboBox2) + ",";
                    strSQL += "'" + textBox1.Text + "',";
                    strSQL += fn.GetDepartment(comboBox1) + ",";
                    strSQL += "'" + textBox2.Text + "',";
                    strSQL += "'" + textBox3.Text + "',";
                    strSQL += "'" + textBox4.Text + "',";
                    strSQL += (checkBox1.Checked ? 1 : 0) + ",";
                    strSQL += (checkBox1.Checked ? "'" + fn.DateToStr(dateTimePicker1.Value.Date) + "'" : "Null") + ")";
                    ClSQL.ExecuteSQL(strSQL);
                    item_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                }
                Close();
            }
        }

        private void Drivers_Item_Load(object sender, EventArgs e)
        {
            if (EditMode)
            {
                DataRow dr = ClSQL.SelectRow("select * from drivers where drv_id=" + item_id);
                textBox1.Text = dr["tab_no"].ToString();
                textBox2.Text = dr["fio"].ToString();
                textBox3.Text = dr["udostov"].ToString();
                textBox4.Text = dr["klass"].ToString();
                fn.UpdateDepartments(comboBox1, (int)dr["dep_id"]);
                fn.UpdateNomDO(comboBox2, dr["nom_do"].ToString());
                if (dr["status"].ToString() == "1")
                {
                    checkBox1.Checked = true;
                    if (dr["end_date"].ToString() == "")
                    {
                        dateTimePicker1.Value = DateTime.Now;
                    }
                    else
                    {
                        dateTimePicker1.Value = DateTime.Parse(dr["end_date"].ToString());
                    }
                    dateTimePicker1.Visible = true;
                }
            }
            else
            {
                fn.UpdateDepartments(comboBox1, 0);
                fn.UpdateNomDO(comboBox2, "");
                dateTimePicker1.Visible = false;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            bool b = false;
            if (checkBox1.Checked) b = true;
            dateTimePicker1.Visible = b;
        }
    }
}
