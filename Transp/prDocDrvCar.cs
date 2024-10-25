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
    public partial class prDocDrvCar : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public bool EditMode = false;
        public int doc_id = -1;

        public prDocDrvCar()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private bool HasErrors()
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните номер приказа !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите водителя !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (comboBox2.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите ТС !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
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
                    strSQL = "update pr_drvcar set ";
                    strSQL += "pr_nom='" + textBox1.Text + "',";
                    strSQL += "pr_date='" + fn.DateToStr(dateTimePicker1.Value) + "',";
                    strSQL += "beg_date='" + fn.DateToStr(dateTimePicker2.Value) + "',";
                    strSQL += "drv_id=" + fn.GetDriver(comboBox1) + ",";
                    strSQL += "car_id=" + fn.GetTranspSr(comboBox2) + " ";
                    strSQL += "where pr_id=" + doc_id.ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
                else
                {
                    strSQL = "insert into pr_drvcar (pr_nom, pr_date, beg_date, drv_id, car_id) values (";
                    strSQL += "'" + textBox1.Text + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker1.Value.Date) + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker2.Value.Date) + "',";
                    strSQL += fn.GetDriver(comboBox1) + ",";
                    strSQL += fn.GetTranspSr(comboBox2) + ")";
                    ClSQL.ExecuteSQL(strSQL);
                    doc_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                }
                Close();
            }
        }

        private void prDocDrvCar_Load(object sender, EventArgs e)
        {
            if (EditMode)
            {
                DataRow dr = ClSQL.SelectRow("select * from pr_drvcar where pr_id=" + doc_id);
                textBox1.Text = dr["pr_nom"].ToString();
                dateTimePicker1.Value = DateTime.Parse(dr["pr_date"].ToString());
                dateTimePicker2.Value = DateTime.Parse(dr["beg_date"].ToString());
                fn.UpdateDrivers(comboBox1, (int)dr["drv_id"]);
                fn.UpdateTranspSr(comboBox2, (int)dr["car_id"]);
            }
            else
            {
                textBox1.Text = (fn.GetMaxNom("pr_nom", "pr_drvcar") + 1).ToString();
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;
                fn.UpdateDrivers(comboBox1, 0);
                fn.UpdateTranspSr(comboBox2, 0);
            }
        }
    }
}
