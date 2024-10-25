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
    public partial class prDocFuelPrice : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public bool EditMode = false;
        public int doc_id = -1;

        public prDocFuelPrice()
        {
            InitializeComponent();
        }

        private bool HasErrors()
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните номер приказа !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите тип топлива !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (!fn.ChkDoubleVal(textBox2))
            {
                MessageBox.Show("Некорректная цена !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
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
                    strSQL = "update pr_fuelprice set ";
                    strSQL += "pr_nom='" + textBox1.Text + "',";
                    strSQL += "pr_date='" + fn.DateToStr(dateTimePicker1.Value) + "',";
                    strSQL += "ft_id=" + fn.GetFuelType(comboBox1) + ",";
                    strSQL += "beg_date='" + fn.DateToStr(dateTimePicker2.Value) + "',";
                    strSQL += "price=" + fn.NumStr(textBox2) + " ";
                    strSQL += "where pr_id=" + doc_id.ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
                else
                {
                    strSQL = "insert into pr_fuelprice (pr_nom, pr_date, ft_id, beg_date, price) values (";
                    strSQL += "'" + textBox1.Text + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker1.Value.Date) + "',";
                    strSQL += fn.GetFuelType(comboBox1) + ",";
                    strSQL += "'" + fn.DateToStr(dateTimePicker2.Value.Date) + "',";
                    strSQL += fn.NumStr(textBox2) + ")";
                    ClSQL.ExecuteSQL(strSQL);
                    doc_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                }
                Close();
            }
        }

        private void prDocFuelPrice_Load(object sender, EventArgs e)
        {
            if (EditMode)
            {
                DataRow dr = ClSQL.SelectRow("select * from pr_fuelprice where pr_id=" + doc_id);
                textBox1.Text = dr["pr_nom"].ToString();
                dateTimePicker1.Value = DateTime.Parse(dr["pr_date"].ToString());
                dateTimePicker2.Value = DateTime.Parse(dr["beg_date"].ToString());
                fn.UpdateFuelType(comboBox1, (int)dr["ft_id"]);
                textBox2.Text = dr["price"].ToString();
            }
            else
            {
                textBox1.Text = (fn.GetMaxNom("pr_nom", "pr_fuelprice") + 1).ToString();
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;
                fn.UpdateFuelType(comboBox1, 0);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
