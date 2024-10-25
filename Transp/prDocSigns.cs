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
    public partial class prDocSigns : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public bool EditMode = false;
        public int doc_id = -1;

        public prDocSigns()
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
                    strSQL = "update pr_signs set ";
                    strSQL += "pr_nom='" + textBox1.Text + "',";
                    strSQL += "pr_date='" + fn.DateToStr(dateTimePicker1.Value) + "',";
                    strSQL += "beg_date='" + fn.DateToStr(dateTimePicker2.Value) + "',";
                    strSQL += "nom_do=" + fn.GetNomDO(comboBox1) + ",";
                    strSQL += "dispatcher='" + textBox2.Text + "',";
                    strSQL += "mechanic='" + textBox3.Text + "' ";
                    strSQL += "where pr_id=" + doc_id.ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
                else
                {
                    strSQL = "insert into pr_signs (pr_nom, pr_date, beg_date, nom_do, dispatcher, mechanic) values (";
                    strSQL += "'" + textBox1.Text + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker1.Value.Date) + "',";
                    strSQL += "'" + fn.DateToStr(dateTimePicker2.Value.Date) + "',";
                    strSQL += fn.GetNomDO(comboBox1) + ",";
                    strSQL += "'" + textBox2.Text + "',";
                    strSQL += "'" + textBox3.Text + "')";
                    ClSQL.ExecuteSQL(strSQL);
                    doc_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                }
                ClSQL.ExecuteSQL("EXEC dbo.reCalcSigns " + fn.GetNomDO(comboBox1) + ",'" + fn.DateToStr(dateTimePicker2.Value.Date) + "'");
                Close();
            }
        }

        private void prDocSigns_Load(object sender, EventArgs e)
        {
            if (EditMode)
            {
                DataRow dr = ClSQL.SelectRow("select * from pr_signs where pr_id=" + doc_id);
                textBox1.Text = dr["pr_nom"].ToString();
                dateTimePicker1.Value = DateTime.Parse(dr["pr_date"].ToString());
                dateTimePicker2.Value = DateTime.Parse(dr["beg_date"].ToString());
                fn.UpdateNomDO(comboBox1, dr["nom_do"].ToString());
                textBox2.Text = dr["dispatcher"].ToString();
                textBox3.Text = dr["mechanic"].ToString();
            }
            else
            {
                textBox1.Text = (fn.GetMaxNom("pr_nom", "pr_signs") + 1).ToString();
                dateTimePicker1.Value = DateTime.Now;
                dateTimePicker2.Value = DateTime.Now;
                fn.UpdateNomDO(comboBox1, "");
            }
        }
    }
}
