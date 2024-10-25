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
    public partial class FuelCards_Item : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public bool EditMode = false;
        public int item_id = -1;

        public FuelCards_Item()
        {
            InitializeComponent();
        }

        private bool HasErrors()
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните номер карты !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите вид топлива !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (!fn.ChkIntVal(textBox2))
            {
                MessageBox.Show("Не верно введен лимит !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
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
                    strSQL = "update fuel_cards set ";
                    strSQL += "ft_id=" + fn.GetFuelType(comboBox1) + ",";
                    strSQL += "nom_do='" + fn.GetNomDO(comboBox2) + "',";
                    strSQL += "fc_nomer='" + textBox1.Text + "',";
                    strSQL += "fc_limit=" + fn.NumStr(textBox2) + ", ";
                    strSQL += "status=" + (checkBox1.Checked ? 1 : 0) + " ";
                    strSQL += "where fc_id=" + item_id.ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
                else
                {
                    strSQL = "insert into fuel_cards (ft_id, nom_do, fc_nomer, fc_limit, status) values (";
                    strSQL += fn.GetFuelType(comboBox1) + ",";
                    strSQL += fn.GetNomDO(comboBox2) + ",";
                    strSQL += "'" + textBox1.Text + "',";
                    strSQL += "'" + textBox2.Text + "',";
                    strSQL += (checkBox1.Checked ? 1 : 0) + ")";
                    ClSQL.ExecuteSQL(strSQL);
                    item_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                }
                Close();
            }
        }

        private void FuelCards_Item_Load(object sender, EventArgs e)
        {
            if (EditMode)
            {
                DataRow dr = ClSQL.SelectRow("select * from fuel_cards where fc_id=" + item_id);
                fn.UpdateFuelType(comboBox1, (int)dr["ft_id"]);
                textBox1.Text = dr["fc_nomer"].ToString();
                textBox2.Text = dr["fc_limit"].ToString();
                fn.UpdateNomDO(comboBox2, dr["nom_do"].ToString());
                checkBox1.Checked = dr["status"].ToString() == "1" ? true : false;
            }
            else
            {
                fn.UpdateFuelType(comboBox1, 0);
                fn.UpdateNomDO(comboBox2, "");
            }
        }
    }
}
