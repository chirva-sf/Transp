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
    public partial class Departments_Item : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public bool EditMode = false;
        public int dep_id = -1;

        public Departments_Item()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "")
            {
                MessageBox.Show("Выберите филиал/доп.офис !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else if (textBox1.Text == "")
            {
                MessageBox.Show("Введите название подразделения !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else if (textBox2.Text == "")
            {
                MessageBox.Show("Введите ФИО руководителя подразделения !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                string strSQL = "";
                if (EditMode)
                {
                    strSQL = "update departments set ";
                    strSQL += "nom_do = '" + fn.GetNomDO(comboBox1) + "',";
                    strSQL += "name='" + textBox1.Text + "',";
                    strSQL += "head='" + textBox2.Text + "',";
                    strSQL += "address='" + textBox3.Text + "' ";
                    strSQL += "where dep_id=" + dep_id.ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
                else
                {
                    strSQL = "insert into departments (nom_do,name,head,address) values (";
                    strSQL += "'" + fn.GetNomDO(comboBox1) + "',";
                    strSQL += "'" + textBox1.Text + "',";
                    strSQL += "'" + textBox2.Text + "',";
                    strSQL += "'" + textBox3.Text + "')";
                    ClSQL.ExecuteSQL(strSQL);
                    dep_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                }
                string s = ClSQL.SelectCell("select set_id from settings where setkod='dep_resttime_" + dep_id.ToString() + "'");
                if (s == "") strSQL = "insert into settings (setkod, name, curvalue) values ('dep_resttime_" + dep_id.ToString() + "','Время на отдых/питание','" + textBox4.Text + "')";
                        else strSQL = "update settings set curvalue='" + textBox4.Text + "' where setkod='dep_resttime_" + dep_id.ToString() + "'";
                ClSQL.ExecuteSQL(strSQL);
                s = ClSQL.SelectCell("select set_id from settings where setkod='dep_plmtime_" + dep_id.ToString() + "'");
                if (s == "") strSQL = "insert into settings (setkod, name, curvalue) values ('dep_plmtime_" + dep_id.ToString() + "','Время м/у выдачей ПЛ и выездом','" + textBox5.Text + "')";
                else strSQL = "update settings set curvalue='" + textBox5.Text + "' where setkod='dep_plmtime_" + dep_id.ToString() + "'";
                ClSQL.ExecuteSQL(strSQL);
                Close();
            }
        }

        private void Department_Item_Load(object sender, EventArgs e)
        {
            if (EditMode)
            {
                DataRow dr = ClSQL.SelectRow("select * from departments where dep_id=" + dep_id.ToString());
                fn.UpdateNomDO(comboBox1, dr["nom_do"].ToString());
                textBox1.Text = dr["name"].ToString();
                textBox2.Text = dr["head"].ToString();
                textBox3.Text = dr["address"].ToString();
                textBox4.Text = ClSQL.SelectCell("select curvalue from settings where setkod='dep_resttime_" + dep_id.ToString() + "'");
                textBox5.Text = ClSQL.SelectCell("select curvalue from settings where setkod='dep_plmtime_" + dep_id.ToString() + "'");
            }
            else
            {
                fn.UpdateNomDO(comboBox1, "");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
