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
    public partial class Users_Item : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public bool EditMode = false;
        public int user_id = -1;

        public Users_Item()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите филиал/доп.офис !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните ФИО !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else if (comboBox2.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите подразделение !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else if (textBox3.Text == "")
            {
                MessageBox.Show("Заполните логин в домене !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else if (comboBox3.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите тип пользователя !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                string strSQL = "";
                if (EditMode)
                {
                    strSQL = "update users set ";
                    strSQL += "nom_do = '" + fn.GetNomDO(comboBox1) + "',";
                    strSQL += "fio='" + textBox1.Text + "',";
                    strSQL += "dep_id=" + fn.GetDepartment(comboBox2) + ",";
                    strSQL += "office='" + textBox2.Text + "',";
                    strSQL += "user_login='" + textBox3.Text + "',";
                    strSQL += "user_type=" + fn.GetUsersType(comboBox3) + " ";
                    strSQL += "where user_id=" + user_id.ToString();
                    ClSQL.ExecuteSQL(strSQL);
                }
                else
                {
                    strSQL = "insert into users (nom_do,fio,dep_id,office,user_login,user_type) values (";
                    strSQL += "'" + fn.GetNomDO(comboBox1) + "',";
                    strSQL += "'" + textBox1.Text + "',";
                    strSQL += fn.GetDepartment(comboBox2) + ",";
                    strSQL += "'" + textBox2.Text + "',";
                    strSQL += "'" + textBox3.Text + "',";
                    strSQL += fn.GetUsersType(comboBox3) + ")";
                    ClSQL.ExecuteSQL(strSQL);
                    user_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                }
                Close();
            }
        }


        private void User_Item_Load(object sender, EventArgs e)
        {
            if (EditMode)
            {
                DataRow dr = ClSQL.SelectRow("select * from users where user_id=" + user_id);
                textBox1.Text = dr["fio"].ToString();
                textBox2.Text = dr["office"].ToString();
                textBox3.Text = dr["user_login"].ToString();
                textBox4.Text = dr["info"].ToString();
                label8.Text = dr["last_visit"].ToString();
                fn.UpdateNomDO(comboBox1, dr["nom_do"].ToString());
                fn.UpdateDepartments(comboBox2, (int)dr["dep_id"]);
                fn.UpdateUsersTypes(comboBox3, (int)dr["user_type"]);
            }
            else
            {
                fn.UpdateNomDO(comboBox1, "");
                fn.UpdateDepartments(comboBox2, 0);
                fn.UpdateUsersTypes(comboBox3, 0);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            fn.UpdateDepartmentsRef(comboBox2);
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            Close();
        }

    }
}
