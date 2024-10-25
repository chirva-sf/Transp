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
    public partial class Users_List : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        BindingSource BSource = new BindingSource();

        public Users_List()
        {
            InitializeComponent();
        }

        private void UpdateUsersGrid()
        {
            string strSQL = "select us.user_id, us.nom_do, us.fio, ut.name, user_login as user_type ";
            strSQL += "from users us, usrtypes ut where us.user_type=ut.ust_id order by us.nom_do, us.fio";
            BSource.DataSource = ClSQL.SelectSQL(strSQL);
            dataGridView1.DataSource = BSource;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[4].Visible = false;
            dataGridView1.Columns[1].HeaderText = "ДО";
            dataGridView1.Columns[1].Width = 50;
            dataGridView1.Columns[2].HeaderText = "ФИО";
            dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[3].HeaderText = "Тип";
            dataGridView1.Columns[3].Width = 225;
        }

        private void AddUser()
        {
            Users_Item clUser = new Users_Item();
            clUser.EditMode = false;
            clUser.ShowDialog();
            UpdateUsersGrid();
            if (clUser.user_id != -1)
            {
                int itemFound = BSource.Find("user_id", clUser.user_id);
                BSource.Position = itemFound;
            }
        }

        private void EditUser()
        {
            Users_Item clUser = new Users_Item();
            clUser.EditMode = true;
            clUser.user_id = dataGridView1.CurrentRow == null ? (int)dataGridView1[0, 0].Value : (int)dataGridView1.CurrentRow.Cells[0].Value;
            clUser.ShowDialog();
            int cur_row = BSource.Position;
            UpdateUsersGrid();
            BSource.Position = cur_row;
        }

        private void DelUser()
        {
            string user_id = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            int cur_row = BSource.Position;
            if (MessageBox.Show("Вы уверены, что хотите удалить пользователя ?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ClSQL.ExecuteSQL("delete from users where user_id = " + user_id);
                UpdateUsersGrid();
                BSource.Position = cur_row;
            }
        }

        private void Users_List_Load(object sender, EventArgs e)
        {
            UpdateUsersGrid();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            AddUser();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            EditUser();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            DelUser();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            EditUser();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

    }
}
