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
    public partial class Drivers_List : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        BindingSource BSource = new BindingSource();

        public Drivers_List()
        {
            InitializeComponent();
        }

        private void UpdateGrid()
        {
            string strSQL = "select d.drv_id, d.nom_do, d.tab_no, d.fio, p.name, d.status ";
            strSQL += "from drivers d, departments p where d.dep_id=p.dep_id ";
            if (checkBox1.Checked) strSQL += "and d.status<>1 ";
            strSQL += "order by d.nom_do, d.fio";
            BSource.DataSource = ClSQL.SelectSQL(strSQL);
            dataGridView1.DataSource = BSource;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Фил/ДО";
            dataGridView1.Columns[1].Width = 60;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].HeaderText = "Таб.ном.";
            dataGridView1.Columns[2].Width = 60;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].HeaderText = "ФИО";
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.Columns[4].HeaderText = "Подразделение";
            dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.Columns[5].Visible = false;
        }

        private void dataGridView1_Paint(object sender, PaintEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["status"].Value.ToString() == "1")
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 209, 209);
                }
            }
        }

        private void AddItem()
        {
            Drivers_Item clItem = new Drivers_Item();
            clItem.EditMode = false;
            clItem.ShowDialog();
            UpdateGrid();
            if (clItem.item_id != -1)
            {
                int itemFound = BSource.Find("drv_id", clItem.item_id);
                BSource.Position = itemFound;
            }
        }

        private void EditItem()
        {
            Drivers_Item clItem = new Drivers_Item();
            clItem.EditMode = true;
            clItem.item_id = dataGridView1.CurrentRow == null ? (int)dataGridView1[0, 0].Value : (int)dataGridView1.CurrentRow.Cells[0].Value;
            clItem.ShowDialog();
            int cur_row = BSource.Position;
            UpdateGrid();
            BSource.Position = cur_row;
        }

        private void DelItem()
        {
            string item_id = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            int cur_row = BSource.Position;
            if (MessageBox.Show("Вы уверены, что хотите удалить водителя ?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ClSQL.ExecuteSQL("delete from drivers where drv_id = " + item_id);
                UpdateGrid();
                BSource.Position = cur_row;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            AddItem();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            EditItem();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            DelItem();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            EditItem();
        }

        private void Drivers_List_Load(object sender, EventArgs e)
        {
            if (fn.LoadUserParam("filter_drvs") == "1") checkBox1.Checked = true;
            UpdateGrid();
        }

        private void Drivers_List_FormClosed(object sender, FormClosedEventArgs e)
        {
            fn.SaveUserParam("filter_drvs", (checkBox1.Checked ? "1" : "0"));
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            UpdateGrid();
        }

    }
}
