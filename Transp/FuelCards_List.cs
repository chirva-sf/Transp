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
    public partial class FuelCards_List : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        BindingSource BSource = new BindingSource();
        public int StartMode = 0;
        public int selid = 0;

        public FuelCards_List()
        {
            InitializeComponent();
        }

        private void UpdateGrid()
        {
            string strSQL = "select f.fc_id, f.nom_do, f.fc_nomer, t.ft_name, f.status ";
            strSQL += "from fuel_cards f, fuel_types t where f.ft_id=t.ft_id ";
            if (checkBox1.Checked) strSQL += "and f.status<>1 ";
            strSQL += "order by f.fc_nomer";
            DataTable dt = ClSQL.SelectSQL(strSQL);
            BSource.DataSource = dt;
            dataGridView1.DataSource = BSource;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Ном.ДО";
            dataGridView1.Columns[1].Width = 80;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].HeaderText = "Номер карты";
            dataGridView1.Columns[2].Width = 135;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].HeaderText = "Вид топлива";
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].Visible = false;
            if (StartMode == 1 && selid > 0)
            {
                int itemFound = BSource.Find("fc_id", selid);
                BSource.Position = itemFound;
            }
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
            FuelCards_Item clItem = new FuelCards_Item();
            clItem.EditMode = false;
            clItem.ShowDialog();
            UpdateGrid();
            if (clItem.item_id != -1)
            {
                int itemFound = BSource.Find("fc_id", clItem.item_id);
                BSource.Position = itemFound;
            }
        }

        private void EditItem()
        {
            FuelCards_Item clItem = new FuelCards_Item();
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
            if (MessageBox.Show("Вы уверены, что хотите удалить карту ?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ClSQL.ExecuteSQL("delete from fuel_cards where fc_id = " + item_id);
                UpdateGrid();
                BSource.Position = cur_row;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            selid = 0;
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
            if (StartMode == 1)
            {
                selid = (int)dataGridView1.CurrentRow.Cells[0].Value;
                Close();
            }
            else
            {
                EditItem();
            }
        }

        private void FuelCards_List_Load(object sender, EventArgs e)
        {
            if (StartMode != 1) button2.Visible = false;
            if (fn.LoadUserParam("filter_fc") == "1") checkBox1.Checked = true;
            UpdateGrid();
        }

        private void FuelCards_List_FormClosed(object sender, FormClosedEventArgs e)
        {
            fn.SaveUserParam("filter_fc", (checkBox1.Checked ? "1" : "0"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (StartMode == 1)
            {
                selid = (int)dataGridView1.CurrentRow.Cells[0].Value;
                Close();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            UpdateGrid();
        }

    }
}
