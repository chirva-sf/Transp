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
    public partial class TranspSr_List : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        BindingSource BSource = new BindingSource();

        public TranspSr_List()
        {
            InitializeComponent();
        }

        private void UpdateGrid()
        {
            string strSQL = "select c.car_id, c.nom_do, c.marka, c.gosnomer, f.ft_name, c.mileage_to, c.beg_date, c.status ";
            strSQL += "from cars c, fuel_types f where c.ft_id = f.ft_id ";
            if (checkBox1.Checked) strSQL += "and c.status<>1 ";
            strSQL += "order by c.nom_do, c.marka";
            BSource.DataSource = ClSQL.SelectSQL(strSQL);
            dataGridView1.DataSource = BSource;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Филиал/ДО";
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].HeaderText = "Марка, модель";
            dataGridView1.Columns[2].Width = 150;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.Columns[3].HeaderText = "Госномер";
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].HeaderText = "Топливо";
            dataGridView1.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[5].HeaderText = "ТО";
            dataGridView1.Columns[5].Width = 60;
            dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[6].HeaderText = "Нач.эксп.";
            dataGridView1.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[7].Visible = false;
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
            TranspSr_Item clItem = new TranspSr_Item();
            clItem.EditMode = false;
            clItem.ShowDialog();
            UpdateGrid();
            if (clItem.item_id != -1)
            {
                int itemFound = BSource.Find("car_id", clItem.item_id);
                BSource.Position = itemFound;
            }
        }

        private void EditItem()
        {
            TranspSr_Item clItem = new TranspSr_Item();
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
            if (MessageBox.Show("Вы уверены, что хотите удалить ТС ?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ClSQL.ExecuteSQL("delete from cars where car_id = " + item_id);
                UpdateGrid();
                BSource.Position = cur_row;
            }
        }

        private void TranspSr_List_Load(object sender, EventArgs e)
        {
            if (fn.LoadUserParam("filter_tsr") == "1") checkBox1.Checked = true;
            UpdateGrid();
        }

        private void TranspSr_List_FormClosed(object sender, FormClosedEventArgs e)
        {
            fn.SaveUserParam("filter_tsr", (checkBox1.Checked ? "1" : "0"));
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

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            UpdateGrid();
        }

    }
}
