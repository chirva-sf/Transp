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
    public partial class Departments_List : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        BindingSource BSource = new BindingSource();
        public int StartMode = 0;
        public int SelDepid = 0;
        public int ShowAllDeps = 0;

        public Departments_List()
        {
            InitializeComponent();
        }

        private void UpdateDepartmentsGrid()
        {
            string strSQL = "select dep_id,nom_do,name from departments ";
            if (Program.UserType > 3 && ShowAllDeps != 1) strSQL += "where dep_id=" + Program.UserDepID + " ";
            strSQL += "order by nom_do,name";
            BSource.DataSource = ClSQL.SelectSQL(strSQL);
            dataGridView1.DataSource = BSource;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "ДО";
            dataGridView1.Columns[1].Width = 50;
            dataGridView1.Columns[2].HeaderText = "Наименование";
            dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        private void AddDepartment()
        {
            Departments_Item clDepart = new Departments_Item();
            clDepart.EditMode = false;
            clDepart.ShowDialog();
            UpdateDepartmentsGrid();
            if (clDepart.dep_id != -1)
            {
                int itemFound = BSource.Find("dep_id", clDepart.dep_id);
                BSource.Position = itemFound;
            }
        }

        private void EditDepartment()
        {
            Departments_Item clDepart = new Departments_Item();
            clDepart.EditMode = true;
            clDepart.dep_id = dataGridView1.CurrentRow == null ? (int)dataGridView1[0, 0].Value : (int)dataGridView1.CurrentRow.Cells[0].Value;
            clDepart.ShowDialog();
            int cur_row = BSource.Position;
            UpdateDepartmentsGrid();
            BSource.Position = cur_row;
        }

        private void DelDepartment()
        {
            string dep_id = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            int cur_row = BSource.Position;
            if (MessageBox.Show("Вы уверены, что хотите удалить подразделение ?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ClSQL.ExecuteSQL("delete from departments where dep_id = " + dep_id);
                UpdateDepartmentsGrid();
                BSource.Position = cur_row;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Departments_Load(object sender, EventArgs e)
        {
            UpdateDepartmentsGrid();
            if (StartMode == 1)
            {
                int itemFound = BSource.Find("dep_id", SelDepid);
                BSource.Position = itemFound;
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            if (StartMode == 1)
            {
                SelDepid = (int)dataGridView1.CurrentRow.Cells[0].Value;
                Close();
            }
            else
            {
                EditDepartment();
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            AddDepartment();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            EditDepartment();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            DelDepartment();
        }
    }
}
