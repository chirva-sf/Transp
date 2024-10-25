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
    public partial class FuelTypes : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        BindingSource BSpl = new BindingSource();
        DataTable DTpl = new DataTable();

        public FuelTypes()
        {
            InitializeComponent();
        }

        private void FuelTypesTableCreate()
        {
            DTpl.Columns.Clear();
            DTpl.Columns.Add("ft_id");
            DTpl.Columns.Add("ft_name");
            BSpl.DataSource = DTpl;
            dataGridView1.AutoGenerateColumns = false;
            foreach (DataColumn cl in DTpl.Columns)
            {
                DataGridViewColumn column = new DataGridViewTextBoxColumn();
                column.DataPropertyName = cl.ColumnName;
                dataGridView1.Columns.Add(column);
            }
            dataGridView1.DataSource = BSpl;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Вид топлива";
            dataGridView1.Columns[1].DisplayIndex = 0;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[1].Width = dataGridView1.Width - 12;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string strSQL;
            bool fl, err;
            err = false;
            DataTable dt = ClSQL.SelectSQL("select * from fuel_types");
            foreach (DataRow dr in dt.Rows)
            {
                fl = false; for (int i = 0; i < DTpl.Rows.Count; i++) if (DTpl.Rows[i][1].ToString() == dr[1].ToString()) { fl = true; break; }
                if (!fl)
                {
                    int f1 = ClSQL.SelectIntCell("select top 1 car_id from cars where ft_id=" + dr[0].ToString());
                    int f2 = ClSQL.SelectIntCell("select top 1 fc_id from fuel_cards where ft_id=" + dr[0].ToString());
                    int f3 = ClSQL.SelectIntCell("select top 1 pr_id from pr_fsupl_t where ft_id=" + dr[0].ToString());
                    if (f1 + f2 + f3 > 0)
                    {
                        MessageBox.Show("Нельзя менять или удалять вид топлива " + dr[1].ToString() + " т.к. он используется !");
                        err = true;
                    }
                    else
                    {
                        strSQL = "delete from fuel_types where ft_id=" + dr[0].ToString();
                        ClSQL.ExecuteSQL(strSQL);
                    }
                }
            }
            if (!err)
            {
                for (int i = 0; i < DTpl.Rows.Count; i++)
                {
                    DataRow dr = DTpl.Rows[i];
                    if (dr[1].ToString() != "")
                    {
                        int id = ClSQL.SelectIntCell("select ft_id from fuel_types where ft_name='" + dr[1].ToString() + "'");
                        if (id < 1)
                        {
                            strSQL = "insert into fuel_types (ft_name) values ('" + dr[1].ToString() + "')";
                            ClSQL.ExecuteSQL(strSQL);
                        }
                    }
                }
                Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FuelTypes_Load(object sender, EventArgs e)
        {
            FuelTypesTableCreate();
            object[] rowarr = new object[2];
            DataTable dt = ClSQL.SelectSQL("select * from fuel_types order by ft_id");
            foreach (DataRow r in dt.Rows)
            {
                DataRow row = DTpl.NewRow();
                rowarr[0] = r["ft_id"];
                rowarr[1] = r["ft_name"];
                row.ItemArray = rowarr;
                DTpl.Rows.Add(row);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            DataRow row = DTpl.NewRow();
            DTpl.Rows.Add(row);
            dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[1];
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            dataGridView1.BeginEdit(false);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 0)
            {
                DTpl.Rows[dataGridView1.CurrentCell.RowIndex].Delete();
            }
        }
    }
}
