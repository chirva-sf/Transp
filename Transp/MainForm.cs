using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Threading;

namespace Transp
{
    public partial class MainForm : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        Thread hThread;
        BindingSource BSpl = new BindingSource();
        BindingSource BSpr = new BindingSource();
        BindingSource BSst = new BindingSource();
        DataTable DTst = new DataTable();
        private int[] cur_pr_id = new int[10] { -1, -1, -1, -1, -1, -1, -1, -1, -1, -1 };
        private int[] drvar = new int[300];
        private int[] carar = new int[300];
        private int cur_pl_id = -1;
        private bool fl_st = false;

        public MainForm()
        {
            InitializeComponent();
        }

        public void ThreadProc()
        {
            try
            {
                while (true)
                {
                    if (File.Exists(Program.ProgPath + "update.flg"))
                    {
                        System.Windows.Forms.Application.Exit();
                        return;
                    }
                    Thread.Sleep(1000);
                }
            }
            catch
            {
            }
        }

        private void KillThread(object sender, System.EventArgs e)
        {
            if (hThread != null)
            {
                try
                {
                    hThread.Join(10);
                    hThread.Abort();
                }
                catch
                {
                }
                hThread = null;
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            hThread = new Thread(new ThreadStart(ThreadProc));
            hThread.Start();
            this.Closed += new EventHandler(this.KillThread);
            listBox1.SelectedIndex = 0;
            string dep = ClSQL.SelectCell("select name from departments where dep_id=" + Program.UserDepID);
            string username = ClSQL.SelectCell("select fio from users where user_id=" + Program.UserID.ToString());
            this.Text = " Транспорт - " + Program.UserNomDo + " " + dep + " - " + username + (Program.DataBase.ToLower() != "transp"?" - ТЕСТОВАЯ база":"");
            UpdatePlGrid();
            if (Program.UserType > 3)
            {
                toolStripComboBox1.Dispose();
                toolStripLabel1.Dispose();
                toolStripComboBox2.Dispose();
                toolStripLabel2.Dispose();
                toolStripComboBox3.Dispose();
                toolStripLabel3.Dispose();
                toolStripComboBox6.Dispose();
                toolStripLabel6.Dispose();
            }
            else
            {
                toolStripComboBox1.Items.Clear();
                toolStripComboBox1.Items.Add("Все");
                toolStripComboBox4.Items.Clear();
                toolStripComboBox4.Items.Add("Все");
                for (int i = 0; i <= Program.KolvoDO; i++)
                {
                    string s = i.ToString();
                    if (s.Length < 2) s = "0" + s;
                    toolStripComboBox1.Items.Add(Program.FilialPrefix + s);
                    toolStripComboBox4.Items.Add(Program.FilialPrefix + s);
                }
                int id, ti, id2, ti2;
                string ss = fn.LoadUserParam("filter_do");
                toolStripComboBox1.SelectedIndex = ss == "" ? 0 : Int32.Parse(ss);
                ss = fn.LoadUserParam("filter_do2");
                toolStripComboBox4.SelectedIndex = ss == "" ? 0 : Int32.Parse(ss);
                toolStripComboBox2.Items.Clear();
                toolStripComboBox2.Items.Add("Все");
                ss = fn.LoadUserParam("filter_drv");
                ti = 0; id = ss == "" ? 0 : Int32.Parse(ss);
                DataTable dt = ClSQL.SelectSQL("select * from drivers order by nom_do, fio");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    drvar[i] = (int)dt.Rows[i]["drv_id"];
                    if (drvar[i] == id) ti = i + 1;
                    toolStripComboBox2.Items.Add(dt.Rows[i]["nom_do"].ToString() + "  " + dt.Rows[i]["fio"].ToString());
                }
                toolStripComboBox2.SelectedIndex = ti;
                toolStripComboBox3.Items.Clear();
                toolStripComboBox3.Items.Add("Все");
                toolStripComboBox5.Items.Clear();
                toolStripComboBox5.Items.Add("Все");
                ss = fn.LoadUserParam("filter_car");
                ti = 0; id = ss == "" ? 0 : Int32.Parse(ss);
                ss = fn.LoadUserParam("filter_car2");
                ti2 = 0; id2 = ss == "" ? 0 : Int32.Parse(ss);
                dt = ClSQL.SelectSQL("select * from cars order by nom_do, marka");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    carar[i] = (int)dt.Rows[i]["car_id"];
                    if (carar[i] == id) ti = i + 1;
                    if (carar[i] == id2) ti2 = i + 1;
                    toolStripComboBox3.Items.Add(dt.Rows[i]["nom_do"].ToString() + "  " + dt.Rows[i]["marka"].ToString() + "  " + dt.Rows[i]["gosnomer"].ToString());
                    toolStripComboBox5.Items.Add(dt.Rows[i]["nom_do"].ToString() + "  " + dt.Rows[i]["marka"].ToString() + "  " + dt.Rows[i]["gosnomer"].ToString());
                }
                toolStripComboBox3.SelectedIndex = ti;
                toolStripComboBox5.SelectedIndex = ti2;
                ss = fn.LoadUserParam("filter_period");
                toolStripComboBox6.SelectedIndex = (ss == "" || ss == "-1") ? 1 : Int32.Parse(ss);
            }
            if (Program.UserType > 1)
            {
                button1.Dispose();
                button2.Dispose();
                button3.Location = new Point(button3.Location.X, button3.Location.Y - 70);
                button4.Location = new Point(button4.Location.X, button4.Location.Y - 70);
                button5.Location = new Point(button5.Location.X, button5.Location.Y - 70);
                button9.Location = new Point(button5.Location.X, button9.Location.Y - 70);
            }
            switch (Program.UserType)
            {
                case 2:
                    /* Все документы, справочники, отчеты */
                    break;
                case 3:
                    /* Все документы и справочники */
                    tabControl1.TabPages[2].Dispose();
                    break;
                case 4:
                case 5:
                    /* только путевые листы */
                    tabControl1.TabPages[1].Dispose();
                    tabControl1.TabPages[1].Dispose();
                    tabControl1.TabPages[1].Dispose();
                    break;
            }
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (ClSQL.CheckConnection())
            {
                if (Program.UserType <= 3)
                {
                    fn.SaveUserParam("filter_do", toolStripComboBox1.SelectedIndex.ToString());
                    fn.SaveUserParam("filter_do2", toolStripComboBox4.SelectedIndex.ToString());
                    fn.SaveUserParam("filter_drv", toolStripComboBox2.SelectedIndex < 1 ? "-1" : drvar[toolStripComboBox2.SelectedIndex - 1].ToString());
                    fn.SaveUserParam("filter_car", toolStripComboBox3.SelectedIndex < 1 ? "-1" : carar[toolStripComboBox3.SelectedIndex - 1].ToString());
                    fn.SaveUserParam("filter_car2", toolStripComboBox5.SelectedIndex < 1 ? "-1" : carar[toolStripComboBox5.SelectedIndex - 1].ToString());
                    fn.SaveUserParam("filter_period", toolStripComboBox6.SelectedIndex.ToString());
                }
            }
        }

        // ************************ Закладка "Путевые листы" ************************

        private void UpdatePlGrid()
        {
            string strSQL = "select p.pl_id, p.status, p.pl_nom, p.pl_date, p.nom_do, d.name, v.fio, c.marka+'   '+gosnomer as marka ";
            strSQL += "from put_lists p, departments d, drivers v, cars c where p.dep_id=d.dep_id and p.drv_id=v.drv_id and p.car_id=c.car_id ";
            if (Program.UserType > 3)
            {
                strSQL += "and p.nom_do='" + Program.UserNomDo + "' ";
                string ss = DateTime.Now.Date.AddDays(-93).ToString("MM-dd-yyyy");
                strSQL += "and p.pl_date>='" + ss + "' ";
            }
            else
            {
                if (toolStripComboBox1.SelectedIndex > 0) strSQL += "and p.nom_do='" + toolStripComboBox1.Text + "' ";
                if (toolStripComboBox2.SelectedIndex > 0) strSQL += "and p.drv_id=" + drvar[toolStripComboBox2.SelectedIndex - 1].ToString() + " ";
                if (toolStripComboBox3.SelectedIndex > 0) strSQL += "and p.car_id=" + carar[toolStripComboBox3.SelectedIndex - 1].ToString() + " ";
                if (toolStripComboBox6.SelectedIndex > 0)
                {
                    string ss = "";
                    if (toolStripComboBox6.SelectedIndex == 1) ss = DateTime.Now.Date.AddDays(-93).ToString("MM-dd-yyyy");
                    if (toolStripComboBox6.SelectedIndex == 2) ss = DateTime.Now.Date.AddDays(-365).ToString("MM-dd-yyyy");
                    strSQL += "and p.pl_date>='" + ss + "' ";
                }
            }
            strSQL += "order by p.pl_date desc, p.pl_id desc";
            DataTable dt = ClSQL.SelectSQL(strSQL);
            BSpl.DataSource = dt;
            dataGridView1.DataSource = BSpl;
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[0].DisplayIndex = 0;
            dataGridView1.Columns[1].Visible = false;
            dataGridView1.Columns[1].DisplayIndex = 0;
            dataGridView1.Columns[2].HeaderText = "Номер";
            dataGridView1.Columns[2].DisplayIndex = 2;
            dataGridView1.Columns[2].Width = 80;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].HeaderText = "Дата";
            dataGridView1.Columns[3].DisplayIndex = 3;
            dataGridView1.Columns[3].Width = 80;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].HeaderText = "Филиал/ДО";
            dataGridView1.Columns[4].DisplayIndex = 4;
            dataGridView1.Columns[4].Width = 80;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[5].HeaderText = "Подразделение";
            dataGridView1.Columns[5].DisplayIndex = 5;
            dataGridView1.Columns[5].Width = 250;
            dataGridView1.Columns[6].HeaderText = "Водитель";
            dataGridView1.Columns[6].DisplayIndex = 6;
            dataGridView1.Columns[6].Width = 280;
            dataGridView1.Columns[7].HeaderText = "Автомобиль";
            dataGridView1.Columns[7].DisplayIndex = 7;
            dataGridView1.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        private void SavePlIndex()
        {
            if (dataGridView1.CurrentRow != null) cur_pl_id = (int)dataGridView1.CurrentRow.Cells[0].Value;
        }

        private void AddPlDoc()
        {
            if (Program.UserType == 5)
            {
                MessageBox.Show("Нет доступа !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            PutList clDoc = new PutList();
            clDoc.ShowDialog();
            UpdatePlGrid();
            if (clDoc.doc_id != -1)
            {
                int itemFound = BSpl.Find("pl_id", clDoc.doc_id);
                BSpl.Position = itemFound;
            }
        }

        private void EditPlDoc()
        {
            PutList clDoc = new PutList();
            clDoc.EditMode = true;
            clDoc.doc_id = dataGridView1.CurrentRow == null ? (int)dataGridView1[0, 0].Value : (int)dataGridView1.CurrentRow.Cells[0].Value;
            clDoc.ShowDialog();
            int cur_row = BSpl.Position;
            UpdatePlGrid();
            BSpl.Position = cur_row;
        }

        private void DelPlDoc()
        {
            if (Program.UserType == 5)
            {
                MessageBox.Show("Нет доступа !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string doc_id = dataGridView1.CurrentRow == null ? dataGridView1[0, 0].Value.ToString() : dataGridView1.CurrentRow.Cells[0].Value.ToString();
            int cur_row = BSpl.Position;
            if (MessageBox.Show("Вы уверены, что хотите удалить документ ?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ClSQL.ExecuteSQL("delete from put_lists_t where pl_id = " + doc_id);
                ClSQL.ExecuteSQL("delete from put_lists where pl_id = " + doc_id);
                UpdatePlGrid();
                BSpl.Position = cur_row;
            }
        }

        private void PrintPlList()
        {
            if (Program.UserType == 5)
            {
                MessageBox.Show("Нет доступа !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int doc_id = dataGridView1.CurrentRow == null ? (int)dataGridView1[0, 0].Value : (int)dataGridView1.CurrentRow.Cells[0].Value;
            PutList clDoc = new PutList();
            clDoc.PrintPutList(doc_id);
        }

        // ************************ Закладка "Документы" ************************

        private void UpdatePrGrid()
        {
            bool fldo = false;
            string strSQL = "select p.pr_id, p.pr_nom, ";
            if (listBox1.SelectedIndex < 5 || listBox1.SelectedIndex == 8) fldo = true;
            toolStripLabel4.Visible = fldo; toolStripComboBox4.Visible = fldo;
            if (listBox1.SelectedIndex < 2) fldo = true; else fldo = false;
            toolStripLabel5.Visible = fldo; toolStripComboBox5.Visible = fldo;
            if (listBox1.SelectedIndex == 0) // Закрепление ТС за водителем
            {
                strSQL += "p.beg_date, d.nom_do, d.fio, c.marka from pr_drvcar p, drivers d, cars c where ";
                if (toolStripComboBox4.SelectedIndex > 0) strSQL += "d.nom_do='" + toolStripComboBox4.Text + "' and ";
                if (toolStripComboBox5.SelectedIndex > 0) strSQL += "c.car_id='" + carar[toolStripComboBox5.SelectedIndex - 1].ToString() + "' and ";
                strSQL += "p.drv_id=d.drv_id and p.car_id=c.car_id order by d.nom_do, p.beg_date desc, p.pr_nom desc";
            }
            else if (listBox1.SelectedIndex == 1) // Расход топлива по норме
            {
                strSQL += "p.pr_date, c.nom_do, c.marka, c.gosnomer from pr_rasxod p, cars c where ";
                if (toolStripComboBox4.SelectedIndex > 0) strSQL += "c.nom_do='" + toolStripComboBox4.Text + "' and ";
                if (toolStripComboBox5.SelectedIndex > 0) strSQL += "c.car_id='" + carar[toolStripComboBox5.SelectedIndex - 1].ToString() + "' and ";
                strSQL += "p.car_id=c.car_id order by c.nom_do, p.pr_date desc, p.pr_nom desc";
            }
            else if (listBox1.SelectedIndex == 2) // Выдача водителю топливной карты
            {
                strSQL += "p.beg_date, d.nom_do, d.fio, f.fc_nomer from pr_fcdrv p, drivers d, fuel_cards f where ";
                if (toolStripComboBox4.SelectedIndex > 0) strSQL += "d.nom_do='" + toolStripComboBox4.Text + "' and ";
                strSQL += "p.drv_id=d.drv_id and p.fc_id=f.fc_id order by d.nom_do, p.beg_date desc, p.pr_nom desc";
            }
            else if (listBox1.SelectedIndex == 3) // Назначение диспетчера, механика
            {
                strSQL += "p.beg_date, p.nom_do, p.dispatcher, p.mechanic from pr_signs p ";
                if (toolStripComboBox4.SelectedIndex > 0) strSQL += "where p.nom_do='" + toolStripComboBox4.Text + "' ";
                strSQL += "order by p.beg_date desc, p.nom_do, p.pr_nom desc";
            }
            else if (listBox1.SelectedIndex == 4) // Карточки учета ТО и ремонта
            {
                strSQL = "select r.rem_id as pr_id, r.rem_nom, r.nom_do, c.marka, c.gosnomer from cars_rem r, cars c where ";
                if (toolStripComboBox4.SelectedIndex > 0) strSQL += "r.nom_do='" + toolStripComboBox4.Text + "' and ";
                strSQL += "r.car_id = c.car_id order by r.nom_do, c.marka, c.gosnomer";
            }
            else if (listBox1.SelectedIndex == 6) // Договора на поставку топлива
            {
                strSQL += "p.pr_date, p.beg_date, p.supplier from pr_fsupl p where p.pr_id>0 ";
                strSQL += "order by p.pr_date desc, p.pr_nom desc";
            }
            else if (listBox1.SelectedIndex == 7) // Сверка с топливными компаниями
            {
                dataGridView2.DataSource = null; return;
            }
            else if (listBox1.SelectedIndex == 8) // Ввод начальных остатков по ТС
            {
                strSQL += "c.nom_do, c.marka, c.gosnomer from pr_cars_in p, cars c where ";
                if (toolStripComboBox4.SelectedIndex > 0) strSQL += "c.nom_do='" + toolStripComboBox4.Text + "' and ";
                strSQL += "p.car_id=c.car_id order by c.nom_do, c.marka";
            }
            else
            {
                dataGridView2.DataSource = null; return;
            }
            DataTable dt = ClSQL.SelectSQL(strSQL);
            BSpr.DataSource = dt;
            dataGridView2.DataSource = BSpr;
            dataGridView2.Columns[0].Visible = false;
            dataGridView2.Columns[0].DisplayIndex = 0;
            dataGridView2.Columns[1].HeaderText = "Номер";
            dataGridView2.Columns[1].DisplayIndex = 1;
            dataGridView2.Columns[1].Width = 100;
            dataGridView2.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            if (cur_pr_id[listBox1.SelectedIndex] > 0)
            {
                int itemFound = 0;
                itemFound = BSpr.Find("pr_id", cur_pr_id[listBox1.SelectedIndex]);
                BSpr.Position = itemFound;
            }
            if (listBox1.SelectedIndex == 0) // Закрепление ТС за водителем
            {
                dataGridView2.Columns[2].HeaderText = "Действует с";
                dataGridView2.Columns[2].DisplayIndex = 2;
                dataGridView2.Columns[2].Width = 120;
                dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[3].HeaderText = "Филиал/ДО";
                dataGridView2.Columns[3].DisplayIndex = 3;
                dataGridView2.Columns[3].Width = 100;
                dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[4].HeaderText = "Водитель";
                dataGridView2.Columns[4].DisplayIndex = 4;
                dataGridView2.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView2.Columns[5].HeaderText = "ТС";
                dataGridView2.Columns[5].DisplayIndex = 5;
                dataGridView2.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            else if (listBox1.SelectedIndex == 1) // Расход топлива по норме
            {
                dataGridView2.Columns[2].HeaderText = "Дата приказа";
                dataGridView2.Columns[2].DisplayIndex = 2;
                dataGridView2.Columns[2].Width = 120;
                dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[3].HeaderText = "Филиал/ДО";
                dataGridView2.Columns[3].DisplayIndex = 3;
                dataGridView2.Columns[3].Width = 100;
                dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[4].HeaderText = "ТС";
                dataGridView2.Columns[4].DisplayIndex = 4;
                dataGridView2.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView2.Columns[5].HeaderText = "Госномер";
                dataGridView2.Columns[5].DisplayIndex = 5;
                dataGridView2.Columns[5].Width = 120;
                dataGridView2.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            else if (listBox1.SelectedIndex == 2) // Выдача водителю топливной карты
            {
                dataGridView2.Columns[2].HeaderText = "Действует с";
                dataGridView2.Columns[2].DisplayIndex = 2;
                dataGridView2.Columns[2].Width = 120;
                dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[3].HeaderText = "Филиал/ДО";
                dataGridView2.Columns[3].DisplayIndex = 3;
                dataGridView2.Columns[3].Width = 100;
                dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[4].HeaderText = "Водитель";
                dataGridView2.Columns[4].DisplayIndex = 4;
                dataGridView2.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView2.Columns[5].HeaderText = "Карта";
                dataGridView2.Columns[5].DisplayIndex = 5;
                dataGridView2.Columns[5].Width = 150;
                dataGridView2.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            else if (listBox1.SelectedIndex == 3) // Назначение диспетчера, механика
            {
                dataGridView2.Columns[2].HeaderText = "Действует с";
                dataGridView2.Columns[2].DisplayIndex = 2;
                dataGridView2.Columns[2].Width = 120;
                dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[3].HeaderText = "Филиал/ДО";
                dataGridView2.Columns[3].DisplayIndex = 3;
                dataGridView2.Columns[3].Width = 100;
                dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[4].HeaderText = "Диспетчер";
                dataGridView2.Columns[4].DisplayIndex = 4;
                dataGridView2.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView2.Columns[5].HeaderText = "Механик";
                dataGridView2.Columns[5].DisplayIndex = 5;
                dataGridView2.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            else if (listBox1.SelectedIndex == 4) // Карточки учета ТО и ремонта
            {
                dataGridView2.Columns[2].HeaderText = "Филиал/ДО";
                dataGridView2.Columns[2].DisplayIndex = 2;
                dataGridView2.Columns[2].Width = 150;
                dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[3].HeaderText = "Марка";
                dataGridView2.Columns[3].DisplayIndex = 3;
                dataGridView2.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[4].HeaderText = "Госномер";
                dataGridView2.Columns[4].DisplayIndex = 4;
                dataGridView2.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            else if (listBox1.SelectedIndex == 6) // Договора на поставку топлива
            {
                dataGridView2.Columns[2].HeaderText = "Дата";
                dataGridView2.Columns[2].DisplayIndex = 2;
                dataGridView2.Columns[2].Width = 120;
                dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[3].HeaderText = "Действует с";
                dataGridView2.Columns[3].DisplayIndex = 3;
                dataGridView2.Columns[3].Width = 120;
                dataGridView2.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[4].HeaderText = "Компания";
                dataGridView2.Columns[4].DisplayIndex = 4;
                dataGridView2.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }
            else if (listBox1.SelectedIndex == 7) // Сверка с топливными компаниями
            {
            }
            else if (listBox1.SelectedIndex == 8) // Ввод начальных остатков по ТС
            {
                dataGridView2.Columns[2].HeaderText = "Филиал/ДО";
                dataGridView2.Columns[2].DisplayIndex = 2;
                dataGridView2.Columns[2].Width = 120;
                dataGridView2.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView2.Columns[3].HeaderText = "ТС";
                dataGridView2.Columns[3].DisplayIndex = 3;
                dataGridView2.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dataGridView2.Columns[4].HeaderText = "Госномер";
                dataGridView2.Columns[4].DisplayIndex = 4;
                dataGridView2.Columns[4].Width = 120;
                dataGridView2.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void SavePrIndex()
        {
            if (dataGridView2.CurrentRow != null) cur_pr_id[listBox1.SelectedIndex] = (int)dataGridView2.CurrentRow.Cells[0].Value;
        }

        private void AddPrDoc()
        {
            int doc_id = -1;
            if (listBox1.SelectedIndex == 0)
            {
                prDocDrvCar clDoc = new prDocDrvCar();
                clDoc.ShowDialog();
                doc_id = clDoc.doc_id;
            }
            else if (listBox1.SelectedIndex == 1)
            {
                prDocRasxod clDoc = new prDocRasxod();
                clDoc.ShowDialog();
                doc_id = clDoc.doc_id;
            }
            else if (listBox1.SelectedIndex == 2)
            {
                prDocFcDrv clDoc = new prDocFcDrv();
                clDoc.ShowDialog();
                doc_id = clDoc.doc_id;
            }
            else if (listBox1.SelectedIndex == 3)
            {
                prDocSigns clDoc = new prDocSigns();
                clDoc.ShowDialog();
                doc_id = clDoc.doc_id;
            }
            else if (listBox1.SelectedIndex == 4)
            {
                prDocCarsTO clDoc = new prDocCarsTO();
                clDoc.ShowDialog();
                doc_id = clDoc.doc_id;
            }
            else if (listBox1.SelectedIndex == 6)
            {
                prDocFSupl clDoc = new prDocFSupl();
                clDoc.ShowDialog();
                doc_id = clDoc.doc_id;
            }
            else if (listBox1.SelectedIndex == 8)
            {
                prDocCarsIn clDoc = new prDocCarsIn();
                clDoc.ShowDialog();
                doc_id = clDoc.doc_id;
            }
            if (doc_id != -1)
            {
                UpdatePrGrid();
                int itemFound = 0;
                if (listBox1.SelectedIndex == 4)
                {
                    itemFound = BSpr.Find("rem_id", doc_id);
                }
                else
                {
                    itemFound = BSpr.Find("pr_id", doc_id);
                }
                BSpr.Position = itemFound;
            }
        }

        private void EditPrDoc()
        {
            if (dataGridView2.RowCount == 0) return;
            int doc_id = dataGridView2.CurrentRow == null ? (int)dataGridView2[0, 0].Value : (int)dataGridView2.CurrentRow.Cells[0].Value;
            if (listBox1.SelectedIndex == 0)
            {
                prDocDrvCar clDoc = new prDocDrvCar();
                clDoc.EditMode = true;
                clDoc.doc_id = doc_id;
                clDoc.ShowDialog();
            }
            else if (listBox1.SelectedIndex == 1)
            {
                prDocRasxod clDoc = new prDocRasxod();
                clDoc.EditMode = true;
                clDoc.doc_id = doc_id;
                clDoc.ShowDialog();
            }
            else if (listBox1.SelectedIndex == 2)
            {
                prDocFcDrv clDoc = new prDocFcDrv();
                clDoc.EditMode = true;
                clDoc.doc_id = doc_id;
                clDoc.ShowDialog();
            }
            else if (listBox1.SelectedIndex == 3)
            {
                prDocSigns clDoc = new prDocSigns();
                clDoc.EditMode = true;
                clDoc.doc_id = doc_id;
                clDoc.ShowDialog();
            }
            else if (listBox1.SelectedIndex == 4)
            {
                prDocCarsTO clDoc = new prDocCarsTO();
                clDoc.EditMode = true;
                clDoc.doc_id = doc_id;
                clDoc.ShowDialog();
            }
            else if (listBox1.SelectedIndex == 6)
            {
                prDocFSupl clDoc = new prDocFSupl();
                clDoc.EditMode = true;
                clDoc.doc_id = doc_id;
                clDoc.ShowDialog();
            }
            else if (listBox1.SelectedIndex == 8)
            {
                prDocCarsIn clDoc = new prDocCarsIn();
                clDoc.EditMode = true;
                clDoc.doc_id = doc_id;
                clDoc.ShowDialog();
            }
            int cur_row = BSpr.Position;
            UpdatePrGrid();
            BSpr.Position = cur_row;
        }

        private void DelPrDoc()
        {
            if (dataGridView2.RowCount == 0) return;
            string doc_id = dataGridView2.CurrentRow == null ? dataGridView2[0, 0].Value.ToString() : dataGridView2.CurrentRow.Cells[0].Value.ToString();
            int cur_row = BSpl.Position;
            if (MessageBox.Show("Вы уверены, что хотите удалить документ ?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string tname = "";
                string idname = "pr_id";
                if (listBox1.SelectedIndex == 0) tname = "pr_drvcar";
                else if (listBox1.SelectedIndex == 1) tname = "pr_rasxod";
                else if (listBox1.SelectedIndex == 2) tname = "pr_fcdrv";
                else if (listBox1.SelectedIndex == 3) tname = "pr_signs";
                else if (listBox1.SelectedIndex == 4) 
                {
                    ClSQL.ExecuteSQL("delete from cars_rem_t where rem_id=" + doc_id);
                    tname = "cars_rem"; idname = "rem_id"; 
                }
                else if (listBox1.SelectedIndex == 6)
                {
                    ClSQL.ExecuteSQL("delete from pr_fsupl_t where pr_id=" + doc_id);
                    tname = "pr_fsupl"; idname = "pr_id";
                }
                else if (listBox1.SelectedIndex == 8) tname = "pr_cars_in";
                ClSQL.ExecuteSQL("delete from " + tname + " where " + idname + " = " + doc_id);
                UpdatePrGrid();
                BSpl.Position = cur_row;
            }
        }

        private void toolStripComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePrGrid();
            dataGridView2.Focus();
        }

        private void toolStripComboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePrGrid();
            dataGridView2.Focus();
        }


        // ************************ Настройки ***********************

        private void UpdateStGrid()
        {
            fl_st = true;
            object[] rowArray = new object[8];
            string strSQL = "select set_id, setkod, name, curvalue from settings where setkod not like 'dep_%' ";
            if (Program.UserType > 1) strSQL += "and setkod<>'filial_prefix' and setkod<>'kolvo_do' ";
            strSQL += "order by vorder, setkod";
            DataTable dt = ClSQL.SelectSQL(strSQL);
            DTst.Rows.Clear();
            DTst.Columns.Clear();
            DTst.Columns.Add("set_id");
            DTst.Columns.Add("setkod");
            DTst.Columns.Add("name");
            DTst.Columns.Add("curvalue");
            foreach (DataRow r in dt.Rows)
            {
                DataRow row = DTst.NewRow();
                for (int i = 0; i < 4; i++) row[i] = r[i];
                DTst.Rows.Add(row);
            }
            BSst.DataSource = DTst;
            dataGridView3.DataSource = BSst;
            dataGridView3.Columns[0].Visible = false;
            dataGridView3.Columns[0].DisplayIndex = 0;
            dataGridView3.Columns[1].HeaderText = "Код";
            dataGridView3.Columns[1].DisplayIndex = 1;
            dataGridView3.Columns[1].Width = 100;
            dataGridView3.Columns[2].HeaderText = "Наименование";
            dataGridView3.Columns[2].DisplayIndex = 2;
            dataGridView3.Columns[2].Width = 280;
            dataGridView3.Columns[3].HeaderText = "Значение";
            dataGridView3.Columns[3].DisplayIndex = 3;
            dataGridView3.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            fl_st = false;
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            DataRow row = DTst.NewRow();
            DTst.Rows.Add(row);
            dataGridView3.CurrentCell = dataGridView3.Rows[dataGridView3.RowCount - 1].Cells[1];
            dataGridView3.Focus();
        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            dataGridView3.BeginEdit(false);
        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            if (dataGridView3.RowCount > 0)
            {
                if (MessageBox.Show("Вы уверенны, что хотите удалить настройку ?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    ClSQL.ExecuteSQL("delete from settings where set_id=" + DTst.Rows[dataGridView3.CurrentCell.RowIndex][0].ToString());
                    DTst.Rows[dataGridView3.CurrentCell.RowIndex].Delete();
                }
            }
        }

        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int set_id = 0;
            if (fl_st) return;
            if (dataGridView3.CurrentRow == null) return;
            DataRow dr = DTst.Rows[dataGridView3.CurrentCell.RowIndex];
            if (dr[0].ToString() == "")
            {
                ClSQL.ExecuteSQL("insert into settings (setkod, name, curvalue) values ('','','')");
                set_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                ClSQL.ExecuteSQL("update settings set vorder=set_id where set_id=" + set_id.ToString());
                dr[0] = set_id;
            }
            else set_id = Int32.Parse(dr[0].ToString());
            string strSQL = "update settings set setkod='" + dr[1].ToString() + "',";
            strSQL += "name = '" + dr[2].ToString() + "',";
            strSQL += "curvalue = '" + dr[3].ToString() + "' ";
            strSQL += "where set_id = " + set_id.ToString();
            ClSQL.ExecuteSQL(strSQL);
        }

        // ************************

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SavePlIndex();
            SavePrIndex();
            if (tabControl1.SelectedIndex == 0)
            {
                UpdatePlGrid();
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                UpdatePrGrid();
            }
            else if ((tabControl1.SelectedIndex == 3 && Program.UserType < 3) || (tabControl1.SelectedIndex == 2 && Program.UserType == 3))
            {
                UpdateStGrid();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Users_List clList = new Users_List();
            clList.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Departments_List clList = new Departments_List();
            clList.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            TranspSr_List clList = new TranspSr_List();
            clList.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Drivers_List clList = new Drivers_List();
            clList.ShowDialog();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            FuelTypes ft = new FuelTypes();
            ft.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FuelCards_List clList = new FuelCards_List();
            clList.ShowDialog();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePrGrid();
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            AddPrDoc();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            EditPrDoc();
        }

        private void dataGridView2_DoubleClick(object sender, EventArgs e)
        {
            EditPrDoc();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            DelPrDoc();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            AddPlDoc();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            EditPlDoc();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            EditPlDoc();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            DelPlDoc();
        }

        private void dataGridView1_Paint(object sender, PaintEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["status"].Value.ToString() == "1")
                {
                    row.DefaultCellStyle.BackColor = Color.FromArgb(209, 255, 209);
                }
            }

        }

        private void toolStripButton10_Click(object sender, EventArgs e)
        {
            PrintPlList();
        }

        private void toolStripButton11_Click(object sender, EventArgs e)
        {
            UpdatePlGrid();
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePlGrid();
            dataGridView1.Focus();
        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePlGrid();
            dataGridView1.Focus();
        }

        private void toolStripComboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePlGrid();
            dataGridView1.Focus();
        }

        private void toolStripComboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdatePlGrid();
            dataGridView1.Focus();
        }

        // *************** ОТЧЕТЫ ****************

        // Ведомость расхода ГСМ
        private void button6_Click(object sender, EventArgs e)
        {
            RepVedomost rep = new RepVedomost();
            rep.ShowDialog();
        }

        // Расход по топливным картам
        private void button7_Click(object sender, EventArgs e)
        {
            RepFuelCards rep = new RepFuelCards();
            rep.ShowDialog();
        }

        // План проведения ТО
        private void button8_Click(object sender, EventArgs e)
        {
            RepPlanTO rep = new RepPlanTO();
            rep.ShowDialog();
        }

        // Годовой план эксплуатации
        private void button12_Click(object sender, EventArgs e)
        {
            RepPlanEkspl rep = new RepPlanEkspl();
            rep.ShowDialog();
        }

        // Паспорт учета работы ТС
        private void button13_Click(object sender, EventArgs e)
        {
            RepPasportTS rep = new RepPasportTS();
            rep.ShowDialog();
        }

        // Время работы водителей
        private void button15_Click_1(object sender, EventArgs e)
        {
            RepDriversTime rep = new RepDriversTime();
            rep.ShowDialog();
        }

        // Маршруты водителей
        private void button14_Click(object sender, EventArgs e)
        {
            RepMarshrut rep = new RepMarshrut();
            rep.ShowDialog();
        }

        // Транспортный налог
        private void button10_Click(object sender, EventArgs e)
        {
            RepTranspNalog rep = new RepTranspNalog();
            rep.ShowDialog();
        }

        // Загрязнение ОС
        private void button11_Click(object sender, EventArgs e)
        {
            RepZagrOkrSr rep = new RepZagrOkrSr();
            rep.ShowDialog();
        }

        // Время работы водителей
        private void button15_Click(object sender, EventArgs e)
        {
            RepDriversTime rep = new RepDriversTime();
            rep.ShowDialog();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!ClSQL.CheckConnection())
            {
                Thread.Sleep(3000);
                if (!ClSQL.CheckConnection())
                {
                    MessageBox.Show("Прервано соединение с БД !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    System.Diagnostics.Process.GetCurrentProcess().Kill();
                }
            }
        }

    }
}
