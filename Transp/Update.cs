using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;

namespace Transp
{
    public partial class Update : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        public int res = -1;

        public Update()
        {
            InitializeComponent();
        }

        private void DoUpdate()
        {
            string[] ss = File.ReadAllLines(Program.ProgPath + "update.sql", Encoding.UTF8);
            string sql = "";
            foreach (string s in ss)
            {
                if (s.Length > 0)
                {
                    if (s.Substring(0, 1) != "-" && s != "ON [PRIMARY]")
                    {
                        if (s == "GO")
                        {
                            if (sql.Trim() != "")
                            {
                                ClSQL.ExecuteSQL(sql);
                                sql = "";
                            }
                        }
                        else
                        {
                            sql = sql + s + Environment.NewLine;
                        }
                    }
                }
            }
            File.Delete(Program.ProgPath + "update.sql");
            res = 1;
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false; button2.Enabled = false;
            DoUpdate();
            button1.Enabled = true; button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

    }
}
