using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;

namespace Transp
{
    public partial class FirstStart : Form
    {
        private XmlDocument document;

        public FirstStart()
        {
            InitializeComponent();
        }

        private void AddAttr(string AttrName, string AttrValue)
        {
            XmlNode element = document.CreateElement(AttrName);
            document.DocumentElement.AppendChild(element);
            XmlAttribute attribute = document.CreateAttribute("value");
            attribute.Value = AttrValue;
            element.Attributes.Append(attribute);
        }

        private void CreateBase()
        {
            if (!File.Exists(Program.ProgPath + "transp.sql"))
            {
                MessageBox.Show("Отсутствует файл transp.sql !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                File.Delete(Program.ProgPath + "transp.xml");
            }
            else
            {
                ClSQL ClSQL = new ClSQL("server=" + textBox1.Text + ";uid=" + textBox2.Text + ";pwd=" + textBox3.Text);
                ClSQL.ExecuteSQL(string.Format("CREATE DATABASE [{0}]", textBox4.Text));
                ClSQL.ExecuteSQL(string.Format("USE [{0}]", textBox4.Text));
                string[] ss = File.ReadAllLines(Program.ProgPath + "transp.sql", Encoding.UTF8);
                string sql = "", sm;
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
                                sm = s;
                                if (sm.Trim() == "[ust_id] int IDENTITY(1, 1) NOT NULL,") sm = sm.Replace("1, 1","0, 1");
                                sql = sql + sm + Environment.NewLine;
                            }
                        }
                    }
                }
                string nom_do = textBox5.Text;
                ClSQL.ExecuteSQL("insert into departments (nom_do,name) values ('" + nom_do + "','Отдел информационного обеспечения')");
                ClSQL.ExecuteSQL("insert into users (nom_do,fio,dep_id,user_login,user_type) values ('" + nom_do + "','Администратор',1,'" + Environment.UserName + "',1)");
                ClSQL.ExecuteSQL("insert into settings (setkod,name,vorder) values ('filial_name','Наименование филиала',1)");
                ClSQL.ExecuteSQL("insert into settings (setkod,name,vorder) values ('filial_addr','Адрес филиала',2)");
                ClSQL.ExecuteSQL("insert into settings (setkod,name,vorder) values ('filial_phone','Телефон филиала',3)");
                ClSQL.ExecuteSQL("insert into settings (setkod,name,curvalue,vorder) values ('filial_prefix','Префикс филиала','" + nom_do.Substring(0, 2) + "',4)");
                ClSQL.ExecuteSQL("insert into settings (setkod,name,vorder) values ('kolvo_do','Количество доп.офисов',5)");
                ClSQL.ExecuteSQL("insert into settings (setkod,name,vorder) values ('sign1','Заместитель директора',6)");
                ClSQL.ExecuteSQL("insert into settings (setkod,name,vorder) values ('sign2','Зам.гл.бухгалтера - нач.ОБУиО',7)");
                ClSQL.ExecuteSQL("insert into settings (setkod,name,vorder) values ('empty_lists_days','Кол-во дней разрешено для записи пустых пут.листов',8)");
                ClSQL.ExecuteSQL("insert into settings (setkod,name,curvalue,vorder) values ('enable_downtime','Разрешить часы простоя','Нет',9)");
                ClSQL.ExecuteSQL("insert into settings (setkod,name,vorder) values ('rest_time','Время на перерыв для отдыха и питания (мин)',10)");
                ClSQL.ExecuteSQL("insert into settings (setkod,name,vorder) values ('plm_time','Время м/у выдачей пут.листа и выездом (мин)',11)");
                ClSQL.ExecuteSQL("insert into usrtypes (name) values ('Нет доступа')");
                ClSQL.ExecuteSQL("insert into usrtypes (name) values ('Полный доступ (администраторы)')");
                ClSQL.ExecuteSQL("insert into usrtypes (name) values ('Все документы, справочники, отчеты')");
                ClSQL.ExecuteSQL("insert into usrtypes (name) values ('Все документы и справочники')");
                ClSQL.ExecuteSQL("insert into usrtypes (name) values ('Только путевые листы')");
                ClSQL.ExecuteSQL("insert into usrtypes (name) values ('Путевые листы ограниченно (водители)')");
                ClSQL.ExecuteSQL("insert into fuel_types (ft_name) values ('АИ-92')");
                ClSQL.ExecuteSQL("insert into fuel_types (ft_name) values ('АИ-93')");
                ClSQL.ExecuteSQL("insert into fuel_types (ft_name) values ('АИ-95')");
                ClSQL.ExecuteSQL("insert into fuel_types (ft_name) values ('АИ-96')");
                ClSQL.ExecuteSQL("insert into fuel_types (ft_name) values ('АИ-98')");
                ClSQL.ExecuteSQL("insert into fuel_types (ft_name) values ('ДТ')");
                ClSQL.ExecuteSQL("insert into pr_fsupl (pr_nom,supplier) values ('Б/н','Без договора')");
                ClSQL.ExecuteSQL("insert into version (ver) values ('" + Program.ProgVer + "')");
                ClSQL.DisconnectSQL();
            }
        }

        private bool CheckDataError()
        {
            string err = "";
            if (textBox1.Text == "") err = "Заполните имя сервера !";
            else if (textBox4.Text == "") err = "Заполните название БД !";
            else if (textBox2.Text == "") err = "Заполните пользователя БД !";
            else if (textBox3.Text == "") err = "Заполните пароль к БД !";
            else if (textBox5.Text == "") err = "Заполните номер филиала !";
            else if (textBox5.Text.Length < 4) err = "Номер филиала должен быть из 4х цифр !";
            if (err == "")
            {
                try 
                { 
                    SqlConnection ClCN = new SqlConnection("server=" + textBox1.Text + ";uid=" + textBox2.Text + ";pwd=" + textBox3.Text);
                    ClCN.Open(); ClCN.Close();
                }
                catch { err = "Не удалось подключиться к серверу !"; }
            }
            if (err == "") return false; else { MessageBox.Show(err, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); return true; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false; button2.Enabled = false;
            if (CheckDataError()) { button1.Enabled = true; button2.Enabled = true; return; }
            string xml_file = Program.ProgPath + "transp.xml";
            String[] args = Environment.GetCommandLineArgs();
            string ProgPath = args[0].Substring(0, args[0].LastIndexOf('\\') + 1);
            XmlTextWriter textWritter = new XmlTextWriter(xml_file, Encoding.ASCII);
            textWritter.WriteStartDocument();
            textWritter.WriteStartElement("Settings");
            textWritter.WriteEndElement();
            textWritter.Close();
            document = new XmlDocument();
            document.Load(xml_file);
            AddAttr("server", textBox1.Text);
            AddAttr("database", textBox4.Text);
            AddAttr("dbuser", Convert.ToBase64String(Encoding.UTF8.GetBytes("za" + textBox2.Text + "m!")));
            AddAttr("dbpass", Convert.ToBase64String(Encoding.UTF8.GetBytes("ch" + textBox3.Text + "3s")));
            document.Save(xml_file);
            bool flbase = false;
            try
            {
                SqlConnection ClCN = new SqlConnection("server=" + textBox1.Text + ";uid=" + textBox2.Text + ";pwd=" + textBox3.Text + ";database=" + textBox4.Text);
                ClCN.Open();
                SqlDataAdapter da = new SqlDataAdapter("select ver from version", ClCN);
                DataSet ds = new DataSet("clsql");
                da.Fill(ds, "cl_sql");
                ClCN.Close();
            }
            catch { flbase = true; }
            if (flbase) CreateBase();
            MessageBox.Show("Настройка завершена. Перезапустите программу.");
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FirstStart_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
        }
    }
}
