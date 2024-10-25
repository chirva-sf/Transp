using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Data;
using System.IO;
using System.Xml;
using System.Net;
using System.Security.Permissions;
using System.Threading;
using System.Text;

[assembly: SecurityPermission(SecurityAction.RequestMinimum, ControlThread = true)]
namespace Transp
{
    static class Program
    {
        public static string ProgVer = "2.53";
        public static string ProgPath;
        public static string TmpPath;
        public static ClSQL ClSQL;
        public static ClFunc ClFunc;
        public static string DataBase = "";
        public static string FilialPrefix = "";
        public static int KolvoDO = 0;
        public static int EmpLstDays = 0;
        public static int UserID = -1;
        public static string UserNomDo = "";
        public static int UserDepID = -1;
        public static int UserType = -1;
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            int ii;
            string Server = "", DBUser = "sa", DBPass = "MSsql05S";
            String[] args = Environment.GetCommandLineArgs();
            ProgPath = args[0].Substring(0, args[0].LastIndexOf('\\') + 1);
            TmpPath = Environment.GetEnvironmentVariable("TEMP");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (CultureInfo.CurrentCulture.Name != "ru-RU")
            {
                MessageBox.Show("Региональные настройки не соответствуют ru-RU !" + Environment.NewLine + "Могут быть проблемы при формировании отчетов !", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Thread.CurrentThread.CurrentCulture = new CultureInfo("ru-RU", false);
            }
            string UserName = Environment.UserName;
            if (!File.Exists(ProgPath + "transp.xml"))
            {
                if (MessageBox.Show("Не найден файл настроек transp.xml !" + Environment.NewLine + "Будет произведена первичная настройка !" + Environment.NewLine + "Продолжить ?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    FirstStart fs = new FirstStart();
                    fs.ShowDialog();
                    return;
                }
                if (!File.Exists(ProgPath + "transp.xml")) return;
            }
            XmlTextReader reader = new XmlTextReader(ProgPath + "transp.xml");
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.Name.Equals("server"))
                    {
                        Server = reader.GetAttribute("value");
                    }
                    if (reader.Name.Equals("database"))
                    {
                        DataBase = reader.GetAttribute("value");
                    }
                    if (reader.Name.Equals("dbuser"))
                    {
                        try
                        {
                            string usr = Encoding.UTF8.GetString(Convert.FromBase64String(reader.GetAttribute("value")));
                            DBUser = usr.Substring(2, usr.Length - 4);
                        } catch {
                            MessageBox.Show("Неверно указаны данные доступа к БД !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    if (reader.Name.Equals("dbpass"))
                    {
                        try
                        {
                            string pwd = Encoding.UTF8.GetString(Convert.FromBase64String(reader.GetAttribute("value")));
                            DBPass = pwd.Substring(2, pwd.Length - 4);
                        } catch {
                            MessageBox.Show("Неверно указаны данные доступа к БД !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    if (reader.Name.Equals("username"))
                    {
                        UserName = reader.GetAttribute("value");
                    }
                }
            }
            reader.Close();
            if ((Server == "") || (DataBase == ""))
            {
                MessageBox.Show("В файле настроек неправильно указанна информация о базе данных !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            ClSQL = new ClSQL("server=" + Server + ";uid=" + DBUser + ";pwd=" + DBPass + ";database=" + DataBase);
            if (ClSQL.Error)
            {
                MessageBox.Show("Не удалось подключиться к базе данных !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            ClFunc = new ClFunc();
            string sqlver = ClSQL.SelectCell("select serverproperty('productversion')");
            int p = sqlver.IndexOf(".");
            if (p > -1) sqlver = sqlver.Substring(0, p);
            int sver = ClFunc.StrToInt(sqlver);
            if (sver < 9)
            {
                MessageBox.Show("Версия MS SQL Server должна быть 2005 или выше !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int k = ClSQL.SelectIntCell("select count(*) from users");
            if (k < 1)
            {
                MessageBox.Show("В программе не заведено ни одного пользователя !" + Environment.NewLine + "Заведите хотя бы одного пользователя с правами администратора !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                k = ClSQL.SelectIntCell("select count(*) from users where user_type=1");
                if (k < 1)
                {
                    MessageBox.Show("В программе нет ни одного пользователя с правами администратора !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            FilialPrefix = ClSQL.SelectCell("select curvalue from settings where setkod='filial_prefix'");
            if (ClSQL.Error)
            {
                MessageBox.Show("Отсутствует обязательная настройка: filial_prefix Префикс филиала !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            KolvoDO = ClSQL.SelectIntCell("select curvalue from settings where setkod='kolvo_do'");
            if (ClSQL.Error)
            {
                MessageBox.Show("Отсутствует обязательная настройка: kolvo_do Количество доп.офисов !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            ii = ClSQL.SelectIntCell("select curvalue from settings where setkod='empty_lists_days'");
            if (ii > 0) EmpLstDays = ii;
            DataRow dr = ClSQL.SelectRow("select * from users where user_login='" + UserName + "'");
            if (dr != null)
            {
                if (!int.TryParse(dr["user_id"].ToString(), NumberStyles.Integer, null, out UserID)) UserID = -1;
                if (!int.TryParse(dr["dep_id"].ToString(), NumberStyles.Integer, null, out UserDepID)) UserDepID = -1;
                if (!int.TryParse(dr["user_type"].ToString(), NumberStyles.Integer, null, out UserType)) UserType = -1;
                UserNomDo = dr["nom_do"].ToString();
            }
            if (UserType < 1)
            {
                MessageBox.Show("У Вас нет доступа к программе !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (UserType == 1)
            {
                if (File.Exists(ProgPath + "update.sql"))
                {
                    Update fu = new Update();
                    fu.ShowDialog();
                    if (fu.res == 1)
                    {
                        MessageBox.Show("Обновление завершено. Перезапустите программу.");
                        return;
                    }
                }
            }
            if (UserType != 1)
            {
                string ProgName = Path.GetFileNameWithoutExtension(Application.ExecutablePath);
                if (Process.GetProcessesByName(ProgName).Length > 1)
                {
                    MessageBox.Show("Программа уже запущена !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
                    return;
                }
            }
            string BaseVer = ClSQL.SelectCell("select ver from version");
            if (BaseVer != ProgVer)
            {
                MessageBox.Show("Версия вашей программы не соответствует версии базы !" + Environment.NewLine + "Обратитесь к администратору !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                string strHostName = Dns.GetHostName();
                string str = strHostName + " ; ";
                IPHostEntry iphostentry = Dns.GetHostEntry(strHostName);
                for (int i = 0; i < iphostentry.AddressList.Length; i++) str += iphostentry.AddressList[i] + " ; ";
                ClSQL.ExecuteSQL("update users set info='" + str.Substring(0, str.Length - 3) + "' where user_id=" + UserID.ToString());
            }
            catch (Exception e)
            {
                ClSQL.ExecuteSQL("update users set info='" + e.Message + "' where user_id=" + UserID.ToString());
            }
            ClSQL.ExecuteSQL("update users set last_visit='" + DateTime.Now.ToString(CultureInfo.InvariantCulture) + "' where user_id=" + UserID.ToString());
            Application.Run(new MainForm());
        }
    }
}
