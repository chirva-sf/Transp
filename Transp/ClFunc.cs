using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;

namespace Transp
{
    class ClFunc
    {
        [DllImport("user32.dll")]
        private static extern int SetForegroundWindow(IntPtr handle);

        ClSQL ClSQL = Program.ClSQL;
        private int[] ustparr = new int[30];
        private int[] ftarr = new int[30];
        private int[] deparr = new int[300];
        private int[] drvarr = new int[300];
        private int[] cararr = new int[300];
        private int[] fcarr = new int[3000];

        /* ******** Int ******** */

        public bool ChkIntVal(string s)
        {
            int intVal;
            if (s == "") return true; else return Int32.TryParse(NumStr(s), (NumberStyles.Integer | NumberStyles.AllowThousands), null, out intVal);
        }

        public int StrToInt(string s)
        {
            int intVal;
            Int32.TryParse(NumStr(s), (NumberStyles.Integer | NumberStyles.AllowThousands), null, out intVal);
            return intVal;
        }

        public bool ChkIntVal(TextBox t)
        {
            return ChkIntVal(t.Text);
        }

        public int StrToInt(TextBox t)
        {
            return StrToInt(t.Text);
        }

        /* ******** Double ******** */

        public bool ChkDoubleVal(string s)
        {
            double dVal;
            if (s == "") return true; else return Double.TryParse(NumStr(s), (NumberStyles.Float | NumberStyles.AllowThousands), new CultureInfo("en-US", false).NumberFormat, out dVal);
        }

        public double StrToDouble(string s)
        {
            double dVal;
            Double.TryParse(NumStr(s), (NumberStyles.Float | NumberStyles.AllowThousands), new CultureInfo("en-US", false).NumberFormat, out dVal);
            return dVal;
        }

        public bool ChkDoubleVal(TextBox t)
        {
            return ChkDoubleVal(t.Text);
        }

        public double StrToDouble(TextBox t)
        {
            return StrToDouble(t.Text);
        }

        /* ******** Decimal ******** */

        public bool ChkDecimalVal(string s)
        {
            decimal dVal;
            if (s == "") return true; else return Decimal.TryParse(NumStr(s), (NumberStyles.Float | NumberStyles.AllowThousands), new CultureInfo("en-US", false).NumberFormat, out dVal);
        }

        public decimal StrToDecimal(string s)
        {
            decimal dVal;
            Decimal.TryParse(NumStr(s), (NumberStyles.Float | NumberStyles.AllowThousands), new CultureInfo("en-US", false).NumberFormat, out dVal);
            return dVal;
        }

        public decimal StrToDecimal(TextBox t)
        {
            return StrToDecimal(t.Text);
        }

        public bool ChkDecimalVal(TextBox t)
        {
            return ChkDecimalVal(t.Text);
        }

        /* ******** Date ******** */

        public string DateToStr(DateTime d)
        {
            return d.Date.ToString("d", new CultureInfo("en-US"));
        }

        public string DateToStrR(DateTime d)
        {
            return d.Date.ToString("d", new CultureInfo("ru-RU"));
        }

        public DateTime StrToDate(string s)
        {
            return DateTime.Parse(s, new CultureInfo("en-US")); 
        }

        public string TimeFromDateTime(string s)
        {
            string r = "";
            int p = s.IndexOf(" ");
            if (p != -1)
            {
                s = s.Substring(p + 1);
                if (s.Substring(1, 1) == ":") s = "0" + s;
                r = s.Substring(0, 5);
            }
            if (r == "00:00") r = "";
            return r;
        }

        public string DateFromDateTime(string s)
        {
            string r = "";
            int p = s.IndexOf(" ");
            if (p != -1)
            {
                r = s.Substring(0, p);
            }
            return r;
        }

        public string StrDateToSQL(string s) 
        {
            return s == "" ? "null" : s.Substring(3, 2) + "." + s.Substring(0, 2) + "." + s.Substring(6);
        }

        /* ******** additionally ******** */

        public string NumStr(TextBox t)
        {
            return NumStr(t.Text);
        }

        public string NumStr(string s)
        {
            if (s == "") s = "0";
            while (s.IndexOf(",") > -1) s = s.Replace(",", ".");
            while (s.IndexOf(" ") > -1) s = s.Replace(" ", "");
            while (s.IndexOf("\u00A0") > -1) s = s.Replace("\u00A0", "");
            while (s.IndexOf("\u00D0") > -1) s = s.Replace("\u00D0", "");
            return s;
        }

        public string Empty(string s)
        {
            if (s == "0") return ""; else return s;
        }

        public string GetFIO(string s)
        {
            int i = -1, j = -1;
            i = s.IndexOf(".");
            if (i < s.Length - 1) j = s.Substring(i + 1).IndexOf(".");
            if (i > 0 && j > 0)
            {
                return s;
            }
            else
            {
                string r = "";
                string p = s.Trim() + " ";
                i = p.IndexOf(" ");
                r = p.Substring(0, i + 1);
                if (p.Length > i + 1)
                {
                    p = p.Substring(p.IndexOf(" ") + 1);
                    r += p.Substring(0, 1) + ".";
                }
                if (p.Length > p.IndexOf(" ") + 1)
                {
                    p = p.Substring(p.IndexOf(" ") + 1);
                    r += p.Substring(0, 1) + ".";
                }
                return r.Trim();
            }
        }

        public void SetActiveWord()
        {
            Process[] processArray = Process.GetProcessesByName("WinWord");
            Process word = processArray[0];
            SetForegroundWindow(word.MainWindowHandle);
        }

        public void SetActiveExcel()
        {
            Process[] processArray = Process.GetProcessesByName("Excel");
            Process word = processArray[0];
            SetForegroundWindow(word.MainWindowHandle);
        }

        // ***** Other functions *****

        public string LoadUserParam(string ParamName, int user_id = 0)
        {
            return ClSQL.SelectCell("select pvalue from usrparams where user_id=" + (user_id > 0 ? user_id : Program.UserID).ToString() + " and name='" + ParamName + "'");
        }

        public void SaveUserParam(string ParamName, string ParamValue, int user_id = 0)
        {
            user_id = user_id > 0 ? user_id : Program.UserID;
            int k = ClSQL.SelectIntCell("select count(*) from usrparams where user_id=" + user_id.ToString() + " and name='" + ParamName + "'");
            if (k > 0)
            {
                ClSQL.ExecuteSQL("update usrparams set pvalue='" + ParamValue + "' where user_id=" + user_id.ToString() + " and name='" + ParamName + "'");
            }
            else
            {
                ClSQL.ExecuteSQL("insert into usrparams values (" + user_id.ToString() + ",'" + ParamName + "','" + ParamValue + "')");
            }
        }

        // *****

        public int GetMaxNom(string pole, string tname)
        {
            int intVal;
            int maxnom = 0;
            string s,t;
            int p = pole.IndexOf("_");
            DataTable dt = ClSQL.SelectSQL("select top 100 " + pole + " from " + tname + " order by " + pole.Substring(0,p) + "_id desc");
            foreach (DataRow dr in dt.Rows)
            {
                t = dr[0].ToString(); s = "";
                for (int i = 0; i < t.Length; i++)
                {
                    if ("0123456789".IndexOf(t.Substring(i, 1)) > -1)
                    {
                        s += t.Substring(i, 1);
                    }
                }
                if (Int32.TryParse(s, (NumberStyles.Integer | NumberStyles.AllowThousands), null, out intVal))
                {
                    if (intVal > maxnom) maxnom = intVal;
                }
            }
            return maxnom;
        }

        // *****

        public void UpdateNomDO(ComboBox cb, string SelNomDO)
        {
            cb.Items.Clear();
            for (int i = 0; i <= Program.KolvoDO; i++)
            {
                string s = i.ToString();
                if (s.Length < 2) s = "0" + s;
                cb.Items.Add(Program.FilialPrefix + s);
            }
            int ti = -1;
            if (SelNomDO.Length > 1)
            {
                if (SelNomDO.Substring(0, 2) == Program.FilialPrefix)
                {
                    ti = int.Parse(SelNomDO) - int.Parse(Program.FilialPrefix) * 100;
                }
            }
            cb.SelectedIndex = ti;
        }

        public string GetNomDO(ComboBox cb)
        {
            return cb.Text;
        }

        // *****

        public void UpdateFuelType(ComboBox cb, int si)
        {
            cb.Items.Clear();
            DataTable dt = ClSQL.SelectSQL("select * from fuel_types order by ft_name");
            int ti = -1;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ftarr[i] = (int)dt.Rows[i]["ft_id"];
                if (ftarr[i] == si) ti = i;
                cb.Items.Add(dt.Rows[i]["ft_name"].ToString());
            }
            cb.SelectedIndex = ti;
        }

        public string GetFuelType(ComboBox cb)
        {
            return cb.SelectedIndex == -1 ? "-1" : ftarr[cb.SelectedIndex].ToString();
        }

        // *****

        public void UpdateDepartments(ComboBox cb, int SelDepid)
        {
            cb.Items.Clear();
            DataTable dt = ClSQL.SelectSQL("select dep_id,nom_do,name from departments order by nom_do,name");
            int ti = -1;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                deparr[i] = (int)dt.Rows[i]["dep_id"];
                if (deparr[i] == SelDepid) ti = i;
                cb.Items.Add(dt.Rows[i]["name"].ToString());
            }
            cb.SelectedIndex = ti;
        }

        public void UpdateDepartmentsRef(ComboBox cb)
        {
            int cur_id = deparr[cb.SelectedIndex];
            Departments_List frm = new Departments_List();
            frm.StartMode = 1;
            frm.ShowDialog();
            if (frm.SelDepid > 0)
            {
                UpdateDepartments(cb, frm.SelDepid);
            }
            else
            {
                UpdateDepartments(cb, cur_id);
            }
        }

        public string GetDepartment(ComboBox cb)
        {
            return cb.SelectedIndex == -1 ? "-1" : deparr[cb.SelectedIndex].ToString();
        }

        // *****

        public void UpdateUsersTypes(ComboBox cb, int SelUstid)
        {
            cb.Items.Clear();
            DataTable dt = ClSQL.SelectSQL("select * from usrtypes");
            int ti = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ustparr[i] = (int)dt.Rows[i]["ust_id"];
                if (ustparr[i] == SelUstid) ti = i;
                cb.Items.Add(dt.Rows[i]["name"].ToString());
            }
            cb.SelectedIndex = ti;
        }

        public string GetUsersType(ComboBox cb)
        {
            return cb.SelectedIndex == -1 ? "-1" : ustparr[cb.SelectedIndex].ToString();
        }

        // *****

        public void UpdateDrivers(ComboBox cb, int si)
        {
            cb.Items.Clear();
            string strSQL = "select * from drivers where status<>1 ";
            if (Program.UserType > 3) strSQL += "and nom_do='" + Program.UserNomDo + "' ";
            strSQL += "order by nom_do, fio";
            DataTable dt = ClSQL.SelectSQL(strSQL);
            int ti = -1;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                drvarr[i] = (int)dt.Rows[i]["drv_id"];
                if (drvarr[i] == si) ti = i;
                cb.Items.Add(dt.Rows[i]["nom_do"].ToString() + "  " + dt.Rows[i]["fio"].ToString());
            }
            cb.SelectedIndex = ti;
        }

        public string GetDriver(ComboBox cb)
        {
            return cb.SelectedIndex == -1 ? "-1" : drvarr[cb.SelectedIndex].ToString();
        }

        // *****

        public void UpdateTranspSr(ComboBox cb, int si)
        {
            cb.Items.Clear();
            DataTable dt = ClSQL.SelectSQL("select * from cars order by nom_do, gosnomer");
            int ti = -1;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cararr[i] = (int)dt.Rows[i]["car_id"];
                if (cararr[i] == si) ti = i;
                cb.Items.Add(dt.Rows[i]["nom_do"].ToString() + "  -  " + dt.Rows[i]["marka"].ToString() + "   " + dt.Rows[i]["gosnomer"].ToString());
            }
            cb.SelectedIndex = ti;
        }

        public string GetTranspSr(ComboBox cb)
        {
            return cb.SelectedIndex == -1 ? "-1" : cararr[cb.SelectedIndex].ToString();
        }

        // *****

        public void UpdateFuelCards(ComboBox cb, int si)
        {
            cb.Items.Clear();
            string s = "select * from fuel_cards ";
            if (Program.UserType > 3)
            {
                s += "where nom_do='" + Program.UserNomDo + "' ";
                if (si > 0) s += "or fc_id=" + si.ToString();
            }
            s += "order by fc_nomer";
            DataTable dt = ClSQL.SelectSQL(s);
            int ti = -1;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                fcarr[i] = (int)dt.Rows[i]["fc_id"];
                if (fcarr[i] == si) ti = i;
                cb.Items.Add(dt.Rows[i]["fc_nomer"].ToString());
            }
            cb.SelectedIndex = ti;
        }

        public string GetFuelCard(ComboBox cb)
        {
            return cb.SelectedIndex == -1 ? "-1" : fcarr[cb.SelectedIndex].ToString();
        }

        public int GetFuelCardInt(ComboBox cb)
        {
            return cb.SelectedIndex == -1 ? -1 : fcarr[cb.SelectedIndex];
        }

        // ************

        public double GetRasxNorm(int car_id, DateTime d, int mileage, int mtype)
        {
            return ClSQL.SelectDoubleCell("select dbo.getRasxNorm (" + car_id.ToString() + ",'" + DateToStr(d) + "'," + mileage.ToString() + "," + mtype.ToString() + ")");
        }

        public double GetRasxBase(int car_id, DateTime d)
        {
            return ClSQL.SelectDoubleCell("select top 1 base_rasxod from pr_rasxod where car_id=" + car_id.ToString() + " order by beg_date desc");
        }

        public int GetBegMileage(int doc_id, int car_id, DateTime d)
        {
            return ClSQL.SelectIntCell("select dbo.getBegMileage (" + doc_id.ToString() + "," + car_id.ToString() + ",'" + DateToStr(d) + "')");
        }

        public double GetBegFuel(int doc_id, int car_id, DateTime d)
        {
            return ClSQL.SelectDoubleCell("select dbo.getBegFuel (" + doc_id.ToString() + "," + car_id.ToString() + ",'" + DateToStr(d) + "')");
        }
    }
}
