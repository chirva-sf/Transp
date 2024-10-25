using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using Excel;

namespace Transp
{
    class ClExcel
    {
        Excel.Application hExcel;
        Excel.Workbook hBook;
        Excel._Worksheet hTab;
        Excel.Range rng;
        object missing = Missing.Value;

        public bool Open(string FileName)
        {
            hExcel = new Excel.ApplicationClass();
            hExcel.Visible = false;
            hExcel.DisplayAlerts = false;
            hBook = hExcel.Workbooks.Add(FileName);
            bool fl = false;
            for (int i = 1; i <= hBook.Worksheets.Count; i++)
            {
                hTab = (Excel._Worksheet)hBook.Worksheets.get_Item(i);
                if (hTab.Name.Trim() == "Лист1")
                {
                    fl = true; break;
                }
            }
            if (!fl)
            {
                MessageBox.Show("Не найден \"Лист1\" в файле " + FileName + " !");
                hExcel.Quit();
                return false;
            }
            else
            {
                return true;
            }
        }

        public void Show()
        {
            hExcel.DisplayAlerts = true;
            hExcel.Visible = true;
        }

        public void Close()
        {
            hExcel.Quit();
        }

        public string GetStrCell(string RC)
        {
            rng = hTab.get_Range(RC + ":" + RC, missing);
            if (rng.Value2 != null)
            {
                return rng.Value2.ToString().Trim();
            }
            else
            {
                return "";
            }
        }

        public DateTime GetDateCell(string RC)
        {
            DateTime dt = DateTime.MinValue;
            object value = hTab.get_Range(RC + ":" + RC, missing).Value2;
            if (value != null)
            {
                if (value is double)
                {
                    try
                    {
                        dt = DateTime.FromOADate((double)value);
                    }
                    catch
                    {
                        //SaveLog(nom_do, fn, "Ячейка " + RC, "Не верный формат даты !");
                    }
                }
                else
                {
                    if (!DateTime.TryParse((string)value, out dt))
                    {
                        //SaveLog(nom_do, fn, "Ячейка " + RC, "Не верный формат даты !");
                    }
                }
            }
            return dt;
        }

        public Double GetDoubleCell(string RC)
        {
            Double d = 0;
            object value = hTab.get_Range(RC + ":" + RC, missing).Value2;
            if (value != null)
            {
                try
                {
                    d = (double)value;
                }
                catch
                {
                    //SaveLog(nom_do, fn, "Ячейка " + RC, "Не верный формат числа !");
                    d = 0;
                }
            }
            return d;
        }

        public void SetCellValue(string RC, string Value)
        {
            hTab.get_Range(RC, missing).Value2 = Value;
        }

        public void SetCellYXValue(int y, int x, string Value)
        {
            try
            {
                hTab.get_Range(GetColumnLetter(x) + y.ToString(), missing).Value2 = Value;
            }
            catch
            {
            }
        }

        public void SetCellYXValue(int y, int x, double Value)
        {
            try
            {
                hTab.get_Range(GetColumnLetter(x) + y.ToString(), missing).Value2 = Value;
            }
            catch
            {
            }
        }

        public void SetCellYXValue(int y, int x, int Value)
        {
            try
            {
                hTab.get_Range(GetColumnLetter(x) + y.ToString(), missing).Value2 = Value;
            }
            catch
            {
            }
        }

        private string GetColumnLetter(int x)
        {
            string s = "";
            int c = (x - 1) / 26;
            if (c > 0) s = (char)(64 + c) + "";
            s = s + (char)(64 + (x - c * 26));
            return s;
        }

        public void SetByNameValue(string range_name, string Value)
        {
            try
            {
                hTab.get_Range(range_name, missing).Value2 = Value;
            }
            catch
            {
            }
        }

        public void SetByNameValue(string range_name, double Value)
        {
            try
            {
                hTab.get_Range(range_name, missing).Value2 = Value;
            }
            catch
            {
            }
        }

        public void SetByNameValue(string range_name, int Value)
        {
            try
            {
                hTab.get_Range(range_name, missing).Value2 = Value;
            }
            catch
            {
            }
        }

        public void InsertLine(int row, int count)
        {
            for (int i = 0; i < count; i++)
            {
                hTab.get_Range("A" + row.ToString() + ":A" + row.ToString(), missing).EntireRow.Insert(3, true);
            }
        }

        public void RunVBA(string vba_name)
        {
            hExcel.Run(vba_name, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }

        public bool SelectSheet(string sh_name)
        {
            bool fl = false;
            for (int i = 1; i <= hBook.Worksheets.Count; i++)
            {
                hTab = (Excel._Worksheet)hBook.Worksheets.get_Item(i);
                if (hTab.Name.Trim() == sh_name)
                {
                    fl = true; break;
                }
            }
            if (!fl)
            {
                MessageBox.Show("Не найден лист \"" + sh_name + "\" !");
                hExcel.Quit();
                return false;
            }
            else
            {
                return true;
            }
        }

        public string GetSymRange(int nRow, int nCol)
        {
            byte[] bt = new byte[1];
            byte ACol = (byte)nCol;
            byte A = (byte)'A' - 1;
            byte Z = (byte)((byte)'Z' - A);
            byte t = (byte)(ACol / Z);
            byte m = (byte)(ACol % Z);
            if (m == 0) t--;
            string s = "";
            if (t > 0)
            {
                bt[0] = (byte)(A + t);
                s = System.Text.Encoding.ASCII.GetString(bt, 0, bt.Length);
            }
            if (m == 0) t = Z; else t = m;
            bt[0] = (byte)(A + t);
            s += System.Text.Encoding.ASCII.GetString(bt, 0, bt.Length);
            s += nRow.ToString();
            return s;
        }

    }
}
