using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Threading;

namespace Transp
{
    public partial class PutList : Form
    {
        ClSQL ClSQL = Program.ClSQL;
        ClFunc fn = Program.ClFunc;
        BindingSource BSpl = new BindingSource();
        DataTable DTpl = new DataTable();
        public bool EditMode = false;
        public bool ReadOnly = false;
        public int doc_id = -1;
        private bool fl_load = false;
        private bool Changed = false;
        private bool FEnabled = true;
        private string nom_do = "";
        private int dep_id = 0;
        private int car_id = 0;
        private int fc_id = 0;
        private int beg_mileage = 0;
        private double beg_fuel = 0;
        private double end_fuel = 0;
        private string ft_name = "";
        private double rasx_gorod = 0;
        private double rasx_trassa = 0;
        private double rasx_base = 0;
        private decimal rasx_fuel = 0;
        private double fuel_ekonom = 0;
        private string dispatcher = "";
        private string mechanic = "";

        public PutList()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ShowInfoDriver()
        {
            string strSQL;
            DateTime dt1 = DateTime.MinValue;
            DateTime dt2 = DateTime.MinValue;
            DataRow dr;
            string ds = fn.DateToStr(dateTimePicker1.Value);
            string dsr = fn.DateToStrR(dateTimePicker1.Value);
            string drv_id = fn.GetDriver(comboBox3);
            if (drv_id == "-1") return;
            if (fl_load) return;

            strSQL = "select dr.*, dp.name as dpname from drivers dr, departments dp where dr.dep_id=dp.dep_id and dr.drv_id=" + drv_id;
            dr = ClSQL.SelectRow(strSQL);
            if (dr != null)
            {
                if (dr["nom_do"].ToString() == "") MessageBox.Show("Не указан Филиал/ДО водителя !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (dr["tab_no"].ToString() == "") MessageBox.Show("Не указан таб.номер водителя !", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (dr["udostov"].ToString() == "") MessageBox.Show("Не указано удостоверение водителя !", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox2.Text = dr["nom_do"].ToString();
                textBox3.Text = dr["dpname"].ToString();
                textBox4.Text = dr["udostov"].ToString();
                nom_do = dr["nom_do"].ToString();
                dep_id = (int)dr["dep_id"];
            }
            else
            {
                textBox2.Text = ""; textBox3.Text = ""; textBox4.Text = "";
                nom_do = ""; dep_id = -1;
                MessageBox.Show("Не нашел водителя !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            strSQL = "select top 1 fc.fc_id, ft.ft_name from pr_fcdrv p, fuel_cards fc, fuel_types ft ";
            strSQL += "where p.fc_id=fc.fc_id and fc.ft_id=ft.ft_id and p.drv_id=" + drv_id + " and p.beg_date <= '" + ds + "' ";
            strSQL += "order by p.beg_date desc, p.pr_id desc";
            dr = ClSQL.SelectRow(strSQL);
            if (dr != null)
            {
                fc_id = (int)dr["fc_id"];
                ft_name = dr["ft_name"].ToString();
                fn.UpdateFuelCards(comboBox2, (int)dr["fc_id"]);
            }
            else
            {
                fn.UpdateFuelCards(comboBox2, 0);
                textBox12.Text = "";
                MessageBox.Show("На " + dsr + " водителю не выдана топливная карта !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            strSQL = "select top 1 c.car_id, c.marka, c.gosnomer, c.garnomer from pr_drvcar p, cars c ";
            strSQL += "where p.car_id=c.car_id and p.drv_id=" + drv_id + " and p.beg_date <= '" + ds + "' ";
            strSQL += "order by p.beg_date desc, p.pr_id desc";
            dr = ClSQL.SelectRow(strSQL);
            if (dr != null)
            {
                if (dr["garnomer"].ToString() == "") MessageBox.Show("Не указан гаражный номер автомобиля !", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information);
                car_id = (int)dr["car_id"];
                textBox5.Text = dr["marka"].ToString();
                textBox6.Text = dr["gosnomer"].ToString();
            }
            else
            {
                textBox5.Text = "";
                textBox6.Text = "";
                MessageBox.Show("На " + dsr + " за водителем не закреплен автомобль !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            strSQL = "select top 1 dispatcher, mechanic from pr_signs ";
            strSQL += "where nom_do='" + nom_do + "' and beg_date <= '" + ds + "' ";
            strSQL += "order by beg_date desc";
            dr = ClSQL.SelectRow(strSQL);
            if (dr != null)
            {
                dispatcher = dr["dispatcher"].ToString();
                mechanic = dr["mechanic"].ToString();
                textBox18.Text = dispatcher;
                textBox19.Text = mechanic;
            }
            else
            {
                textBox18.Text = "";
                textBox19.Text = "";
                MessageBox.Show("На " + dsr + " не назначен диспетчер, механик !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            beg_mileage = fn.GetBegMileage(doc_id, car_id, dateTimePicker1.Value);
            textBox7.Text = beg_mileage.ToString();
            if (beg_mileage == -1) { MessageBox.Show("На " + dsr + " нет начальных данных спидометра !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); textBox7.Text = ""; }

            beg_fuel = fn.GetBegFuel(doc_id, car_id, dateTimePicker1.Value);
            textBox9.Text = Decimal.Round((decimal)beg_fuel, 1, MidpointRounding.AwayFromZero).ToString();
            if (beg_fuel == -1) { MessageBox.Show("На " + dsr + " нет начальных данных топлива !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); textBox9.Text = ""; }

            if (dateTimePicker1.Value.Date < new DateTime(2017, 09, 04))
            {
                rasx_gorod = fn.GetRasxNorm(car_id, dateTimePicker1.Value, beg_mileage, 1);
                textBox13.Text = fn.Empty(rasx_gorod.ToString());
                if (rasx_gorod == -1) { MessageBox.Show("На " + dsr + " не указан расход топлива по норме !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); textBox13.Text = ""; rasx_gorod = 0;  }
            }
            else
            {
                rasx_gorod = fn.GetRasxNorm(car_id, dateTimePicker1.Value, beg_mileage, 1);
                textBox13.Text = fn.Empty(rasx_gorod.ToString());
                if (rasx_gorod == -1) { MessageBox.Show("На " + dsr + " не указан расход топлива \"Город\" по норме !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); textBox13.Text = ""; rasx_gorod = 0;  }

                rasx_trassa = fn.GetRasxNorm(car_id, dateTimePicker1.Value, beg_mileage, 2);
                textBox25.Text = fn.Empty(rasx_trassa.ToString());
                if (rasx_trassa == -1) { MessageBox.Show("На " + dsr + " не указан расход топлива \"Трасса\" по норме !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); textBox25.Text = ""; rasx_trassa = 0; }
            }
            rasx_base = fn.GetRasxBase(car_id, dateTimePicker1.Value);
        }

        private void EmptyRaschet()
        {
            textBox10.Text = ""; textBox14.Text = "";  textBox15.Text = "";
            textBox16.Text = ""; textBox17.Text = "";
            end_fuel = 0; rasx_fuel = 0; fuel_ekonom = 0;
        }

        private void ShowInfoRasch()
        {
            bool fldt = true; string ss = ClSQL.SelectCell("select curvalue from settings where setkod='enable_downtime'");
            if (ss != "Да" && ss != "да" && ss != "ДА") fldt = false;
            string drv_id = fn.GetDriver(comboBox3);
            if (drv_id != "-1")
            {
                /* Расчет расхода и остатка топлива */
                int totkm = fn.StrToInt(textBox8) - beg_mileage;
                if (totkm < 0) { EmptyRaschet(); return; }
                if (dateTimePicker1.Value.Date < new DateTime(2013, 09, 16))
                {
                    rasx_fuel = Decimal.Round((totkm * (decimal)rasx_gorod) / 100, 1, MidpointRounding.AwayFromZero);
                    if (fldt) rasx_fuel = rasx_fuel + (decimal)rasx_base * fn.StrToDecimal(textBox11);
                    end_fuel = (double)((decimal)beg_fuel + Decimal.Round(fn.StrToDecimal(textBox12), 1, MidpointRounding.AwayFromZero) - rasx_fuel); 
                }
                else if (dateTimePicker1.Value.Date < new DateTime(2017, 09, 04))
                {
                    rasx_fuel = (totkm * (decimal)rasx_gorod) / 100;
                    if (fldt) rasx_fuel = rasx_fuel + (decimal)rasx_base * fn.StrToDecimal(textBox11);
                    end_fuel = (double)((decimal)beg_fuel + fn.StrToDecimal(textBox12) - rasx_fuel);
                }
                else
                {
                    int mileage_gorod = fn.StrToInt(textBox23.Text);
                    int mileage_trassa = fn.StrToInt(textBox24.Text);
                    if (mileage_gorod > 0 || mileage_trassa > 0)
                    {
                        rasx_fuel = ((mileage_gorod * (decimal)rasx_gorod) / 100) + ((mileage_trassa * (decimal)rasx_trassa) / 100);
                        if (fldt) rasx_fuel = rasx_fuel + (decimal)rasx_base * fn.StrToDecimal(textBox11);
                        end_fuel = (double)((decimal)beg_fuel + fn.StrToDecimal(textBox12) - rasx_fuel);
                    }
                    else
                    {
                        rasx_fuel = (totkm * (decimal)rasx_gorod) / 100;
                        if (fldt) rasx_fuel = rasx_fuel + (decimal)rasx_base * fn.StrToDecimal(textBox11);
                        end_fuel = (double)((decimal)beg_fuel + fn.StrToDecimal(textBox12) - rasx_fuel);
                    }
                }
                fuel_ekonom = 0;
                textBox10.Text = Decimal.Round((decimal)end_fuel, 1, MidpointRounding.AwayFromZero).ToString();
                textBox14.Text = Decimal.Round(rasx_fuel, 1, MidpointRounding.AwayFromZero).ToString();
                textBox15.Text = totkm.ToString();
                textBox16.Text = fuel_ekonom < 0 ? (-fuel_ekonom).ToString() : "0";
                textBox17.Text = fuel_ekonom > 0 ? fuel_ekonom.ToString() : "0";
            }
            else
            {
                EmptyRaschet();
            }
        }

        private void MileageSum()
        {
            decimal sum1 = 0, sum2 = 0, total = 0;
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                total += fn.StrToDecimal(dataGridView1.Rows[i].Cells[6].Value.ToString());
                if (dataGridView1.Rows[i].Cells[7].Value.ToString() == "Город")
                {
                    sum1 += fn.StrToDecimal(dataGridView1.Rows[i].Cells[6].Value.ToString());
                }
                else if (dataGridView1.Rows[i].Cells[7].Value.ToString() == "Трасса")
                {
                    sum2 += fn.StrToDecimal(dataGridView1.Rows[i].Cells[6].Value.ToString());
                }
            }
            sum1 = Decimal.Round(sum1, 0, MidpointRounding.AwayFromZero);
            textBox23.Text = sum1.ToString();
            sum2 = Decimal.Round(sum2, 0, MidpointRounding.AwayFromZero);
            textBox24.Text = sum2.ToString();
            total = Decimal.Round(total, 0, MidpointRounding.AwayFromZero);
            textBox20.Text = total.ToString();
            if (dateTimePicker1.Value.Date >= new DateTime(2017, 09, 04) && !fl_load)
            {
                textBox21.Text = sum1.ToString();
                textBox22.Text = sum2.ToString();
                ShowInfoRasch();
            }
        }

        /* ****************** Check errors ****************** */

        private bool HasErrors()
        {
            bool fl1_er, fl2_er;
            bool flag = false;
            string s = "", ss = "";
            int intVal;
            decimal decVal;
            decimal decVal2;
            decimal totkm = 0;
            if (textBox1.Text == "")
            {
                MessageBox.Show("Заполните номер документа !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (comboBox3.SelectedIndex < 0)
            {
                MessageBox.Show("Выберите водителя !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (nom_do == "")
            {
                MessageBox.Show("Не указан Филиал/ДО водителя !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (car_id <= 0)
            {
                MessageBox.Show("Не указан автомобиль !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (fc_id <= 0)
            {
                MessageBox.Show("Не указана топливная карта !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            else if (rasx_gorod <= 0)
            {
                if (dateTimePicker1.Value.Date >= new DateTime(2017, 09, 04))
                {
                    MessageBox.Show("Не указан расход по норме \"Город\" !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
                }
                else
                {
                    MessageBox.Show("Не указан расход по норме !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
                }
            }
            else if (dateTimePicker1.Value.Date >= new DateTime(2017, 09, 04) && rasx_trassa <= 0)
            {
                MessageBox.Show("Не указан расход по норме \"Трасса\" !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); return true;
            }
            for (int i = 0; i < DTpl.Rows.Count; i++)
            {
                DataRow dr = DTpl.Rows[i];
                if (dr[0].ToString() != "" || dr[1].ToString() != "" || dr[2].ToString() != "" || dr[3].ToString() != "" ||
                    dr[4].ToString() != "" || dr[5].ToString() != "" || dr[6].ToString() != "")
                {
                    flag = true;

                    fl1_er = !Int32.TryParse(fn.NumStr(dr[2].ToString()), (NumberStyles.Integer | NumberStyles.AllowThousands), null, out intVal);
                    if (!fl1_er && (intVal < 0 || intVal > 23)) fl1_er = true;
                    fl2_er = !Int32.TryParse(fn.NumStr(dr[3].ToString()), (NumberStyles.Integer | NumberStyles.AllowThousands), null, out intVal);
                    if (!fl2_er && (intVal < 0 || intVal > 59)) fl2_er = true;
                    if (fl1_er || fl2_er) s += (s == "" ? "" : Environment.NewLine) + "Не верно заполнено время выезда на оборотной стороне !";

                    fl1_er = !Int32.TryParse(fn.NumStr(dr[4].ToString()), (NumberStyles.Integer | NumberStyles.AllowThousands), null, out intVal);
                    if (!fl1_er && (intVal < 0 || intVal > 23)) fl1_er = true;
                    fl2_er = !Int32.TryParse(fn.NumStr(dr[5].ToString()), (NumberStyles.Integer | NumberStyles.AllowThousands), null, out intVal);
                    if (!fl2_er && (intVal < 0 || intVal > 59)) fl2_er = true;
                    if (fl1_er || fl2_er) s += (s == "" ? "" : Environment.NewLine) + "Не верно заполнено время возврата на оборотной стороне !";

                    fl1_er = !Decimal.TryParse(fn.NumStr(dr[6].ToString()), (NumberStyles.Float | NumberStyles.AllowThousands), new CultureInfo("en-US", false).NumberFormat, out decVal);
                    if (!fl1_er && decVal < 0) fl1_er = true; else totkm += decVal;
                    if (fl1_er) s += (s == "" ? "" : Environment.NewLine) + "Не верно заполнен пробег на оборотной стороне !";

                    if (dr[7].ToString() == "" && dateTimePicker1.Value.Date >= new DateTime(2017, 09, 04))
                    {
                        s += (s == "" ? "" : Environment.NewLine) + "Не указан Вид пробега на обратной стороне !";
                    }
                }
                if (s != "") break;
            }
            if (s == "" && flag)
            {
                totkm = Decimal.Round(totkm, 0, MidpointRounding.AwayFromZero);
                decVal2 = fn.StrToDecimal(textBox15.Text);
                if (textBox8.Text != "" && textBox8.Text != "0" && decVal2 != totkm) s = "Не совпадает пробег на обратной стороне путевого листа с разницей показаний спидометра !";
            }
            int id = ClSQL.SelectIntCell("select pl_id from put_lists where car_id=" + car_id.ToString() + " and pl_date='" + fn.DateToStr(dateTimePicker1.Value) + "'");
            if ((EditMode && id > 0 && id != doc_id) || (!EditMode && id > 0))
            {
                if (Program.UserType < 3)
                {
                    MessageBox.Show(fn.DateFromDateTime(dateTimePicker1.Value.ToString()) + " уже есть путевой лист !", "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
                }
                else
                {
                    s = fn.DateFromDateTime(dateTimePicker1.Value.ToString()) + " уже есть путевой лист !" + Environment.NewLine + "За один день на автомобиль может быть только один путевой лист !";
                }
            }
            if (s != "") 
            { 
                MessageBox.Show(s, "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
                return true; 
            }
            if (Program.EmpLstDays > 0)
            {
                ss = "select pl_nom, pl_date from put_lists where car_id=" + car_id.ToString() + " and pl_date<'" + fn.DateToStr(dateTimePicker1.Value) + "' and (end_mileage=0 or CONVERT(varchar(8), beg_time, 108)='00:00:00' or CONVERT(varchar(8), end_time, 108)='00:00:00')";
                if (EditMode) ss += " and pl_id<>" + doc_id;
                DataTable dt = ClSQL.SelectSQL(ss);
                if (dt.Rows.Count > Program.EmpLstDays)
                {
                    s = "По автомобилю " + textBox5.Text + " " + textBox6.Text + " есть более " + Program.EmpLstDays.ToString() + " не заполненных или не до конца заполненных документа !";
                    foreach (DataRow dr in dt.Rows)
                    {
                        s += Environment.NewLine + dr["pl_nom"].ToString() + " от " + dr["pl_date"].ToString().Substring(0, 10);
                    }
                    MessageBox.Show(s, "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return true;
                }
                else if (dt.Rows.Count > 2)
                {
                    s = "По автомобилю " + textBox5.Text + " " + textBox6.Text + " есть не заполненные или не до конца заполненные документы !";
                    foreach (DataRow dr in dt.Rows)
                    {
                        s += Environment.NewLine + dr["pl_nom"].ToString() + " от " + dr["pl_date"].ToString().Substring(0, 10);
                    }
                    MessageBox.Show(s, "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                ss = "select top(50) pl_id, pl_nom, pl_date from put_lists where car_id=" + car_id.ToString() + " and end_mileage - beg_mileage > 0 and pl_date < '" + fn.DateToStr(dateTimePicker1.Value) + "' and pl_date >= '9/4/2017'";
                if (EditMode) ss += " and pl_id<>" + doc_id;
                ss += " order by pl_date";
                DataTable dtt = ClSQL.SelectSQL(ss);
                int k = 0;  s = "";
                foreach (DataRow dr in dtt.Rows)
                {
                    DataRow tdr = ClSQL.SelectRow("select * from (select sum(mileage) as mileage_gorod from put_lists_t where pl_id=" + dr["pl_id"] + " and mtype=1) t1, (select sum(mileage) as mileage_trassa from put_lists_t where pl_id=" + dr["pl_id"] + " and mtype=2) t2");
                    decimal mileage_gorod = tdr["mileage_gorod"].ToString() == "" ? 0 : (decimal)(double)tdr["mileage_gorod"];
                    decimal mileage_trassa = tdr["mileage_trassa"].ToString() == "" ? 0 : (decimal)(double)tdr["mileage_trassa"];
                    if (mileage_gorod == 0 && mileage_trassa == 0)
                    {
                        k++; s += Environment.NewLine + dr["pl_nom"].ToString() + " от " + dr["pl_date"].ToString().Substring(0, 10);
                        if (k > Program.EmpLstDays) break;
                    }
                }
                if (k > Program.EmpLstDays)
                {
                    MessageBox.Show("По автомобилю " + textBox5.Text + " " + textBox6.Text + " не заполнена оборотная сторона более чем в " + Program.EmpLstDays.ToString() + " путевых листах !" + s, "Внимание !!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return true;
                }
                else if (k > 2)
                {
                    MessageBox.Show("По автомобилю " + textBox5.Text + " " + textBox6.Text + " не заполнена оборотная сторона в предыдущих путевых листах !" + s, "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
            return false;
        }

        /* ****************** Save docum ****************** */

        private void SaveOborot()
        {
            string ts;
            int ii = 0;
            string strSQL = "delete from put_lists_t where pl_id=" + doc_id.ToString();
            ClSQL.ExecuteSQL(strSQL);
            for (int i = 0; i < DTpl.Rows.Count; i++)
            {
                DataRow dr = DTpl.Rows[i];
                if (dr[0].ToString() != "" || dr[1].ToString() != "" || dr[2].ToString() != "" || dr[3].ToString() != "" ||
                    dr[4].ToString() != "" || dr[5].ToString() != "" || dr[6].ToString() != "")
                {
                    ii++;
                    strSQL = "insert into put_lists_t values (";
                    strSQL += doc_id.ToString() + ",";
                    strSQL += ii.ToString() + ",";
                    strSQL += "'" + dr[0].ToString() + "',";
                    strSQL += "'" + dr[1].ToString() + "',";
                    ts = fn.NumStr(dr[2].ToString()) + ":" + fn.NumStr(dr[3].ToString());
                    strSQL += "'" + fn.DateToStr(dateTimePicker1.Value) + " " + ts + "',";
                    ts = fn.NumStr(dr[4].ToString()) + ":" + fn.NumStr(dr[5].ToString());
                    strSQL += "'" + fn.DateToStr(dateTimePicker1.Value) + " " + ts + "',";
                    strSQL += fn.NumStr(dr[6].ToString()) + ",";
                    strSQL += (dr[7].ToString() == "Город" ? "1" : (dr[7].ToString() == "Трасса" ? "2" : "0")) + ")";
                    ClSQL.ExecuteSQL(strSQL);
                }
            }
        }

        private void SavePutList()
        {
            string strSQL = "";
            if (EditMode)
            {
                strSQL = "update put_lists set ";
                strSQL += "nom_do='" + nom_do + "',";
                strSQL += "dep_id=" + dep_id.ToString() + ",";
                strSQL += "pl_nom='" + textBox1.Text + "',";
                strSQL += "pl_date='" + fn.DateToStr(dateTimePicker1.Value) + "',";
                strSQL += "car_id=" + car_id.ToString() + ",";
                strSQL += "drv_id=" + fn.GetDriver(comboBox3) + ",";
                strSQL += "beg_time='" + fn.DateToStr(dateTimePicker1.Value) + (maskedTextBox1.Text.Trim() == ":" ? "" : " " + maskedTextBox1.Text) + "',";
                strSQL += "end_time='" + fn.DateToStr(dateTimePicker1.Value) + (maskedTextBox2.Text.Trim() == ":" ? "" : " " + maskedTextBox2.Text) + "',";
                strSQL += "beg_mileage=" + beg_mileage.ToString() + ",";
                strSQL += "end_mileage=" + fn.NumStr(textBox8) + ",";
                strSQL += "beg_fuel=" + fn.NumStr(beg_fuel.ToString()) + ",";
                strSQL += "end_fuel=" + fn.NumStr(end_fuel.ToString()) + ",";
                strSQL += "fc_id=" + fc_id.ToString() + ",";
                strSQL += "fuel_in=" + fn.NumStr(textBox12) + ",";
                strSQL += "downtime=" + fn.NumStr(textBox11) + ",";
                strSQL += "rasx_base=" + fn.NumStr(rasx_base.ToString()) + ",";
                strSQL += "rasx_gorod=" + fn.NumStr(rasx_gorod.ToString()) + ",";
                strSQL += "rasx_trassa=" + fn.NumStr(rasx_trassa.ToString()) + ",";
                strSQL += "rasx_fuel=" + fn.NumStr(rasx_fuel.ToString()) + ",";
                strSQL += "fuel_ekonom=" + fn.NumStr(fuel_ekonom.ToString()) + ",";
                strSQL += "dispatcher='" + dispatcher + "',";
                strSQL += "mechanic='" + mechanic + "',";
                strSQL += "status=" + (checkBox1.Checked ? "1" : "0") + " ";
                strSQL += "where pl_id=" + doc_id.ToString();
                ClSQL.ExecuteSQL(strSQL);
            }
            else
            {
                strSQL = "insert into put_lists (nom_do,dep_id,pl_nom,pl_date,car_id,drv_id,beg_time,end_time,beg_mileage,end_mileage,";
                strSQL += "beg_fuel,end_fuel,fc_id,fuel_in,downtime,rasx_base,rasx_gorod,rasx_trassa,rasx_fuel,fuel_ekonom,dispatcher,mechanic,status) values (";
                strSQL += "'" + nom_do + "',";
                strSQL += dep_id.ToString() + ",";
                strSQL += "'" + textBox1.Text + "',";
                strSQL += "'" + fn.DateToStr(dateTimePicker1.Value) + "',";
                strSQL += car_id.ToString() + ",";
                strSQL += fn.GetDriver(comboBox3) + ",";
                strSQL += "'" + fn.DateToStr(dateTimePicker1.Value) + (maskedTextBox1.Text.Trim() == ":" ? "" : " " + maskedTextBox1.Text) + "',";
                strSQL += "'" + fn.DateToStr(dateTimePicker1.Value) + (maskedTextBox2.Text.Trim() == ":" ? "" : " " + maskedTextBox2.Text) + "',";
                strSQL += beg_mileage.ToString() + ",";
                strSQL += fn.NumStr(textBox8) + ",";
                strSQL += fn.NumStr(beg_fuel.ToString()) + ",";
                strSQL += fn.NumStr(end_fuel.ToString()) + ",";
                strSQL += fc_id.ToString() + ",";
                strSQL += fn.NumStr(textBox12) + ",";
                strSQL += fn.NumStr(textBox11) + ",";
                strSQL += fn.NumStr(rasx_base.ToString()) + ",";
                strSQL += fn.NumStr(rasx_gorod.ToString()) + ",";
                strSQL += fn.NumStr(rasx_trassa.ToString()) + ",";
                strSQL += fn.NumStr(rasx_fuel.ToString()) + ",";
                strSQL += fn.NumStr(fuel_ekonom.ToString()) + ",";
                strSQL += "'" + dispatcher + "',";
                strSQL += "'" + mechanic + "',";
                strSQL += (checkBox1.Checked ? "1" : "0") + ")";
                ClSQL.ExecuteSQL(strSQL);
                doc_id = ClSQL.SelectIntCell("select top 1 scope_identity()");
                // select top 1 ident_current('put_lists')
                // select @@IDENTITY
            }
            EditMode = true;
            SaveOborot();
            CheckNextLists();
        }

        private void CheckNextLists()
        {
            bool fldt = true; string ss = ClSQL.SelectCell("select curvalue from settings where setkod='enable_downtime'");
            if (ss != "Да" && ss != "да" && ss != "ДА") fldt = false;
            DataTable DTpr = new DataTable();
            int tek_mileage = fn.StrToInt(textBox8);
            decimal tek_fuel = (decimal)end_fuel;
            decimal trasx_fuel;
            int totkm;
            bool flpr = false;
            string uslov = "where car_id=" + car_id.ToString() + " and pl_date>='" + fn.DateToStr(dateTimePicker1.Value) + "' ";
            DataTable dt = ClSQL.SelectSQL("select * from put_lists " + uslov + "order by pl_date, pl_id");
            foreach (DataRow dr in dt.Rows)
            {
                if (DateTime.Parse(dr["pl_date"].ToString()) == dateTimePicker1.Value.Date && (int)dr["pl_id"] <= doc_id) continue;
                if ((int)dr["beg_mileage"] != tek_mileage || (decimal)(double)dr["beg_fuel"] != tek_fuel)
                {
                    flpr = true; break;
                }
                /* Расчет расхода и остатка топлива */
                totkm = (int)dr["end_mileage"] - tek_mileage;
                if (totkm < 0)
                {
                    trasx_fuel = 0;
                    tek_fuel = 0;
                }
                else
                {
                    if ((DateTime)dr["pl_date"] < new DateTime(2013, 09, 16))
                    {
                        trasx_fuel = Decimal.Round((totkm * (decimal)(double)dr["rasx_gorod"]) / 100, 1, MidpointRounding.AwayFromZero);
                        tek_fuel = tek_fuel + Decimal.Round((decimal)(double)dr["fuel_in"], 1, MidpointRounding.AwayFromZero) - trasx_fuel;
                    }
                    else if ((DateTime)dr["pl_date"] < new DateTime(2017, 09, 04))
                    {
                        trasx_fuel = (totkm * (decimal)(double)dr["rasx_gorod"]) / 100;
                        if (fldt) trasx_fuel = trasx_fuel + (decimal)(double)dr["rasx_base"] * (decimal)(double)dr["downtime"];
                        tek_fuel = tek_fuel + (decimal)(double)dr["fuel_in"] - trasx_fuel;
                    }
                    else
                    {
                        DataRow tdr = ClSQL.SelectRow("select * from (select sum(mileage) as mileage_gorod from put_lists_t where pl_id=" + dr["pl_id"] + " and mtype=1) t1, (select sum(mileage) as mileage_trassa from put_lists_t where pl_id=" + dr["pl_id"] + " and mtype=2) t2");
                        decimal mileage_gorod = tdr["mileage_gorod"].ToString() == "" ? 0 : (decimal)(double)tdr["mileage_gorod"];
                        decimal mileage_trassa = tdr["mileage_trassa"].ToString() == "" ? 0 : (decimal)(double)tdr["mileage_trassa"];
                        if (mileage_gorod > 0 || mileage_trassa > 0)
                        {
                            trasx_fuel = ((mileage_gorod * (decimal)(double)dr["rasx_gorod"]) / 100) + ((mileage_trassa * (decimal)(double)dr["rasx_trassa"]) / 100);
                            if (fldt) trasx_fuel = trasx_fuel + (decimal)(double)dr["rasx_base"] * (decimal)(double)dr["downtime"];
                            tek_fuel = tek_fuel + (decimal)(double)dr["fuel_in"] - trasx_fuel;
                        }
                        else
                        {
                            trasx_fuel = (totkm * (decimal)(double)dr["rasx_gorod"]) / 100;
                            if (fldt) trasx_fuel = trasx_fuel + (decimal)(double)dr["rasx_base"] * (decimal)(double)dr["downtime"];
                            tek_fuel = tek_fuel + (decimal)(double)dr["fuel_in"] - trasx_fuel;
                        }
                    }
                }
                tek_mileage = (int)dr["end_mileage"];
                if ((decimal)(double)dr["end_fuel"] != tek_fuel)
                {
                    flpr = true; break;
                }
            }
            if (flpr)
            {
                PutList_pr prFrm = new PutList_pr();
                prFrm.doc_id = doc_id;
                prFrm.ShowDialog();
            }
        }

        private void UpdateDriverTime()
        {
            /* происходит при записи ПЛ и при "уходе" с закладки "Обортная сторона" */
            /* пока-что сказали, что это не нужно */
            /* string s1, s2;
            int i;
            if (DTpl.Rows.Count > 0)
            {
                s1 = DTpl.Rows[0][2].ToString(); if (s1.Length == 1) s1 = "0" + s1;
                s2 = DTpl.Rows[0][3].ToString(); if (s2.Length == 1) s2 = "0" + s2;
                if (s1 != "" && s2 != "" && Int32.TryParse(s1, out i) && Int32.TryParse(s2, out i)) maskedTextBox1.Text = s1 + ":" + s2;
                s1 = DTpl.Rows[DTpl.Rows.Count - 1][4].ToString(); if (s1.Length == 1) s1 = "0" + s1;
                s2 = DTpl.Rows[DTpl.Rows.Count - 1][5].ToString(); if (s2.Length == 1) s2 = "0" + s2;
                if (s1 != "" && s2 != "" && Int32.TryParse(s1, out i) && Int32.TryParse(s2, out i)) maskedTextBox2.Text = s1 + ":" + s2;
            }
            */
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!HasErrors())
            {
                UpdateDriverTime();
                SavePutList();
                Close();
            }
        }

        /* ****************** Load form ****************** */

        private void OborotTableCreate()
        {
            DTpl.Columns.Clear();
            DTpl.Columns.Add("place_out");
            DTpl.Columns.Add("place_in");
            DTpl.Columns.Add("time_out_h");
            DTpl.Columns.Add("time_out_m");
            DTpl.Columns.Add("time_in_h");
            DTpl.Columns.Add("time_in_m");
            DTpl.Columns.Add("mileage");
            DTpl.Columns.Add("mtype");
            BSpl.DataSource = DTpl;
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = BSpl;

            DataGridViewColumn column1 = new DataGridViewTextBoxColumn(); column1.DataPropertyName = "place_out"; dataGridView1.Columns.Add(column1);
            DataGridViewColumn column2 = new DataGridViewTextBoxColumn(); column2.DataPropertyName = "place_in"; dataGridView1.Columns.Add(column2);
            DataGridViewColumn column3 = new DataGridViewTextBoxColumn(); column3.DataPropertyName = "time_out_h"; dataGridView1.Columns.Add(column3);
            DataGridViewColumn column4 = new DataGridViewTextBoxColumn(); column4.DataPropertyName = "time_out_m"; dataGridView1.Columns.Add(column4);
            DataGridViewColumn column5 = new DataGridViewTextBoxColumn(); column5.DataPropertyName = "time_in_h"; dataGridView1.Columns.Add(column5);
            DataGridViewColumn column6 = new DataGridViewTextBoxColumn(); column6.DataPropertyName = "time_in_m"; dataGridView1.Columns.Add(column6);
            DataGridViewColumn column7 = new DataGridViewTextBoxColumn(); column7.DataPropertyName = "mileage"; dataGridView1.Columns.Add(column7);
            DataGridViewComboBoxColumn mtypes = new DataGridViewComboBoxColumn();
            mtypes.Items.AddRange("Город", "Трасса");
            mtypes.DataPropertyName = "mtype";
            mtypes.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
            mtypes.FlatStyle = FlatStyle.Flat;
            dataGridView1.Columns.Add(mtypes);

            dataGridView1.Columns[0].HeaderText = "Место отправления";
            dataGridView1.Columns[0].DisplayIndex = 0;
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[1].HeaderText = "Место назначения";
            dataGridView1.Columns[1].DisplayIndex = 1;
            dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridView1.Columns[1].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[2].HeaderText = "Выезд час.";
            dataGridView1.Columns[2].DisplayIndex = 2;
            dataGridView1.Columns[2].Width = 55;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[3].HeaderText = "Выезд мин.";
            dataGridView1.Columns[3].DisplayIndex = 3;
            dataGridView1.Columns[3].Width = 55;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[4].HeaderText = "Приезд час.";
            dataGridView1.Columns[4].DisplayIndex = 4;
            dataGridView1.Columns[4].Width = 55;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[5].HeaderText = "Приезд мин.";
            dataGridView1.Columns[5].DisplayIndex = 5;
            dataGridView1.Columns[5].Width = 55;
            dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[5].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[6].HeaderText = "Пробег";
            dataGridView1.Columns[6].DisplayIndex = 6;
            dataGridView1.Columns[6].Width = 55;
            dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[6].SortMode = DataGridViewColumnSortMode.NotSortable;
            dataGridView1.Columns[7].HeaderText = "Вид пробега";
            dataGridView1.Columns[7].DisplayIndex = 7;
            dataGridView1.Columns[7].Width = 70;
            dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[7].SortMode = DataGridViewColumnSortMode.NotSortable;

            /*DataRow row = DTpl.NewRow();
            DTpl.Rows.Add(row);
            dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0];*/
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (this.dataGridView1.IsCurrentCellDirty)
            {
                if (this.dataGridView1.CurrentCell != null && !fl_load)
                {
                    if (this.dataGridView1.CurrentCell.ColumnIndex == 7)
                    {
                        this.dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    }
                }
            }
        }

        private void PutList_Load(object sender, EventArgs e)
        {
            string s, strSQL;
            fl_load = true;
            OborotTableCreate();
            if (EditMode)
            {
                strSQL = "select p.*, d.name as dep_name, c.marka, c.gosnomer, dr.udostov, f.fc_id, ft.ft_name ";
                strSQL += "from put_lists p, departments d, cars c, drivers dr, fuel_cards f, fuel_types ft ";
                strSQL += "where p.pl_id=" + doc_id + " and p.dep_id=d.dep_id and p.car_id=c.car_id and p.drv_id=dr.drv_id and p.fc_id=f.fc_id and f.ft_id=ft.ft_id";
                DataRow dr = ClSQL.SelectRow(strSQL);
                if ((int)dr["status"] == 1) ReadOnly = true;
                textBox1.Text = dr["pl_nom"].ToString();
                dateTimePicker1.Value = DateTime.Parse(dr["pl_date"].ToString());
                textBox2.Text = dr["nom_do"].ToString();
                textBox3.Text = dr["dep_name"].ToString();
                textBox4.Text = dr["udostov"].ToString();
                textBox5.Text = dr["marka"].ToString();
                textBox6.Text = dr["gosnomer"].ToString();
                fn.UpdateFuelCards(comboBox2, (int)dr["fc_id"]);
                label31.Text = dr["ft_name"].ToString();
                maskedTextBox1.Text = fn.TimeFromDateTime(dr["beg_time"].ToString());
                maskedTextBox2.Text = fn.TimeFromDateTime(dr["end_time"].ToString());
                textBox7.Text = fn.Empty(dr["beg_mileage"].ToString());
                textBox8.Text = fn.Empty(dr["end_mileage"].ToString());
                beg_fuel = (double)dr["beg_fuel"];
                textBox9.Text = fn.Empty(Decimal.Round((decimal)beg_fuel, 1, MidpointRounding.AwayFromZero).ToString());
                end_fuel = (double)dr["end_fuel"];
                textBox10.Text = fn.Empty(Decimal.Round((decimal)end_fuel, 1, MidpointRounding.AwayFromZero).ToString());
                textBox12.Text = fn.Empty(dr["fuel_in"].ToString());
                textBox11.Text = fn.Empty(dr["downtime"].ToString());
                textBox13.Text = fn.Empty(dr["rasx_gorod"].ToString());
                textBox25.Text = fn.Empty(dr["rasx_trassa"].ToString());
                textBox18.Text = dr["dispatcher"].ToString();
                textBox19.Text = dr["mechanic"].ToString();
                checkBox1.Checked = (int)dr["status"] == 1 ? true : false;
                fn.UpdateDrivers(comboBox3, (int)dr["drv_id"]);
                nom_do = dr["nom_do"].ToString();
                dep_id = (int)dr["dep_id"];
                car_id = (int)dr["car_id"];
                fc_id = (int)dr["fc_id"];
                beg_mileage = (int)dr["beg_mileage"];
                ft_name = dr["ft_name"].ToString();
                rasx_gorod = dr["rasx_gorod"].ToString() == "" ? 0 : (double)dr["rasx_gorod"];
                rasx_trassa = dr["rasx_trassa"].ToString()=="" ? 0 : (double)dr["rasx_trassa"];
                rasx_base = (double)dr["rasx_base"];
                if (rasx_base == 0) rasx_base = fn.GetRasxBase(car_id, dateTimePicker1.Value);
                dispatcher = dr["dispatcher"].ToString();
                mechanic = dr["mechanic"].ToString();
                if (dr["rasx_trassa"].ToString() == "" && dateTimePicker1.Value.Date >= new DateTime(2017, 09, 04))
                {
                    rasx_trassa = fn.GetRasxNorm(car_id, dateTimePicker1.Value, beg_mileage, 2);
                    textBox25.Text = fn.Empty(rasx_trassa.ToString());
                    if (rasx_trassa == -1) { MessageBox.Show("На " + fn.DateToStrR(dateTimePicker1.Value) + " не указан расход топлива \"Трасса\" по норме !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); textBox25.Text = ""; }
                }
                /* ************************** */
                object [] rowArray = new object[8];
                strSQL = "select * from put_lists_t where pl_id=" + doc_id.ToString() + " order by npp, time_out";
                DataTable dt = ClSQL.SelectSQL(strSQL);
                if (dt.Rows.Count < 1)
                {
                    OborotAddNewRow();
                }
                else
                {
                    foreach (DataRow r in dt.Rows)
                    {
                        DataRow row = DTpl.NewRow();
                        rowArray[0] = r["place_out"].ToString();
                        rowArray[1] = r["place_in"].ToString();
                        s = fn.TimeFromDateTime(r["time_out"].ToString());
                        rowArray[2] = s.IndexOf(":") == -1 ? "" : s.Substring(0, s.IndexOf(":"));
                        rowArray[3] = s.IndexOf(":") == -1 ? "" : s.Substring(s.IndexOf(":") + 1);
                        s = fn.TimeFromDateTime(r["time_in"].ToString());
                        rowArray[4] = s.IndexOf(":") == -1 ? "" : s.Substring(0, s.IndexOf(":"));
                        rowArray[5] = s.IndexOf(":") == -1 ? "" : s.Substring(s.IndexOf(":") + 1);
                        rowArray[6] = r["mileage"].ToString();
                        rowArray[7] = r["mtype"].ToString() == "" ? "" : ((int)r["mtype"] == 1 ? "Город" : ((int)r["mtype"] == 2 ? "Трасса" : ""));
                        row.ItemArray = rowArray;
                        DTpl.Rows.Add(row);
                    }
                }
                /* ************************** */
                fl_load = false;
                MileageSum();
                ShowInfoRasch();
                if (Program.UserType > 4)
                {
                    SwitchEnabledFront(false);
                }
                if (ReadOnly)
                {
                    SwitchEnabledAll(false);
                }
            }
            else
            {
                textBox1.Text = (fn.GetMaxNom("pl_nom", "put_lists where nom_do='" + Program.UserNomDo + "'") + 1).ToString();
                dateTimePicker1.Value = DateTime.Now;
                strSQL = "select count(*) from drivers where status<>1";
                if (Program.UserType > 3) strSQL += " and nom_do='" + Program.UserNomDo + "'";
                int k = ClSQL.SelectIntCell(strSQL);
                if (k > 1) fn.UpdateDrivers(comboBox3, 0);
                if (k == 1)
                {
                    fn.UpdateDrivers(comboBox3, ClSQL.SelectIntCell(strSQL.Replace("count(*)", "drv_id")));
                    fl_load = false;
                    ShowInfoDriver();
                }
                OborotAddNewRow();
            }
            if (Program.UserType > 2) checkBox1.Visible = false;
            string ss = ClSQL.SelectCell("select curvalue from settings where setkod='enable_downtime'");
            if (ss != "Да" && ss != "да" && ss != "ДА")
            {
                label18.Visible = false;
                textBox11.Visible = false;
            }
            fl_load = false;
        }

        /* ****************** Print ****************** */

        private void button3_Click(object sender, EventArgs e)
        {
            if (Changed || !EditMode)
            {
                if (MessageBox.Show("Перед печатью документ должен быть записан !" + Environment.NewLine + "Записать и распечатать ?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    if (!HasErrors())
                    {
                        SavePutList();
                        button3.Enabled = false;
                        PrintPutList(doc_id);
                        button3.Enabled = true;
                    }
                }
            }
            else
            {
                button3.Enabled = false;
                PrintPutList(doc_id);
                button3.Enabled = true;
            }
        }

        private static void WaitFormShow()
        {
            SplashFrmWait fw = new SplashFrmWait();
            fw.ShowDialog();
        }

        public void PrintPutList(int docid)
        {
            DateTime date1 = new DateTime(2016, 9, 7);
            DataRow pl = ClSQL.SelectRow("select * from put_lists where pl_id=" + docid.ToString());
            if ((DateTime)pl["pl_date"] < date1) PrintPutList_2(docid);
            PrintPutList_1(docid);
        }

        public void PrintPutList_1(int docid)
        {
            Thread WaitFormThread = new Thread(new ThreadStart(WaitFormShow));
            WaitFormThread.Start();
            ClExcel Excel = new ClExcel();
            string FileName = Program.TmpPath + @"\~putlist_" + DateTime.Now.ToString("ddMMHH_mmss") + ".xls";
            File.Copy(Program.ProgPath + @"\Templates\putlist.xls", FileName);
            Excel.Open(FileName);
            Excel.SelectSheet("Лист1");
            DataRow pl = ClSQL.SelectRow("select * from put_lists where pl_id=" + docid.ToString());
            DataRow dep = ClSQL.SelectRow("select * from departments where dep_id=" + pl["dep_id"].ToString());
            DataRow drv = ClSQL.SelectRow("select * from drivers where drv_id=" + pl["drv_id"].ToString());
            DataRow car = ClSQL.SelectRow("select * from cars where car_id=" + pl["car_id"].ToString());
            DataRow fc = ClSQL.SelectRow("select fc.*,ft.ft_name from fuel_cards fc, fuel_types ft where fc.ft_id=ft.ft_id and fc.fc_id=" + pl["fc_id"].ToString());
            string sdate = ((DateTime)pl["pl_date"]).ToString("D");
            string vrasp = dep["name"].ToString().Trim();
            string vadr = dep["address"].ToString().Trim();
            string[] regnom = car["reg_nom"].ToString().Split(' ');
            if (regnom.Count() == 1) { Array.Resize(ref regnom, 3); regnom[1] = ""; regnom[2] = ""; }
            if (regnom.Count() == 2) { Array.Resize(ref regnom, 3); regnom[2] = ""; }
            string filial = ClSQL.SelectCell("select curvalue from settings where setkod='filial_name'") + ", " +
                            ClSQL.SelectCell("select curvalue from settings where setkod='filial_addr'") + ", " +
                            ClSQL.SelectCell("select curvalue from settings where setkod='filial_phone'");
            string addstr = "";
            bool fldt = true; string ss = ClSQL.SelectCell("select curvalue from settings where setkod='enable_downtime'");
            if (ss != "Да" && ss != "да" && ss != "ДА") fldt = false;
            if (fldt) addstr = pl["downtime"].ToString() == "0" || pl["downtime"].ToString() == "" ? "" : "Простой: " + pl["downtime"].ToString() + " ч.";

            double plmtime = 5;
            ss = ClSQL.SelectCell("select curvalue from settings where setkod='dep_plmtime_" + dep_id.ToString() + "'");
            if (ss.Trim() == "") ss = ClSQL.SelectCell("select curvalue from settings where setkod='plm_time'");
            if (ss.Trim() != "" && ss.Trim() != "0") plmtime = fn.StrToDouble(ss);
            DateTime begtime = (DateTime)pl["beg_time"];      // время выезда
            DateTime endtime = (DateTime)pl["end_time"];      // время возвращения
            DateTime mbegtime = begtime.AddMinutes(-plmtime); // время выдачи ПЛ
            DateTime mendtime = endtime.AddMinutes(plmtime);  // время сдачи ПЛ
            string vplout = fn.TimeFromDateTime(begtime.ToString());
            string vplin = fn.TimeFromDateTime(endtime.ToString());
            string vmplout = vplout == "" ? "" : fn.TimeFromDateTime(mbegtime.ToString());
            string vmplin = vplin == "" ? "" : fn.TimeFromDateTime(mendtime.ToString());
            
            double rest_time = 0;
            ss = ClSQL.SelectCell("select curvalue from settings where setkod='dep_resttime_" + dep_id.ToString() + "'");
            if (ss.Trim() == "") ss = ClSQL.SelectCell("select curvalue from settings where setkod='rest_time'");
            if (ss.Trim() != "" && ss.Trim() != "0") rest_time = fn.StrToDouble(ss);
            double pltime1 = vplout == "" || vplin == "" ? 0 : Math.Round((endtime - begtime).TotalHours - rest_time / 60, 2);
            double pltime2 = vplout == "" || vplin == "" ? 0 : Math.Round((mendtime - mbegtime).TotalHours - rest_time / 60, 2);

            Excel.SetByNameValue("Филиал", filial);
            Excel.SetByNameValue("НомДок", pl["pl_nom"].ToString() == "" ? "" : pl["pl_nom"].ToString());
            Excel.SetByNameValue("ДатаДок", sdate);
            Excel.SetByNameValue("ДатаДок_до", sdate);
            Excel.SetByNameValue("МаркаАвто", car["marka"].ToString() == "" ? "" : car["marka"].ToString());
            Excel.SetByNameValue("Госномер", car["gosnomer"].ToString() == "" ? "" : car["gosnomer"].ToString());
            Excel.SetByNameValue("Регном1", regnom[0] == "" ? "" : regnom[0]);
            Excel.SetByNameValue("Регном2", regnom[1] == "" ? "" : regnom[1]);
            Excel.SetByNameValue("Регном3", regnom[2] == "" ? "" : regnom[2]);
            Excel.SetByNameValue("Гарномер", car["garnomer"].ToString() == "" ? "" : car["garnomer"].ToString());
            Excel.SetByNameValue("Табномер", drv["tab_no"].ToString() == "" ? "" : drv["tab_no"].ToString());
            Excel.SetByNameValue("Водитель", drv["fio"].ToString() == "" ? "" : drv["fio"].ToString());
            Excel.SetByNameValue("ВодитСокр1", drv["fio"].ToString() == "" ? "" : fn.GetFIO(drv["fio"].ToString()));
            Excel.SetByNameValue("ВодитСокр2", drv["fio"].ToString() == "" ? "" : fn.GetFIO(drv["fio"].ToString()));
            Excel.SetByNameValue("Удостоверение", drv["udostov"].ToString() == "" ? "" : drv["udostov"].ToString());
            Excel.SetByNameValue("Класс", drv["klass"].ToString() == "" ? "" : drv["klass"].ToString());
            Excel.SetByNameValue("ТоплКарта", fc["fc_nomer"].ToString() == "" ? "" : fc["fc_nomer"].ToString());
            Excel.SetByNameValue("ВРаспоряж", vrasp);
            Excel.SetByNameValue("ВАдрес", vadr);
            Excel.SetByNameValue("ВремяВыдачиПЛ", vmplout);
            Excel.SetByNameValue("ВремяСдачиПЛ", vmplin);
            Excel.SetByNameValue("ВремяВыезда", vplout);
            Excel.SetByNameValue("ВремяВозвращ", vplin);
            Excel.SetByNameValue("Диспетчер1", pl["dispatcher"].ToString() == "" ? "" : fn.GetFIO(pl["dispatcher"].ToString()));
            Excel.SetByNameValue("Диспетчер2", pl["dispatcher"].ToString() == "" ? "" : fn.GetFIO(pl["dispatcher"].ToString()));
            Excel.SetByNameValue("Механик1", pl["mechanic"].ToString() == "" ? "" : fn.GetFIO(pl["mechanic"].ToString()));
            Excel.SetByNameValue("Механик2", pl["mechanic"].ToString() == "" ? "" : fn.GetFIO(pl["mechanic"].ToString()));
            Excel.SetByNameValue("Дополнительно1", addstr);
            Excel.SetByNameValue("МаркаТопл", fc["ft_name"].ToString() == "" ? "" : fc["ft_name"].ToString());
            if (pl["fuel_in"].ToString() != "0" && pl["fuel_in"].ToString() != "")         Excel.SetByNameValue("Залил",  (double)pl["fuel_in"]);
            if (pl["beg_mileage"].ToString() != "0" && pl["beg_mileage"].ToString() != "") Excel.SetByNameValue("ПробегНач",  (int)pl["beg_mileage"]);
            if (pl["end_mileage"].ToString() != "0" && pl["end_mileage"].ToString() != "") Excel.SetByNameValue("ПробегКон",  (int)pl["end_mileage"]);
            if (pl["beg_fuel"].ToString() != "0" && pl["beg_fuel"].ToString() != "")       Excel.SetByNameValue("ОстТоплНач",  Math.Round((double)pl["beg_fuel"], 1));
            if (pl["end_fuel"].ToString() != "0" && pl["end_fuel"].ToString() != "")       Excel.SetByNameValue("ОстТоплКон",  Math.Round((double)pl["end_fuel"], 1));
            if (pl["rasx_fuel"].ToString() != "0" && pl["rasx_fuel"].ToString() != "")     Excel.SetByNameValue("РасхНорм", Math.Round((double)pl["rasx_fuel"], 1));
            if (pl["rasx_fuel"].ToString() != "0" && pl["rasx_fuel"].ToString() != "")     Excel.SetByNameValue("РасхФакт", Math.Round((double)pl["rasx_fuel"], 1));
            Excel.SetByNameValue("Перерасход", "");
            Excel.SetByNameValue("Экономия", "");
            // *****
            DataTable dt = ClSQL.SelectSQL("select * from put_lists_t where pl_id=" + docid.ToString() + " order by npp, time_out");
            int npp = 0;
            foreach (DataRow dr in dt.Rows)
            {
                npp++;
                string vrout = fn.TimeFromDateTime(dr["time_out"].ToString());
                string vrin = fn.TimeFromDateTime(dr["time_in"].ToString());
                Excel.SetCellYXValue(70 + npp,  1, npp.ToString());
                Excel.SetCellYXValue(70 + npp,  4, dr["place_out"].ToString());
                Excel.SetCellYXValue(70 + npp, 29, dr["place_in"].ToString());
                Excel.SetCellYXValue(70 + npp, 54, vrout.Trim() == "" ? "" : vrout.Substring(0, vrout.IndexOf(":")));
                Excel.SetCellYXValue(70 + npp, 57, vrout.Trim() == "" ? "" : vrout.Substring(vrout.IndexOf(":") + 1));
                Excel.SetCellYXValue(70 + npp, 60, vrin.Trim() == "" ? "" : vrin.Substring(0, vrin.IndexOf(":")));
                Excel.SetCellYXValue(70 + npp, 63, vrin.Trim() == "" ? "" : vrin.Substring(vrin.IndexOf(":") + 1));
                Excel.SetCellYXValue(70 + npp, 66, (double)dr["mileage"]);
            }
            if (pl["beg_mileage"].ToString() != "0" && pl["beg_mileage"].ToString() != "" && pl["end_mileage"].ToString() != "0" && pl["end_mileage"].ToString() != "") Excel.SetByNameValue("ПройденоКм", (int)pl["end_mileage"] - (int)pl["beg_mileage"]);
            if (pltime1 != 0) Excel.SetByNameValue("ВремяВНаряде", pltime1);
            if (pltime2 != 0) Excel.SetByNameValue("ВремяОтработ", pltime2);
            if (rest_time != 0) Excel.SetByNameValue("ВремяОтдых", Math.Round(rest_time / 60, 2));
            Excel.SetByNameValue("Механик3", pl["mechanic"].ToString() == "" ? "" : fn.GetFIO(pl["mechanic"].ToString()));
            WaitFormThread.Abort();
            Excel.Show();
            fn.SetActiveExcel();
        }

        public void PrintPutList_2(int docid)
        {
            Thread WaitFormThread = new Thread(new ThreadStart(WaitFormShow));
            WaitFormThread.Start();
            string strSQL;
            string FileName = Program.TmpPath + @"\~putlist_" + DateTime.Now.ToString("ddMMHH_mmss") + ".doc";
            File.Copy(Program.ProgPath + @"\Templates\putlist.doc", FileName);
            ClWord ClWord = new ClWord();
            ClWord.Open(FileName);
            DataRow pl = ClSQL.SelectRow("select * from put_lists where pl_id=" + docid.ToString());
            DataRow dep = ClSQL.SelectRow("select * from departments where dep_id=" + pl["dep_id"].ToString());
            DataRow drv = ClSQL.SelectRow("select * from drivers where drv_id=" + pl["drv_id"].ToString());
            DataRow car = ClSQL.SelectRow("select * from cars where car_id=" + pl["car_id"].ToString());
            DataRow fc = ClSQL.SelectRow("select fc.*,ft.ft_name from fuel_cards fc, fuel_types ft where fc.ft_id=ft.ft_id and fc.fc_id=" + pl["fc_id"].ToString());
            string sdate = ((DateTime)pl["pl_date"]).ToString("D");
            string dd1 = sdate.Substring(0, sdate.IndexOf(" "));
            sdate = sdate.Substring(sdate.IndexOf(" ") + 1);
            string dd2 = sdate.Substring(0, sdate.IndexOf(" "));
            sdate = sdate.Substring(sdate.IndexOf(" ") + 1);
            string dd3 = sdate.Substring(0, sdate.IndexOf(" "));
            string vrasp = dep["name"].ToString() + " ";
            int pp = -1; int p = vrasp.IndexOf(" ");
            while (p != -1 && p < 33) { pp = p; p = vrasp.IndexOf(" ", pp + 1); }
            string vrasp1 = vrasp.Substring(0, pp).Trim();
            string vrasp2 = vrasp.Substring(pp + 1).Trim();
            if (vrasp2 == "") vrasp2 = " ";
            string vadr = dep["address"].ToString().Trim() + " ";
            pp = -1; p = vadr.IndexOf(" ");
            while (p != -1 && p < 33) { pp = p; p = vadr.IndexOf(" ", pp + 1); }
            string vadr1 = vadr.Substring(0, pp).Trim();
            string vadr2 = vadr.Substring(pp + 1).Trim();
            if (vadr2 == "") vadr2 = " ";
            string vrout = fn.TimeFromDateTime(pl["beg_time"].ToString());
            string vrout_h = vrout.Trim() == "" ? " " : vrout.Substring(0, vrout.IndexOf(":"));
            string vrout_m = vrout.Trim() == "" ? " " : vrout.Substring(vrout.IndexOf(":") + 1);
            string vrin = fn.TimeFromDateTime(pl["end_time"].ToString());
            string vrin_h = vrin.Trim() == "" ? " " : vrin.Substring(0, vrin.IndexOf(":"));
            string vrin_m = vrin.Trim() == "" ? " " : vrin.Substring(vrin.IndexOf(":") + 1);
            string [] regnom = car["reg_nom"].ToString().Split(' ');
            if (regnom.Count() == 1) { Array.Resize(ref regnom, 3); regnom[1] = ""; regnom[2] = ""; }
            if (regnom.Count() == 2) { Array.Resize(ref regnom, 3); regnom[2] = ""; }
            string filial = ClSQL.SelectCell("select curvalue from settings where setkod='filial_name'") + ", " +
                            ClSQL.SelectCell("select curvalue from settings where setkod='filial_addr'") + ", " +
                            ClSQL.SelectCell("select curvalue from settings where setkod='filial_phone'");
            string addstr = " ";
            bool fldt = true; string ss = ClSQL.SelectCell("select curvalue from settings where setkod='enable_downtime'");
            if (ss != "Да" && ss != "да" && ss != "ДА") fldt = false;
            if (fldt) addstr = pl["downtime"].ToString() == "0" || pl["downtime"].ToString() == "" ? " " : "Простой: " + pl["downtime"].ToString() + " ч.";
            ClWord.SetVar("Филиал", filial);
            ClWord.SetVar("НомДок", pl["pl_nom"].ToString() == "" ? " " : pl["pl_nom"].ToString());
            ClWord.SetVar("ДатаДок1", dd1);
            ClWord.SetVar("ДатаДок2", dd2);
            ClWord.SetVar("ДатаДок3", dd3);
            ClWord.SetVar("МаркаАвто", car["marka"].ToString() == "" ? " " : car["marka"].ToString());
            ClWord.SetVar("Госномер", car["gosnomer"].ToString() == "" ? " " : car["gosnomer"].ToString());
            ClWord.SetVar("Регном1", regnom[0] == "" ? " " : regnom[0]);
            ClWord.SetVar("Регном2", regnom[1] == "" ? " " : regnom[1]);
            ClWord.SetVar("Регном3", regnom[2] == "" ? " " : regnom[2]);
            ClWord.SetVar("Гарномер", car["garnomer"].ToString() == "" ? " " : car["garnomer"].ToString());
            ClWord.SetVar("Табномер", drv["tab_no"].ToString() == "" ? " " : drv["tab_no"].ToString());
            ClWord.SetVar("Водитель", drv["fio"].ToString() == "" ? " " : drv["fio"].ToString());
            ClWord.SetVar("ВодитСокр", drv["fio"].ToString() == "" ? " " : fn.GetFIO(drv["fio"].ToString()));
            ClWord.SetVar("Удостоверение", drv["udostov"].ToString() == "" ? " " : drv["udostov"].ToString());
            ClWord.SetVar("Класс", drv["klass"].ToString() == "" ? " " : drv["klass"].ToString());
            ClWord.SetVar("ТоплКарта", fc["fc_nomer"].ToString() == "" ? " " : fc["fc_nomer"].ToString());
            ClWord.SetVar("ВРаспоряж1", vrasp1);
            ClWord.SetVar("ВРаспоряж2", vrasp2);
            ClWord.SetVar("ВАдрес1", vadr1);
            ClWord.SetVar("ВАдрес2", vadr2);
            ClWord.SetVar("OH", vrout_h == "" ? " " : vrout_h);
            ClWord.SetVar("OM", vrout_m == "" ? " " : vrout_m);
            ClWord.SetVar("IH", vrin_h == "" ? " " : vrin_h);
            ClWord.SetVar("IM", vrin_m == "" ? " " : vrin_m);
            ClWord.SetVar("Диспетчер", pl["dispatcher"].ToString() == "" ? " " : fn.GetFIO(pl["dispatcher"].ToString()));
            ClWord.SetVar("Механик", pl["mechanic"].ToString() == "" ? " " : fn.GetFIO(pl["mechanic"].ToString()));
            ClWord.SetVar("Дополнительно", addstr);
            ClWord.SetVar("МаркаТопл", fc["ft_name"].ToString() == "" ? " " : fc["ft_name"].ToString());
            ClWord.SetVar("Залил", pl["fuel_in"].ToString() == "0" || pl["fuel_in"].ToString() == "" ? " " : pl["fuel_in"].ToString());
            ClWord.SetVar("ПробегНач", pl["beg_mileage"].ToString() == "0" || pl["beg_mileage"].ToString() == "" ? " " : pl["beg_mileage"].ToString());
            ClWord.SetVar("ПробегКон", pl["end_mileage"].ToString() == "0" || pl["end_mileage"].ToString() == "" ? " " : pl["end_mileage"].ToString());
            ClWord.SetVar("ПройденоКм", pl["beg_mileage"].ToString() == "0" || pl["beg_mileage"].ToString() == "" || pl["end_mileage"].ToString() == "0" || pl["end_mileage"].ToString() == "" ? " " : ((int)pl["end_mileage"] - (int)pl["beg_mileage"]).ToString());
            ClWord.SetVar("ОстТоплНач", pl["beg_fuel"].ToString() == "0" || pl["beg_fuel"].ToString() == "" ? " " : Decimal.Round((decimal)(double)pl["beg_fuel"], 1, MidpointRounding.AwayFromZero).ToString());
            ClWord.SetVar("ОстТоплКон", pl["end_fuel"].ToString() == "0" || pl["end_fuel"].ToString() == "" ? " " : Decimal.Round((decimal)(double)pl["end_fuel"], 1, MidpointRounding.AwayFromZero).ToString());
            ClWord.SetVar("РасхНорм", pl["rasx_fuel"].ToString() == "0" || pl["rasx_fuel"].ToString() == "" ? " " : Decimal.Round((decimal)(double)pl["rasx_fuel"], 1, MidpointRounding.AwayFromZero).ToString());
            ClWord.SetVar("РасхФакт", pl["rasx_fuel"].ToString() == "0" || pl["rasx_fuel"].ToString() == "" ? " " : Decimal.Round((decimal)(double)pl["rasx_fuel"], 1, MidpointRounding.AwayFromZero).ToString());
            ClWord.SetVar("Перерасход", " ");
            ClWord.SetVar("Экономия", " ");
            // *****
            strSQL = "select * from put_lists_t where pl_id=" + docid.ToString() + " order by npp, time_out";
            DataTable dt = ClSQL.SelectSQL(strSQL);
            int npp = 0;
            foreach (DataRow dr in dt.Rows)
            {
                npp++;
                vrout = fn.TimeFromDateTime(dr["time_out"].ToString());
                vrin = fn.TimeFromDateTime(dr["time_in"].ToString());
                ClWord.SetCellValue(6, 1, 3 + npp, npp.ToString());
                ClWord.SetCellValue(6, 3, 3 + npp, dr["place_out"].ToString());
                ClWord.SetCellValue(6, 4, 3 + npp, dr["place_in"].ToString());
                ClWord.SetCellValue(6, 5, 3 + npp, vrout.Trim() == "" ? " " : vrout.Substring(0, vrout.IndexOf(":")));
                ClWord.SetCellValue(6, 6, 3 + npp, vrout.Trim() == "" ? " " : vrout.Substring(vrout.IndexOf(":") + 1));
                ClWord.SetCellValue(6, 7, 3 + npp, vrin.Trim() == "" ? " " : vrin.Substring(0, vrin.IndexOf(":")));
                ClWord.SetCellValue(6, 8, 3 + npp, vrin.Trim() == "" ? " " : vrin.Substring(vrin.IndexOf(":") + 1));
                ClWord.SetCellValue(6, 9, 3 + npp, dr["mileage"].ToString());
            }
            WaitFormThread.Abort();
            ClWord.Complete();
            fn.SetActiveWord();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (fl_load) return;
            Changed = true;
            ShowInfoDriver();
            ShowInfoRasch();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (fl_load) return;
            Changed = true;
            ShowInfoDriver();
            ShowInfoRasch();
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            if (fl_load) return;
            Changed = true;
            ShowInfoRasch();
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            if (fl_load) return;
            Changed = true;
            ShowInfoRasch();
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            if (fl_load) return;
            Changed = true;
            ShowInfoRasch();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) Close();
        }

        /* *************************************** */

        private void SwitchEnabledAll(bool b)
        {
            textBox1.Enabled = b;
            dateTimePicker1.Enabled = b;
            comboBox2.Enabled = b;
            comboBox3.Enabled = b;
            maskedTextBox1.Enabled = b;
            maskedTextBox2.Enabled = b;
            textBox8.Enabled = b;
            textBox12.Enabled = b;
            dataGridView1.ReadOnly = !b;
            toolStripButton1.Enabled = b;
            toolStripButton2.Enabled = b;
            toolStripButton3.Enabled = b;
            toolStripButton4.Enabled = b;
            toolStripButton5.Enabled = b;
            if (Program.UserType > 2) button1.Enabled = b; // кнопка "Записать"
            FEnabled = b;
        }

        private void SwitchEnabledFront(bool b)
        {
            // водители
            textBox1.Enabled = b;
            dateTimePicker1.Enabled = b;
            comboBox2.Enabled = b;
            comboBox3.Enabled = b;
            button3.Enabled = b;
        }

        private void OborotAddNewRow()
        {
            DataRow row = DTpl.NewRow();
            DTpl.Rows.Add(row);
            dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0];
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (!FEnabled) return;
            OborotAddNewRow();
            dataGridView1.Focus();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (!FEnabled) return;
            dataGridView1.BeginEdit(false);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 0)
            {
                DTpl.Rows[dataGridView1.CurrentCell.RowIndex].Delete();
                MileageSum();
            }
        }

        private void Control_TextChanged(object sender, System.EventArgs e)
        {
            if (dataGridView1.CurrentCell != null && !fl_load)
            {
                Changed = true;
                if (dataGridView1.CurrentCell.ColumnIndex == 6)
                {
                    decimal sum1 = 0, sum2 = 0, total = 0;
                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        if (i != dataGridView1.CurrentCell.RowIndex)
                        {
                            total += fn.StrToDecimal(dataGridView1.Rows[i].Cells[6].Value.ToString());
                            if (dataGridView1.Rows[i].Cells[7].Value.ToString() == "Город")
                            {
                                sum1 += fn.StrToDecimal(dataGridView1.Rows[i].Cells[6].Value.ToString());
                            }
                            else if (dataGridView1.Rows[i].Cells[7].Value.ToString() == "Трасса")
                            {
                                sum2 += fn.StrToDecimal(dataGridView1.Rows[i].Cells[6].Value.ToString());
                            }
                        }
                    }
                    total += fn.StrToDecimal(((TextBox)sender).Text);
                    if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString() == "Город")
                    {
                        sum1 += fn.StrToDecimal(((TextBox)sender).Text);
                    }
                    else if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[7].Value.ToString() == "Трасса")
                    {
                        sum2 += fn.StrToDecimal(((TextBox)sender).Text);
                    }
                    sum1 = Decimal.Round(sum1, 0, MidpointRounding.AwayFromZero);
                    textBox23.Text = sum1.ToString();
                    sum2 = Decimal.Round(sum2, 0, MidpointRounding.AwayFromZero);
                    textBox24.Text = sum2.ToString();
                    total = Decimal.Round(total, 0, MidpointRounding.AwayFromZero);
                    textBox20.Text = total.ToString();
                    if (dateTimePicker1.Value.Date >= new DateTime(2017, 09, 04) && !fl_load)
                    {
                        textBox21.Text = sum1.ToString();
                        textBox22.Text = sum2.ToString();
                        ShowInfoRasch();
                    }
                }
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentCell != null && !fl_load)
            {
                Changed = true;
                if (dataGridView1.CurrentCell.ColumnIndex == 7)
                {
                    MileageSum();
                }
            }
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.TextChanged -= new EventHandler(Control_TextChanged);
            e.Control.TextChanged += new EventHandler(Control_TextChanged);
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (!FEnabled) return;
            if (e.KeyCode == Keys.Delete)
            {
                dataGridView1.CurrentCell.Value = "";
                if (dataGridView1.CurrentCell.ColumnIndex == 6) MileageSum();
            }
            if (e.KeyCode == Keys.Insert)
            {
                OborotAddNewRow();
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1) dataGridView1.Focus();
        }

        private void PutList_Shown(object sender, EventArgs e)
        {
            if (EditMode)
            {
                button2.Focus();
            }
            else
            {
                if (comboBox3.SelectedIndex < 0) comboBox3.Focus(); else maskedTextBox1.Focus();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) SwitchEnabledAll(false); else SwitchEnabledAll(true);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (fl_load) return;
            Changed = true;
            label31.Text = ClSQL.SelectCell("select ft.ft_name from fuel_cards fc, fuel_types ft where fc.ft_id=ft.ft_id and fc.fc_id=" + fn.GetFuelCard(comboBox2));
            fc_id = fn.GetFuelCardInt(comboBox2);
            ft_name = label31.Text;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (fl_load) return;
            Changed = true;
        }

        private void maskedTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (fl_load) return;
            Changed = true;
        }

        private void maskedTextBox2_TextChanged(object sender, EventArgs e)
        {
            if (fl_load) return;
            Changed = true;
        }

        private void button2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)27) Close();
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell.RowIndex > 0)
            {
                int index = dataGridView1.CurrentCell.RowIndex;
                int column = dataGridView1.CurrentCell.ColumnIndex;
                DataRow row = DTpl.Rows[index];
                object[] values = DTpl.Rows[index].ItemArray;
                DTpl.Rows.RemoveAt(index);
                DTpl.Rows.InsertAt(row, index - 1);
                DTpl.Rows[index - 1].ItemArray = values;
                dataGridView1.CurrentCell = dataGridView1.Rows[index - 1].Cells[column];
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell.RowIndex < dataGridView1.Rows.Count - 1)
            {
                int index = dataGridView1.CurrentCell.RowIndex;
                int column = dataGridView1.CurrentCell.ColumnIndex;
                DataRow row = DTpl.Rows[index];
                object[] values = DTpl.Rows[index].ItemArray;
                DTpl.Rows.RemoveAt(index);
                DTpl.Rows.InsertAt(row, index + 1);
                DTpl.Rows[index + 1].ItemArray = values;
                dataGridView1.CurrentCell = dataGridView1.Rows[index + 1].Cells[column];
            }
        }

        private void tabPage2_Leave(object sender, EventArgs e)
        {
            UpdateDriverTime();
        }

    }
}
