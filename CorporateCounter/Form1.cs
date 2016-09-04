using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Configuration;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace CorporateCounter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region static string
        public static string compName;
        public static string licence;
        public static string addr;
        public static string scope;
        public static string tax;
        public static string account;
        public static string moneyType;
        public static string money;
        public static string genre;
        public static string onDate;
        public static string thruDate;
        public static string owner;
        public static string ownerID;
        public static string ownerPhone;
        public static string come;
        public static string comeID;
        public static string comePhone;
        public static string rely;
        public static string relyID;
        public static string relyPhone;
        public static string basee;
        public static string dateinchn = string.Format("{0:D}", DateTime.Now);
        #endregion

        //GetUpdateAppConfig http://www.cnblogs.com/luxiaoxun/p/3599341.html
        public static string GetAppConfig(string strKey)
        {
            string file = Application.ExecutablePath;
            Configuration config = ConfigurationManager.OpenExeConfiguration(file);
            foreach (string key in config.AppSettings.Settings.AllKeys)
            {
                if (key == strKey)
                {
                    return config.AppSettings.Settings[strKey].Value.ToString();
                }
            }
            return null;
        }

        public static void UpdateAppConfig(string newKey, string newValue)
        {
            string file = Application.ExecutablePath;
            Configuration config = ConfigurationManager.OpenExeConfiguration(file);
            bool exist = false;
            foreach (string key in config.AppSettings.Settings.AllKeys)
            {
                if (key == newKey)
                {
                    exist = true;
                }
            }
            if (exist)
            {
                config.AppSettings.Settings.Remove(newKey);
            }
            config.AppSettings.Settings.Add(newKey, newValue);
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }


        private void button1_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(requisition_PrintPage);
            try
            {
                pd.Print();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void requisition_PrintPage(object sender, PrintPageEventArgs e)
        {
            DateTime dt = DateTime.Now;
            Font ft = new Font("宋体", 11, FontStyle.Bold);
            SolidBrush sb = new SolidBrush(Color.Black);
            string CompName = compNametextBox1.Text;
            string Addr = addrtextBox1.Text;
            string OwnerPhone = ownerPhonetextBox5.Text;
            string Postal = GetAppConfig("Postal");      //Postal---224221
            string OwnerName = ownertextBox3.Text;
            string OwnerID = ownerIDtextBox4.Text;
            string Scope = scopetextBox2.Text;
            string AreaCode = GetAppConfig("AreaCode");  //AreaCode---311700
            string OwnerCer = GetAppConfig("OwnerCer");  //OwnerCer---身份证
            string CompCer = GetAppConfig("CompCer");    //CompCer---营业执照
            string Licence = licencetextBox2.Text;
            string Tax = taxtextBox7.Text;
            string BankName = GetAppConfig("BankName");  //BankName---
            string BankCode = GetAppConfig("BankCode");  //314311711013
            //账户性质 switch
            //行业分类 switch
            e.Graphics.DrawString(CompName, ft, sb, new PointF(230.0F, 140.0F));
            e.Graphics.DrawString(OwnerPhone, ft, sb, new PointF(600.0F, 140.0F));
            e.Graphics.DrawString(Addr, ft, sb, new PointF(230.0F, 160.0F));
            e.Graphics.DrawString(Postal, ft, sb, new PointF(600.0F, 160.0F));
            e.Graphics.DrawString(OwnerName, ft, sb, new PointF(380.0F, 219.0F));
            e.Graphics.DrawString(OwnerCer, ft, sb, new PointF(380.0F, 242.0F));
            e.Graphics.DrawString(OwnerID, ft, sb, new PointF(585.0F, 242.0F));
            e.Graphics.DrawString(AreaCode, ft, sb, new PointF(530.0F, 308.5F));
            e.Graphics.DrawString(Scope, ft, sb, new PointF(230.0F, 330.0F));
            e.Graphics.DrawString(CompCer, ft, sb, new PointF(230.0F, 354.0F));
            e.Graphics.DrawString(Licence, ft, sb, new PointF(527.0F, 354.0F));
            e.Graphics.DrawString(Tax, ft, sb, new PointF(380.0F, 407.0F));
            e.Graphics.DrawString(BankName, ft, sb, new PointF(225.0F, 700.0F));
            e.Graphics.DrawString(BankCode, ft, sb, new PointF(530.0F, 700.0F));
            e.Graphics.DrawString(CompName, new Font("宋体", 9.5F, FontStyle.Bold), 
                sb, new PointF(225.0F, 725.0F));
            e.Graphics.DrawString(dt.Year.ToString(), ft, sb, new PointF(180.0F, 895.0F));
            e.Graphics.DrawString(dt.Month.ToString(), ft, sb, new PointF(228.0F, 895.0F));
            e.Graphics.DrawString(dt.Day.ToString(), ft, sb, new PointF(262.0F, 895.0F));
            e.Graphics.DrawString(dt.Year.ToString(), ft, sb, new PointF(368.0F, 895.0F));
            e.Graphics.DrawString(dt.Month.ToString(), ft, sb, new PointF(415.0F, 895.0F));
            e.Graphics.DrawString(dt.Day.ToString(), ft, sb, new PointF(448.0F, 895.0F));
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            if(checkBox1.CheckState == CheckState.Checked)
            {
                taxtextBox7.Text = licencetextBox2.Text;
            }
            else if(checkBox1.CheckState == CheckState.Unchecked)
            {

            }
            else
            {
                MessageBox.Show("Are you kidding me?");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string file = Application.ExecutablePath;
            Configuration cm = ConfigurationManager.OpenExeConfiguration(file);
            cm.AppSettings.Settings.Add("AreaCode", "311700");
            cm.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        private void licencetextBox2_TextChanged(object sender, EventArgs e)
        {
            if(checkBox1.CheckState == CheckState.Checked)
            {
                taxtextBox7.Text = licencetextBox2.Text;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void mailFill_PrintPage(object sender, PrintPageEventArgs e)
        {
            DateTime dt = DateTime.Now;
            Font ft = new Font("宋体", 11, FontStyle.Bold);
            SolidBrush sb = new SolidBrush(Color.Black);
            string CompName = compNametextBox1.Text;
            string Addr = addrtextBox1.Text;
            string OwnerPhone = ownerPhonetextBox5.Text;
            string Postal = GetAppConfig("Postal");      //Postal---224221
            string OwnerName = ownertextBox3.Text;
            string OwnerID = ownerIDtextBox4.Text;
            string RelyName = relytextBox3.Text;
            string RelyID = relyIDtextBox2.Text;
            string RelyPhone = relyPhonetextBox1.Text;
            string BankShort = GetAppConfig("BankNameShort");
            e.Graphics.DrawString(CompName, ft, sb, new PointF(210.0F, 240.0F));
            e.Graphics.DrawString("存款", ft, sb, new PointF(596.0F, 240.0F));
            e.Graphics.DrawString(Addr, ft, sb, new PointF(210.0F, 280.0F));
            e.Graphics.DrawString(Postal, ft, sb, new PointF(596.0F, 280.0F));
            e.Graphics.DrawString(OwnerName, ft, sb, new PointF(210.0F, 316.0F));
            e.Graphics.DrawString(RelyName, ft, sb, new PointF(596.0F, 316.0F));
            e.Graphics.DrawString(OwnerID, ft, sb, new PointF(210.0F, 350.0F));
            e.Graphics.DrawString(RelyID, ft, sb, new PointF(596.0F, 350.0F));
            e.Graphics.DrawString(OwnerPhone, ft, sb, new PointF(210.0F, 386.0F));
            e.Graphics.DrawString(RelyPhone, ft, sb, new PointF(596.0F, 386.0F));
            e.Graphics.DrawString(CompName, ft, sb, new PointF(495.0F, 628.0F));
            e.Graphics.DrawString(BankShort, ft, sb, new PointF(667.0F, 843.0F));
            e.Graphics.DrawString(dt.Year.ToString(), ft, sb, new PointF(175.0F, 1000.0F));
            e.Graphics.DrawString(dt.Month.ToString(), ft, sb, new PointF(226.0F, 1000.0F));
            e.Graphics.DrawString(dt.Day.ToString(), ft, sb, new PointF(293.0F, 1000.0F));
            e.Graphics.DrawString(dt.Year.ToString(), ft, sb, new PointF(514.0F, 1000.0F));
            e.Graphics.DrawString(dt.Month.ToString(), ft, sb, new PointF(576.0F, 1000.0F));
            e.Graphics.DrawString(dt.Day.ToString(), ft, sb, new PointF(640.0F, 1000.0F));
        }

        private void checkBox2_CheckStateChanged(object sender, EventArgs e)
        {
            if(checkBox2.CheckState == CheckState.Checked)
            {
                checkBox3.Checked = false;
                relyIDtextBox2.Text = ownerIDtextBox4.Text;
                relytextBox3.Text = ownertextBox3.Text;
                relyPhonetextBox1.Text = ownerPhonetextBox5.Text;
            }
            else if(checkBox2.CheckState == CheckState.Unchecked)
            {
                checkBox3.Checked = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(mailFill_PrintPage);
            try
            {
                pd.Print();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void agreement_PrintPage(object sender, PrintPageEventArgs e)
        {
            Font ft = new Font("宋体", 10, FontStyle.Bold);
            SolidBrush sb = new SolidBrush(Color.Black);
            string CompName = compNametextBox1.Text;
            string Addr = addrtextBox1.Text;
            string OwnerPhone = ownerPhonetextBox5.Text;
            string Postal = GetAppConfig("Postal");      //Postal---224221
            string OwnerName = ownertextBox3.Text;
            string BankFull = GetAppConfig("BankNameFull");
            string hz = GetAppConfig("BankCharge");
            string BankPhone = GetAppConfig("BankPhone");
            string BankAddr = GetAppConfig("BankAddr");
            e.Graphics.DrawString(BankFull, ft, sb, new PointF(188.0F, 180.0F));
            e.Graphics.DrawString(hz, ft, sb, new PointF(245.0F, 199.0F));
            e.Graphics.DrawString(BankPhone, ft, sb, new PointF(510.0F, 199.0F));
            e.Graphics.DrawString(Postal, ft, sb, new PointF(190.0F, 219.0F));
            e.Graphics.DrawString(BankAddr, ft, sb, new PointF(395.0F, 219.0F));
            e.Graphics.DrawString(CompName, ft, sb, new PointF(188.0F, 245.0F));
            e.Graphics.DrawString(OwnerName, ft, sb, new PointF(245.0F, 265.0F));
            e.Graphics.DrawString(OwnerPhone, ft, sb, new PointF(510.0F, 265.0F));
            e.Graphics.DrawString(Postal, ft, sb, new PointF(190.0F, 285.0F));
            e.Graphics.DrawString(Addr, ft, sb, new PointF(395.0F, 285.0F));
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(agreement_PrintPage);
            try
            {
                pd.Print();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox4_CheckStateChanged(object sender, EventArgs e)
        {
            if(checkBox4.CheckState == CheckState.Unchecked)
            {
                comeIDtextBox2.Text = "";
                comeIDtextBox2.Enabled = false;
                cometextBox3.Text = "";
                cometextBox3.Enabled = false;
                comePhonetextBox1.Text = "";
                comePhonetextBox1.Enabled = false;
            }
            else if(checkBox1.CheckState == CheckState.Checked)
            {
                comeIDtextBox2.Enabled = true;
                cometextBox3.Enabled = true;
                comePhonetextBox1.Enabled = true;
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox3_CheckStateChanged(object sender, EventArgs e)
        {
            if(checkBox3.CheckState == CheckState.Checked)
            {
                checkBox2.Checked = false;
                relyIDtextBox2.Text = comeIDtextBox2.Text;
                relytextBox3.Text = cometextBox3.Text;
                relyPhonetextBox1.Text = comePhonetextBox1.Text;
            }
            else if(checkBox3.CheckState == CheckState.Unchecked)
            {
                checkBox2.Checked = true;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            compName = compNametextBox1.Text;
            licence = licencetextBox2.Text;
            addr = addrtextBox1.Text;
            scope = scopetextBox2.Text;
            tax = taxtextBox7.Text;
            money = moneytextBox4.Text;
            moneyType = moneyTypetextBox3.Text;
            onDate = ondatetextBox1.Text;
            thruDate = thrudatetextBox1.Text;
            genre = genrecomboBox1.Text;
            owner = ownertextBox3.Text;
            ownerID = ownerIDtextBox4.Text;
            ownerPhone = ownerPhonetextBox5.Text;
            basee = basetextBox1.Text;
            onDate = ondatetextBox1.Text;
            if(thrudatetextBox1.Text == "")
            {
                thruDate = "2099年01月01日";
            }
            else
            {
                thruDate = thrudatetextBox1.Text;
            }
            shenqinbiao sqb = new shenqinbiao();
            sqb.Show();
            
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            WordExcelFuck.toReplace.Update("aaa", "AAA");
            WordExcelFuck.toReplace.Update("bbb", "AAA");
            //WordExcelFuck.toReplace.Update("aaa", "d");
            MessageBox.Show(WordExcelFuck.toReplace["aaa"]);
        }

        private void ondatetextBox1_TextChanged(object sender, EventArgs e)
        {
            if(ondatetextBox1.Text.Length == 8)
            {
                ondatetextBox1.Text = string.Format("{0}年{1}月{2}日", ondatetextBox1.Text.Substring(0, 4),
                    ondatetextBox1.Text.Substring(4, 2), ondatetextBox1.Text.Substring(6, 2));
            }
        }

        private void thrudatetextBox1_TextChanged(object sender, EventArgs e)
        {
            if (thrudatetextBox1.Text.Length == 8)
            {
                thrudatetextBox1.Text = string.Format("{0}年{1}月{2}日", thrudatetextBox1.Text.Substring(0, 4),
                    thrudatetextBox1.Text.Substring(4, 2), thrudatetextBox1.Text.Substring(6, 2));
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            string currDir = Environment.CurrentDirectory;
            DirectoryInfo theDir = new DirectoryInfo(currDir);
            foreach(FileInfo fi in theDir.GetFiles())
            {
                if (fi.Name.EndsWith("_temp.doc") || fi.Name.EndsWith("_temp.xls"))
                {
                    fi.Delete();
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            WordExcelFuck.toReplace.Update("%CompanyName%", compNametextBox1.Text);
            WordExcelFuck.toReplace.Update("%DateInChinese%", dateinchn);
            #region switch开户
            switch (powercomboBox1.Text)
            {
                case "开":
                    WordExcelFuck.toReplace.Update("%Account%", "                       ");
                    WordExcelFuck.toReplace.Update("%Come1%", cometextBox3.Text.PadRight(7 - cometextBox3.Text.Length));
                    WordExcelFuck.toReplace.Update("%ComeID1%", comeIDtextBox2.Text);
                    WordExcelFuck.toReplace.Update("%Come2%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID2%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come3%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID3%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come4%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID4%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come5%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID5%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come6%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID6%", "".PadRight(18));
                    break;
                case "变":
                    WordExcelFuck.toReplace.Update("%Account%", accounttextBox1.Text);
                    WordExcelFuck.toReplace.Update("%Come2%", cometextBox3.Text.PadRight(7 - cometextBox3.Text.Length));
                    WordExcelFuck.toReplace.Update("%ComeID2%", comeIDtextBox2.Text);
                    WordExcelFuck.toReplace.Update("%Come1%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID1%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come3%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID3%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come4%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID4%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come5%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID5%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come6%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID6%", "".PadRight(18));
                    break;
                case "撤":
                    WordExcelFuck.toReplace.Update("%Account%", accounttextBox1.Text);
                    WordExcelFuck.toReplace.Update("%Come3%", cometextBox3.Text.PadRight(7 - cometextBox3.Text.Length));
                    WordExcelFuck.toReplace.Update("%ComeID3%", comeIDtextBox2.Text);
                    WordExcelFuck.toReplace.Update("%Come1%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID1%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come2%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID2%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come4%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID4%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come5%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID5%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come6%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID6%", "".PadRight(18));
                    break;
                case "网":
                    WordExcelFuck.toReplace.Update("%Account%", accounttextBox1.Text);
                    WordExcelFuck.toReplace.Update("%Come4%", cometextBox3.Text.PadRight(7 - cometextBox3.Text.Length));
                    WordExcelFuck.toReplace.Update("%ComeID4%", comeIDtextBox2.Text);
                    WordExcelFuck.toReplace.Update("%Come1%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID1%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come2%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID2%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come3%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID3%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come5%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID5%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come6%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID6%", "".PadRight(18));
                    break;
                case "印":
                    WordExcelFuck.toReplace.Update("%Account%", accounttextBox1.Text);
                    WordExcelFuck.toReplace.Update("%Come5%", cometextBox3.Text.PadRight(7 - cometextBox3.Text.Length));
                    WordExcelFuck.toReplace.Update("%ComeID5%", comeIDtextBox2.Text);
                    WordExcelFuck.toReplace.Update("%Come1%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID1%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come2%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID2%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come4%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID4%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come3%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID3%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come6%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID6%", "".PadRight(18));
                    break;
                default:
                    WordExcelFuck.toReplace.Update("%Account%", accounttextBox1.Text);
                    WordExcelFuck.toReplace.Update("%Come6%", cometextBox3.Text.PadRight(7 - cometextBox3.Text.Length));
                    WordExcelFuck.toReplace.Update("%ComeID6%", comeIDtextBox2.Text);
                    WordExcelFuck.toReplace.Update("%Come1%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID1%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come2%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID2%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come4%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID4%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come5%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID5%", "".PadRight(18));
                    WordExcelFuck.toReplace.Update("%Come3%", "".PadRight(7));
                    WordExcelFuck.toReplace.Update("%ComeID3%", "".PadRight(18));
                    break;
            }
            #endregion
            WordExcelFuck.genTemp("PowerTemplate.doc");
            string currDir = Environment.CurrentDirectory;
            string dest = Path.Combine(currDir, "PowerTemplate_temp.doc");
            WordExcelFuck.wReplaceNormal(WordExcelFuck.toReplace, dest);
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                Word.Application wa = new Word.ApplicationClass();
                object wb = wa.WordBasic;
                object[] argValues = new object[] { printDialog1.PrinterSettings.PrinterName, 1 };
                string[] argNames = new string[] { "Printer", "DoNotSetAsSysDefault" };
                wa.WordBasic.GetType().InvokeMember("FilePrintSetup", BindingFlags.InvokeMethod, null, wb, argValues, null, null, argNames);
                object missing = Missing.Value;
                Word.Document d = wa.Documents.Open(dest, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                d.PrintOut();
                d.Close(ref missing, ref missing, ref missing);
                wa.Quit(ref missing, ref missing, ref missing);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if(checkBox2.CheckState == CheckState.Unchecked)
            {
                WordExcelFuck.toReplace.Update("%BankNameShort%", GetAppConfig("BankNameShort"));
                WordExcelFuck.toReplace.Update("%DateInChinese%", dateinchn);
                WordExcelFuck.toReplace.Update("%RelyName%", relytextBox3.Text);
                WordExcelFuck.toReplace.Update("%RelyID%", relyIDtextBox2.Text);
                WordExcelFuck.toReplace.Update("%RelyPhone%", relyPhonetextBox1.Text);
                WordExcelFuck.toReplace.Update("%OwnerID%", ownerIDtextBox4.Text);
                WordExcelFuck.toReplace.Update("%OwnerPhone%", ownerPhonetextBox5.Text);
                WordExcelFuck.genTemp("MailTemplate.doc");
                string currDir = Environment.CurrentDirectory;
                string dest = Path.Combine(currDir, "MailTemplate_temp.doc");
                WordExcelFuck.wReplaceNormal(WordExcelFuck.toReplace, dest);
                if (printDialog1.ShowDialog() == DialogResult.OK)
                {
                    Word.Application wa = new Word.ApplicationClass();
                    object wb = wa.WordBasic;
                    object[] argValues = new object[] { printDialog1.PrinterSettings.PrinterName, 1 };
                    string[] argNames = new string[] { "Printer", "DoNotSetAsSysDefault" };
                    wa.WordBasic.GetType().InvokeMember("FilePrintSetup", BindingFlags.InvokeMethod, null, wb, argValues, null, null, argNames);
                    object missing = Missing.Value;
                    Word.Document d = wa.Documents.Open(dest, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    d.PrintOut();
                    d.Close(ref missing, ref missing, ref missing);
                    wa.Quit(ref missing, ref missing, ref missing);
                }
            }
            else
            {
                MessageBox.Show("同法人，不需要打印");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            WordExcelFuck.toReplace.Update("%BankNameFull%", GetAppConfig("BankNameFull"));
            WordExcelFuck.toReplace.Update("%DateInChinese%", dateinchn);
            WordExcelFuck.toReplace.Update("%RelyName%", relytextBox3.Text);
            WordExcelFuck.toReplace.Update("%RelyID%", relyIDtextBox2.Text);
            WordExcelFuck.genTemp("CodeLetterTemplate.doc");
            string currDir = Environment.CurrentDirectory;
            string dest = Path.Combine(currDir, "CodeLetterTemplate_temp.doc");
            WordExcelFuck.wReplaceNormal(WordExcelFuck.toReplace, dest);
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {
                Word.Application wa = new Word.ApplicationClass();
                object wb = wa.WordBasic;
                object[] argValues = new object[] { printDialog1.PrinterSettings.PrinterName, 1 };
                string[] argNames = new string[] { "Printer", "DoNotSetAsSysDefault" };
                wa.WordBasic.GetType().InvokeMember("FilePrintSetup", BindingFlags.InvokeMethod, null, wb, argValues, null, null, argNames);
                object missing = Missing.Value;
                Word.Document d = wa.Documents.Open(dest, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                d.PrintOut();
                d.Close(ref missing, ref missing, ref missing);
                wa.Quit(ref missing, ref missing, ref missing);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            WordExcelFuck.toReplace.Update("%CompanyName%", compNametextBox1.Text);
            WordExcelFuck.toReplace.Update("%Addr%", addrtextBox1.Text);
            WordExcelFuck.toReplace.Update("%Bureau%", "工商部门");
            WordExcelFuck.toReplace.Update("%Tax%", taxtextBox7.Text);
            WordExcelFuck.toReplace.Update("%Genre%", genrecomboBox1.Text);
            WordExcelFuck.toReplace.Update("%OnDate%", ondatetextBox1.Text);
            WordExcelFuck.toReplace.Update("%CompPhone%", ownerPhonetextBox5.Text);
            WordExcelFuck.toReplace.Update("%Money%", moneytextBox4.Text);
            WordExcelFuck.toReplace.Update("%OwnerID%", ownerIDtextBox4.Text);
            WordExcelFuck.toReplace.Update("%OwnerName%", ownertextBox3.Text);
            WordExcelFuck.genTemp("CodeTemplate.xls");
            string currDir = Environment.CurrentDirectory;
            string dest = Path.Combine(currDir, "CodeTemplate_temp.xls");
            WordExcelFuck.eReplace(WordExcelFuck.toReplace, dest);
            Excel.Application ea = new Excel.ApplicationClass();
            object missing = Missing.Value;
            Excel.Workbook wb = ea.Workbooks.Open(dest, missing, missing, missing, missing, missing, missing,
                 missing, missing, missing, missing, missing, missing, missing, missing);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets["Sheet1"];
            ea.Visible = true;
            ws.PrintPreview(true);
            wb.Saved = true;
            wb.Close(missing, missing, missing);
            ea.Visible = false;
            ea.Quit();
        }
    }
}
