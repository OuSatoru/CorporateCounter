﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Configuration;

namespace CorporateCounter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

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
            pd.Print();
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
                relyIDtextBox2.Text = ownerIDtextBox4.Text;
                relytextBox3.Text = ownertextBox3.Text;
                relyPhonetextBox1.Text = ownerPhonetextBox5.Text;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(mailFill_PrintPage);
            pd.Print();
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
            pd.Print();
        }
    }
}
