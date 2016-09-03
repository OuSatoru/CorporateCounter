﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CorporateCounter
{
    public partial class shenqinbiao : Form
    {
        public shenqinbiao()
        {
            InitializeComponent();
        }

        

        private void shenqinbiao_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            WordExcelFuck.toReplace.Add("%CompanyName%", Form1.compName);
            WordExcelFuck.toReplace.Add("%CompCer%", "营业执照");
            WordExcelFuck.toReplace.Add("%Licence%", Form1.licence);
            WordExcelFuck.toReplace.Add("%Addr%", Form1.addr);
            WordExcelFuck.toReplace.Add("%CompPhone%", Form1.ownerPhone);
            WordExcelFuck.toReplace.Add("%Postal%", Form1.GetAppConfig("Postal"));
            WordExcelFuck.toReplace.Add("%Bureau%", "工商部门");
            WordExcelFuck.toReplace.Add("%MoneyType%", Form1.moneyType);
            WordExcelFuck.toReplace.Add("%Money%", Form1.money);
            WordExcelFuck.toReplace.Add("%Scope%", Form1.scope);
            WordExcelFuck.toReplace.Add("%Tax%", Form1.tax);
            WordExcelFuck.toReplace.Add("%Base%", Form1.basee);
            WordExcelFuck.toReplace.Add("Genre", Form1.genre);
            WordExcelFuck.toReplace.Add("%RelyCer%", "身份证");
            WordExcelFuck.toReplace.Add("%RelyID%", Form1.ownerID);
            WordExcelFuck.toReplace.Add("%RelyPhone%", Form1.ownerPhone);
            WordExcelFuck.toReplace.Add("%RelyName%", Form1.owner);
            WordExcelFuck.toReplace.Add("%OwnerCer%", "身份证");
            WordExcelFuck.toReplace.Add("%OwnerID%", Form1.ownerID);
            WordExcelFuck.toReplace.Add("%OwnerPhone%", Form1.ownerPhone);
            WordExcelFuck.toReplace.Add("%OwnerName%", Form1.owner);
            WordExcelFuck.toReplace.Add("OnDate", Form1.onDate);
            WordExcelFuck.toReplace.Add("ThruDate", Form1.thruDate);
            WordExcelFuck.toReplace.Add("OnCountry", textBox6.Text);
            WordExcelFuck.toReplace.Add("InCountry", textBox7.Text);
            WordExcelFuck.genTemp("10104Template.doc");
            string currDir = Environment.CurrentDirectory;
            string dest = Path.Combine(currDir, "10104Template_temp.doc");
            WordExcelFuck.wReplaceStory(WordExcelFuck.toReplace, dest);
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
    }
}
