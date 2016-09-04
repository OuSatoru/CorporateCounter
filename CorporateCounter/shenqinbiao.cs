using System;
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
            WordExcelFuck.toReplace.Update("%CompanyName%", Form1.compName);
            WordExcelFuck.toReplace.Update("%CompCer%", "营业执照");
            WordExcelFuck.toReplace.Update("%Licence%", Form1.licence);
            WordExcelFuck.toReplace.Update("%Addr%", Form1.addr);
            WordExcelFuck.toReplace.Update("%CompPhone%", Form1.ownerPhone);
            WordExcelFuck.toReplace.Update("%Postal%", Form1.GetAppConfig("Postal"));
            WordExcelFuck.toReplace.Update("%Bureau%", "工商部门");
            WordExcelFuck.toReplace.Update("%MoneyType%", Form1.moneyType);
            WordExcelFuck.toReplace.Update("%Money%", Form1.money);
            WordExcelFuck.toReplace.Update("%Scope%", Form1.scope);
            WordExcelFuck.toReplace.Update("%Tax%", Form1.tax);
            WordExcelFuck.toReplace.Update("%Base%", Form1.basee);
            WordExcelFuck.toReplace.Update("%Genre%", Form1.genre);
            WordExcelFuck.toReplace.Update("%RelyCer%", "身份证");
            WordExcelFuck.toReplace.Update("%RelyID%", Form1.ownerID);
            WordExcelFuck.toReplace.Update("%RelyPhone%", Form1.ownerPhone);
            WordExcelFuck.toReplace.Update("%RelyName%", Form1.owner);
            WordExcelFuck.toReplace.Update("%OwnerCer%", "身份证");
            WordExcelFuck.toReplace.Update("%OwnerID%", Form1.ownerID);
            WordExcelFuck.toReplace.Update("%OwnerPhone%", Form1.ownerPhone);
            WordExcelFuck.toReplace.Update("%OwnerName%", Form1.owner);
            WordExcelFuck.toReplace.Update("%OnDate%", Form1.onDate);
            WordExcelFuck.toReplace.Update("%ThruDate%", Form1.thruDate);
            WordExcelFuck.toReplace.Update("%OnCountry%", textBox6.Text);
            WordExcelFuck.toReplace.Update("%InCountry%", textBox7.Text);
            WordExcelFuck.genTemp("10104Template.doc");
            string currDir = Environment.CurrentDirectory;
            string dest = Path.Combine(currDir, "10104Template_temp.doc");
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
    }
}
