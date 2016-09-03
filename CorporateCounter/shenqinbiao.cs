using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

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
            WordExcelFuck.toReplace.Add("%Addr", Form1.addr);
            WordExcelFuck.toReplace.Add("%CompPhone", Form1.ownerPhone);
            WordExcelFuck.toReplace.Add("%Postal%", Form1.GetAppConfig("Postal"));
            WordExcelFuck.toReplace.Add("Bureau", "工商部门");
            WordExcelFuck.toReplace.Add("MoneyType", Form1.moneyType);
            WordExcelFuck.toReplace.Add("Money", Form1.money);
            WordExcelFuck.toReplace.Add("%Scope%", Form1.scope);
            WordExcelFuck.toReplace.Add("Tax", Form1.tax);
            WordExcelFuck.toReplace.Add("Base", Form1.basee);
        }
    }
}
