using System;
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

        

        private void button1_Click(object sender, EventArgs e)
        {
            PrintDocument pd = new PrintDocument();
            pd.PrintPage += new PrintPageEventHandler(this.requisition_PrintPage);
            pd.Print();
        }

        private void requisition_PrintPage(object sender, PrintPageEventArgs e)
        {
            Font ft = new Font("宋体", 11, FontStyle.Bold);
            SolidBrush sb = new SolidBrush(Color.Black);
            PointF pf1 = new PointF(150.0F, 150.0F);
            PointF pf2 = new PointF(150.0F, 300.0F);
            e.Graphics.DrawString("aaa", ft, sb, pf1);
            e.Graphics.DrawString("bbb", ft, sb, pf2);
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
    }
}
