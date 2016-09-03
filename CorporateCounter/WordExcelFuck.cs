using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;

namespace CorporateCounter
{
    class WordExcelFuck
    {
        public static Dictionary<string, string> toReplace = new Dictionary<string, string>()
        {
            {"%Year%", DateTime.Now.Year.ToString()},
            {"%Month", DateTime.Now.Month.ToString()},
            {"Day", DateTime.Now.Day.ToString()},
        };
        
        public static void genTemp(string src)
        {
            string currDir = Environment.CurrentDirectory;
            string srcFull = Path.Combine(currDir, src);
            string dest = Path.Combine(currDir, Path.GetFileNameWithoutExtension(srcFull) + "_temp" + Path.GetExtension(srcFull));
            if (File.Exists(srcFull))
            {
                File.Copy(srcFull, dest, true);
            }
        }


        public static void wReplaceStory(Dictionary<string,string> dic, string dest)
        {
            Word.Application wa = new Word.ApplicationClass();
            object missing = Missing.Value;
            object replace;
            replace = Word.WdReplace.wdReplaceAll;
            Word.Document d = wa.Documents.Open(dest, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            Word.StoryRanges sr = d.StoryRanges;
            //foreach (var item in dic)
            //{
                foreach (Word.Range rg in sr)
                {
                    rg.Find.ClearFormatting();
                        rg.Find.Replacement.ClearFormatting();
                        rg.Find.Text = "%Year%";
                        rg.Find.Replacement.Text = "2016";
                        rg.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, 
                                ref missing, ref missing, ref missing, ref missing, ref missing, ref replace,
                                ref missing, ref missing, ref missing, ref missing);
                        
                }
                d.Save();
            //}
            d.Close(ref missing, ref missing, ref missing);
            wa.Quit(ref missing, ref missing, ref missing);
        }

        public static void wReplaceNomal(Dictionary<string, string> dic, string dest)
        {
            Word.Application wa = new Word.ApplicationClass();
            object missing = Missing.Value;
            object replace;
            replace = Word.WdReplace.wdReplaceAll;
            Word.Document d = wa.Documents.Open(dest, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            foreach(var item in dic)
            {
                object findtext, replacewith;
                d.Content.Find.Text = item.Key;
                findtext = item.Key;
                replacewith = item.Value;
                d.Content.Find.ClearFormatting();
                d.Content.Find.Execute(ref findtext, ref missing, ref missing, ref missing, ref missing, 
                    ref missing, ref missing, ref missing, ref missing, replacewith, replace, ref missing, 
                    ref missing, ref missing, ref missing);
                d.Save();
            }
            d.Close(ref missing, ref missing, ref missing);
            wa.Quit(ref missing, ref missing, ref missing);
        }

        public static void eReplace(Dictionary<string, string> dic, string dest)
        {

        }

        public static void wPrint(string dest)
        {

        }

        public static void ePrint(string dest)
        {

        }
    }
}
