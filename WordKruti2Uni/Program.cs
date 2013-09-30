using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;

namespace WordKruti2Uni
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //Kruti2uni kr = new Kruti2uni();
           // Document doc;
            SelectWordFile swf = new SelectWordFile();
            swf.ShowDialog();
            /*
            Application app = new Application();
            doc = app.Documents.Open("d:/test.doc",Visible:true);
             doc.Range().Text = kr.converter(doc.Range().Text);
             doc.Range().Font.Name = "Mangal";
             doc.Save();
             doc.Close();
             * */
            /*
            Find findObject = app.Selection.Find;
            findObject.ClearFormatting();
            findObject.MatchWildcards=true;
            findObject.Text = "<(.)>";
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text =kr.converter(@"\1");
            findObject.Execute();
            doc.Close();
             * */
           // object replaceAll = Word.WdReplace.wdReplaceAll;
           
            
        }
    }
}
