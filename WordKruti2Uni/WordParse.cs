using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace WordKruti2Uni
{
    class WordParse
    {
        
        public static void wordParse(Document doc, Kruti2uni kr)
        {
            //Range rg = doc.Range();
            Selection sl = doc.Application.Selection;
           // sl1.Find.ClearFormatting();
            //sl.Find.MatchWildcards = true;
           // sl1.Find.Font.Name = "Kruti Dev 010";
           // rg.Find.Replacement.ClearFormatting();
           // sl.Find.Text = @"[ \t]";
           
            bool result = true;
            try
            {
                while (result)
                {
                  //  Selection sl = doc.Application.Selection;
                   // sl = doc.Application.Selection;
                    sl.Find.ClearFormatting();
                    sl.Find.Font.Name = "Kruti Dev 010";
                    // http://social.msdn.microsoft.com/Forums/windows/en-US/82b56ef5-ae05-4fda-a963-25a026c559a4/wordselectionfindexecute-search-and-replace-only-work-on-the-first-search-but-not-on
                    sl.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;

                   result= sl.Find.Execute();
                  
                   if (result)
                   {
                       sl.Text = kr.converter(sl.Text);
                       sl.Font.Name = "Mangal";
                       sl.Font.Size = sl.Font.Size ;
                       //sl.Next().Select(); ;
                       
                   }
                }
            }
            catch (Exception ex) { }
           
           
        }
        public static void wordParsePara(Document doc, Kruti2uni kr)
        {
            int count = doc.Paragraphs.Count;

            for (int i = 1; i <= count; i++)
            {
                Range rg = doc.Paragraphs[i].Range;
                rg.Find.ClearFormatting();
               // rg.Find.MatchWildcards = true;
                rg.Find.Font.Name = "Kruti Dev 010";
                rg.Find.Execute();
                rg.Text = kr.converter(rg.Text);
            }
        }
       
    }
}
