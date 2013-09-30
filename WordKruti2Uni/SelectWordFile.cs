using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
namespace WordKruti2Uni
{
    public partial class SelectWordFile : Form
    {
        public SelectWordFile()
        {
            InitializeComponent();
        }

        private void SelectWordFile_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                string filename = openFileDialog1.FileName;
                textBox1.Text = filename;
                //string savefilename = filename + "_unicode";


            }
            Console.WriteLine(result); // <-- For debugging use.
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            Kruti2uni kr = new Kruti2uni();
             Document doc;
             Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
             try
             {
                 doc = app.Documents.Open(this.textBox1.Text, Visible:true);
                 string filename = System.IO.Path.GetFileNameWithoutExtension(this.textBox1.Text);
                 string directory = System.IO.Path.GetDirectoryName(this.textBox1.Text);
                 WordParse.wordParse(doc, kr);
                // doc.Save();
                 String str = directory + "\\" + filename + "_unicode" + ".doc";
                   doc.SaveAs(str);
                 doc.Close();
                 label4.Text = "Output saved at " + str;
                 //app.Quit();
             }
             catch (Exception ex) { MessageBox.Show("Error in execution " + ex.Message);  }
            /*
             try
             {
                 Characters chr = doc.Characters;
                 int i = 1;
                 bool v = Regex.Match("v", @"\s").Success;
                 int counter = 1;
                 while (i < chr.Count)
                 {
                     while (chr[i].Font.Name != "Kruti Dev 010") { i++;  }
                     int j = counter;
                     int k = i;
                     String str = "";
                     while (!Regex.Match(chr[i].Text, @"\s").Success && chr[i].Font.Name == "Kruti Dev 010")
                     {
                         i++;
                         str += chr[i].Text;
                     }
                     this.label2.Text = str;
                    // this.label2.Font.Name = "Kruti Dev 010";
                     while (Regex.Match(chr[k].Text, @"\s+").Success) i++;
                     Range rg = doc.Range(i, k);
                     String newstr = kr.converter(str);
                     rg.Text = newstr;
                     /*
                     rg.Find.Text = str;
                     rg.Find.Replacement.ClearFormatting();
                     rg.Find.Replacement.Text = newstr;
                     rg.Find.Replacement.Font.Name = "Mangal";
                     rg.Find.Execute();
                      
                     counter = i + newstr.Length - str.Length;

                     i++;





                 }
             }
             catch (Exception ex) { }
    */
            /*
             Words wd = doc.Words;
             for (int i = 0; i < wd.Count; i++)
             {
                 if (wd[i + 1].Font.Name == "Kruti Dev 010")
                 {
                     doc.Words[i + 1].Text = kr.converter(wd[i+1].Text);
                     doc.Words[i + 1].Font.Name = "Mangal";
                    // doc.Words[i + 1].Font.Size = (doc.Words[i + 1].Font.Size) / 2;
                 }
             }
             * */
            // doc.Range().Text = kr.converter(doc.Range().Text);
             //doc.Range().Font.Name = "Mangal";
   
            
        }

        private void requirementsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
