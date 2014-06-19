using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace TestFileFormation
{
    class FileFormator
    {
        Object Nothing = System.Reflection.Missing.Value;
        public string savePath;
        Application wordApp;
        Document Doc;
        public FileFormator()
        {
            wordApp = new ApplicationClass();
        }

        public void CreateWordDoc(float leftMargin = 80f, float rightMargin = 80f, float pageWid = 600f)
        {
            Doc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            wordApp.Selection.PageSetup.LeftMargin = leftMargin;
            wordApp.Selection.PageSetup.RightMargin = rightMargin;
            wordApp.Selection.PageSetup.PageWidth = pageWid;
            wordApp.Selection.ParagraphFormat.LineSpacing = 50f;
        }

        public void InsertHeader()
        {
    

        }

        public void NextLine(int lines = 1)
        {
            object WdLine;
            while (lines != 0)
            {
                WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;
                lines--;
            }
        }

        public void InsertText(string textContent)
        {
            wordApp.Selection.Text = textContent; // Cotent
            wordApp.Selection.Range.Bold = 2; // If bold
            wordApp.Selection.Range.Font.Size = 30; // text size
            wordApp.Selection.Range.Font.TextColor.RGB = 95; // text color
            wordApp.Selection.Range.Font.Name = "Georgia"; // text type
        }

        public void AppendText(string textContent)
        {
            wordApp.Selection.Text += textContent;
            wordApp.Selection.Range.Font.Name = "Aharoni"; // text type

        }

        public void SaveDoc(object fileName)
        {
            Doc.SaveAs2(ref fileName);
            Doc.Close(ref Nothing, ref Nothing, ref Nothing);
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
        }

        
        
    }
}
