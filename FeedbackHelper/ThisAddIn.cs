using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace FeedbackHelper
{
    public partial class ThisAddIn
    {
        public Word.Document m_doc;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //this.Application.DocumentBeforeSave +=
            //   new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);

            //this.Application.DocumentOpen += Application_DocumentOpen;

        }

        //void Application_DocumentOpen(Word.Document Doc)
        //{
        //    m_doc = Doc;
        //}
        //public void InsertText(string Text)
        //{
        //    m_doc.Paragraphs[1].Range.InsertParagraphBefore();
        //    m_doc.Paragraphs[1].Range.Text = Text;
        //}

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            //Doc.Paragraphs[1].Range.InsertParagraphBefore();
            //Doc.Paragraphs[1].Range.Text = "This text was added by using code.";
        }

        

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
