//Copyright 2017-2021 John M Andre (John At JohnMAndre dot COM)

//This file is part of Feedback Helper.

//Feedback Helper is free software: you can redistribute it And/Or modify
//it under the terms Of the GNU General Public License As published by
//the Free Software Foundation, either version 3 Of the License, Or
//(at your option) any later version.

//Feedback Helper is distributed In the hope that it will be useful,
//but WITHOUT ANY WARRANTY; without even the implied warranty of
//MERCHANTABILITY Or FITNESS FOR A PARTICULAR PURPOSE.  See the
//GNU General Public License For more details.

//You should have received a copy Of the GNU General Public License
//along with Feedback Helper.  If Not, see < https: //www.gnu.org/licenses/>.

using Word = Microsoft.Office.Interop.Word;

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
