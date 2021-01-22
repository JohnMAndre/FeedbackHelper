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

using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System.Xml;

namespace FeedbackHelper
{
    public partial class ribbonFeedback
    {
        public ribbonFeedback()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();

            // Load XML settings file
            // Note: This MUST be in the constructor (or grpConstructive.Items.Add(rcButton) will fail)
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(@"C:\Users\John\Documents\Visual Studio 2013\Projects\FeedbackHelper\FeedbackHelper\bin\Debug\FeedbackData.xml");
            XmlNodeList xList=xDoc.SelectNodes("//Item");
            XmlElement xElement;
            RibbonButton rcButton;
            foreach (XmlNode xNode in xList)
            {
                xElement=(XmlElement)xNode;

                rcButton = this.Factory.CreateRibbonButton();
                rcButton.Label = xElement.GetAttribute("Title");
                rcButton.Tag = xElement.InnerText;
                rcButton.Click += btn_Click;
                rcButton.SuperTip = xElement.GetAttribute("Tip");
                
                switch (xElement.GetAttribute("Type"))
                {
                    case "Constructive":
                        // Add to constructive section
                        grpConstructive.Items.Add(rcButton);
                        break;
                    case "Positive":
                        // Add to constructive section
                        grpPositive.Items.Add(rcButton);
                        break;
                }
            }
        }

        Document _oDocument;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            _oDocument = Globals.ThisAddIn.Application.ActiveDocument;
        }

        private void btn_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton rcButton = (RibbonButton)sender;
            AddComment(rcButton.Tag.ToString());
        }

        private void AddComment(string Text)
        {
            Selection curSel = Globals.ThisAddIn.Application.Selection;
            Comment cmt = curSel.Comments.Add(curSel.Range, Text);
            cmt.Edit();
        }

        private void btnDeleteComment_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Comments.Count>0)
                Globals.ThisAddIn.Application.Selection.Comments[1].Delete();
            
                
        }

    }
}
