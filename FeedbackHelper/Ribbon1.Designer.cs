namespace FeedbackHelper
{
    partial class ribbonFeedback : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

       

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab2 = this.Factory.CreateRibbonTab();
            this.grpConstructive = this.Factory.CreateRibbonGroup();
            this.grpPositive = this.Factory.CreateRibbonGroup();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnDeleteComment = this.Factory.CreateRibbonButton();
            this.tab2.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.grpConstructive);
            this.tab2.Groups.Add(this.grpPositive);
            this.tab2.Groups.Add(this.group1);
            this.tab2.Label = "Feedback";
            this.tab2.Name = "tab2";
            // 
            // grpConstructive
            // 
            this.grpConstructive.Label = "Constructive";
            this.grpConstructive.Name = "grpConstructive";
            // 
            // grpPositive
            // 
            this.grpPositive.Label = "Positive";
            this.grpPositive.Name = "grpPositive";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnDeleteComment);
            this.group1.Label = "Support";
            this.group1.Name = "group1";
            // 
            // btnDeleteComment
            // 
            this.btnDeleteComment.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDeleteComment.Description = "Delete selected comment";
            this.btnDeleteComment.Image = global::FeedbackHelper.Properties.Resources.DeleteComment;
            this.btnDeleteComment.Label = "Delete";
            this.btnDeleteComment.Name = "btnDeleteComment";
            this.btnDeleteComment.ShowImage = true;
            this.btnDeleteComment.SuperTip = "Delete selected comment";
            this.btnDeleteComment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteComment_Click);
            // 
            // ribbonFeedback
            // 
            this.Name = "ribbonFeedback";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpConstructive;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpPositive;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteComment;
    }

    partial class ThisRibbonCollection
    {
        internal ribbonFeedback Ribbon1
        {
            get { return this.GetRibbon<ribbonFeedback>(); }
        }
    }
}
