namespace NoteHighlightAddin
{
    partial class SettingsForm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.btnFont = new System.Windows.Forms.Button();
            this.cbShowTableBorder = new System.Windows.Forms.CheckBox();
            this.btnTogglePreview = new System.Windows.Forms.Button();
            this.splitContainer = new System.Windows.Forms.SplitContainer();
            this._previewPane = new NoteHighlightAddin.Preview.PreviewPane();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).BeginInit();
            this.splitContainer.Panel1.SuspendLayout();
            this.splitContainer.Panel2.SuspendLayout();
            this.splitContainer.SuspendLayout();
            this.SuspendLayout();
            //
            // fontDialog1
            //
            this.fontDialog1.AllowScriptChange = false;
            this.fontDialog1.AllowSimulations = false;
            this.fontDialog1.AllowVerticalFonts = false;
            this.fontDialog1.FontMustExist = true;
            this.fontDialog1.ShowEffects = false;
            //
            // btnFont
            //
            this.btnFont.Location = new System.Drawing.Point(34, 20);
            this.btnFont.Name = "btnFont";
            this.btnFont.Size = new System.Drawing.Size(339, 23);
            this.btnFont.TabIndex = 1;
            this.btnFont.Text = "Font";
            this.btnFont.UseVisualStyleBackColor = true;
            this.btnFont.Click += new System.EventHandler(this.BtnFont_Click);
            //
            // cbShowTableBorder
            //
            this.cbShowTableBorder.AutoSize = true;
            this.cbShowTableBorder.Location = new System.Drawing.Point(34, 70);
            this.cbShowTableBorder.Name = "cbShowTableBorder";
            this.cbShowTableBorder.Size = new System.Drawing.Size(117, 17);
            this.cbShowTableBorder.TabIndex = 2;
            this.cbShowTableBorder.Text = "Show Table Border";
            this.cbShowTableBorder.UseVisualStyleBackColor = true;
            this.cbShowTableBorder.CheckedChanged += new System.EventHandler(this.ChShowTableBorder_CheckedChanged);
            //
            // btnTogglePreview
            //
            this.btnTogglePreview.AccessibleName = "Toggle preview pane";
            this.btnTogglePreview.Location = new System.Drawing.Point(34, 110);
            this.btnTogglePreview.Name = "btnTogglePreview";
            this.btnTogglePreview.Size = new System.Drawing.Size(120, 25);
            this.btnTogglePreview.TabIndex = 3;
            this.btnTogglePreview.Text = "Hide preview";
            this.btnTogglePreview.UseVisualStyleBackColor = true;
            this.btnTogglePreview.Click += new System.EventHandler(this.btnTogglePreview_Click);
            //
            // splitContainer
            //
            this.splitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer.Location = new System.Drawing.Point(0, 0);
            this.splitContainer.Name = "splitContainer";
            this.splitContainer.Orientation = System.Windows.Forms.Orientation.Vertical;
            //
            // splitContainer.Panel1
            //
            this.splitContainer.Panel1.Controls.Add(this.btnTogglePreview);
            this.splitContainer.Panel1.Controls.Add(this.cbShowTableBorder);
            this.splitContainer.Panel1.Controls.Add(this.btnFont);
            this.splitContainer.Panel1MinSize = 200;
            this.splitContainer.Panel1.TabIndex = 0;
            //
            // splitContainer.Panel2
            //
            this.splitContainer.Panel2.Controls.Add(this._previewPane);
            this.splitContainer.Panel2MinSize = 200;
            this.splitContainer.Panel2.TabIndex = 1;
            this.splitContainer.Size = new System.Drawing.Size(408, 355);
            this.splitContainer.SplitterDistance = 204;
            this.splitContainer.TabIndex = 0;
            //
            // _previewPane
            //
            this._previewPane.Dock = System.Windows.Forms.DockStyle.Fill;
            this._previewPane.Location = new System.Drawing.Point(0, 0);
            this._previewPane.Name = "_previewPane";
            this._previewPane.QuickStyleSuppress = false;
            this._previewPane.TabIndex = 0;
            //
            // SettingsForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(408, 355);
            this.Controls.Add(this.splitContainer);
            this.Name = "SettingsForm";
            this.Text = "SettingsForm";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SettingsForm_FormClosed);
            this.Shown += new System.EventHandler(this.SettingsForm_Shown);
            this.splitContainer.Panel1.ResumeLayout(false);
            this.splitContainer.Panel1.PerformLayout();
            this.splitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).EndInit();
            this.splitContainer.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.Button btnFont;
        private System.Windows.Forms.CheckBox cbShowTableBorder;
        private System.Windows.Forms.Button btnTogglePreview;
        private System.Windows.Forms.SplitContainer splitContainer;
        private NoteHighlightAddin.Preview.PreviewPane _previewPane;
    }
}
