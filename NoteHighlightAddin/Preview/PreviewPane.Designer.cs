namespace NoteHighlightAddin.Preview
{
    partial class PreviewPane
    {
        /// <summary>
        /// Required designer variable. Disposal is performed in the
        /// hand-written <see cref="Dispose(bool)"/> in <c>PreviewPane.cs</c>,
        /// which also tears down the hosted MSHTML <see cref="System.Windows.Forms.WebBrowser"/>
        /// in a strict order. This Designer file deliberately does NOT declare
        /// a second <c>Dispose</c> override.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.browser = new System.Windows.Forms.WebBrowser();
            this.lblRendering = new System.Windows.Forms.Label();
            this.SuspendLayout();
            //
            // browser
            //
            this.browser.AllowWebBrowserDrop = false;
            this.browser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.browser.IsWebBrowserContextMenuEnabled = false;
            this.browser.Location = new System.Drawing.Point(0, 0);
            this.browser.MinimumSize = new System.Drawing.Size(20, 20);
            this.browser.Name = "browser";
            this.browser.ScriptErrorsSuppressed = true;
            this.browser.Size = new System.Drawing.Size(400, 300);
            this.browser.TabIndex = 0;
            this.browser.WebBrowserShortcutsEnabled = false;
            //
            // lblRendering
            //
            this.lblRendering.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblRendering.AutoSize = true;
            this.lblRendering.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.lblRendering.ForeColor = System.Drawing.Color.DimGray;
            this.lblRendering.Location = new System.Drawing.Point(326, 6);
            this.lblRendering.Name = "lblRendering";
            this.lblRendering.Padding = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.lblRendering.Size = new System.Drawing.Size(68, 17);
            this.lblRendering.TabIndex = 1;
            this.lblRendering.Text = "rendering...";
            this.lblRendering.Visible = false;
            this.lblRendering.AccessibleName = "Preview rendering in progress";
            this.lblRendering.AccessibleRole = System.Windows.Forms.AccessibleRole.StatusBar;
            //
            // PreviewPane
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblRendering);
            this.Controls.Add(this.browser);
            this.Name = "PreviewPane";
            this.AccessibleName = "Highlight preview";
            this.AccessibleRole = System.Windows.Forms.AccessibleRole.Pane;
            this.Size = new System.Drawing.Size(400, 300);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.WebBrowser browser;
        private System.Windows.Forms.Label lblRendering;
    }
}
