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
            // PreviewPane
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.browser);
            this.Name = "PreviewPane";
            this.Size = new System.Drawing.Size(400, 300);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.WebBrowser browser;
    }
}
