namespace NoteHighlightAddin.Preview
{
    partial class GalleryForm
    {
        /// <summary>
        /// Required designer variable. Disposal and the strict MSHTML teardown of
        /// the hosted <see cref="System.Windows.Forms.WebBrowser"/> are performed in
        /// the hand-written <see cref="Dispose(bool)"/> in <c>GalleryForm.cs</c>;
        /// this Designer file deliberately does NOT declare a second override.
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
            this.progressPanel = new System.Windows.Forms.Panel();
            this.lblProgress = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.progressPanel.SuspendLayout();
            this.SuspendLayout();
            //
            // browser
            //
            this.browser.AllowWebBrowserDrop = false;
            this.browser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.browser.IsWebBrowserContextMenuEnabled = false;
            this.browser.Location = new System.Drawing.Point(0, 40);
            this.browser.MinimumSize = new System.Drawing.Size(20, 20);
            this.browser.Name = "browser";
            this.browser.ScriptErrorsSuppressed = true;
            this.browser.Size = new System.Drawing.Size(900, 620);
            this.browser.TabIndex = 1;
            this.browser.WebBrowserShortcutsEnabled = false;
            //
            // progressPanel
            //
            this.progressPanel.Controls.Add(this.lblProgress);
            this.progressPanel.Controls.Add(this.progressBar);
            this.progressPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.progressPanel.Location = new System.Drawing.Point(0, 0);
            this.progressPanel.Name = "progressPanel";
            this.progressPanel.Padding = new System.Windows.Forms.Padding(8, 8, 8, 8);
            this.progressPanel.Size = new System.Drawing.Size(900, 40);
            this.progressPanel.TabIndex = 0;
            //
            // lblProgress
            //
            this.lblProgress.AutoSize = true;
            this.lblProgress.Dock = System.Windows.Forms.DockStyle.Left;
            this.lblProgress.Location = new System.Drawing.Point(8, 8);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Padding = new System.Windows.Forms.Padding(0, 4, 8, 0);
            this.lblProgress.Size = new System.Drawing.Size(120, 21);
            this.lblProgress.TabIndex = 0;
            this.lblProgress.Text = "Rendering themes...";
            this.lblProgress.AccessibleName = "Gallery render progress";
            this.lblProgress.AccessibleRole = System.Windows.Forms.AccessibleRole.StatusBar;
            //
            // progressBar
            //
            this.progressBar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.progressBar.Location = new System.Drawing.Point(128, 8);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(764, 24);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar.TabIndex = 1;
            //
            // GalleryForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(900, 660);
            this.Controls.Add(this.browser);
            this.Controls.Add(this.progressPanel);
            this.MinimumSize = new System.Drawing.Size(480, 320);
            this.Name = "GalleryForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Browse all themes";
            this.AccessibleName = "Theme gallery";
            this.AccessibleRole = System.Windows.Forms.AccessibleRole.Dialog;
            this.progressPanel.ResumeLayout(false);
            this.progressPanel.PerformLayout();
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.WebBrowser browser;
        private System.Windows.Forms.Panel progressPanel;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.ProgressBar progressBar;
    }
}
