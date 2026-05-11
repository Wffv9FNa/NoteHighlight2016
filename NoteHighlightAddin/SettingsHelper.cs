namespace NoteHighlightAddin
{
    /// <summary>
    /// Process-wide serialisation point for writes to <c>user.config</c> via
    /// <c>Properties.Settings.Default.Save()</c>.
    ///
    /// <see cref="System.Configuration.ApplicationSettingsBase.Save"/> is not thread-safe; concurrent
    /// writes from the main STA and the worker STAs can truncate or corrupt the user-scoped
    /// configuration file. Every save site in the add-in must funnel through <see cref="SafeSave"/>
    /// (or take <see cref="SaveLock"/> explicitly) so the writes are serialised across the process.
    /// </summary>
    public static class SettingsHelper
    {
        public static readonly object SaveLock = new object();

        public static void SafeSave()
        {
            lock (SaveLock)
            {
                // The active user-scoped settings file in this add-in is the one declared in
                // NoteHighlightForm.Properties (see Properties\Settings.Designer.cs). Every existing
                // call site uses NoteHighlightForm.Properties.Settings.Default; this helper must
                // funnel writes through that same instance so the process-wide lock actually
                // serialises real saves rather than a parallel-but-unused settings object.
                NoteHighlightForm.Properties.Settings.Default.Save();
            }
        }
    }
}
