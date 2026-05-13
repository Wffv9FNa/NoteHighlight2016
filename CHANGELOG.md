# Changelog

All notable changes to NoteHighlight2016 are documented in this file. The format loosely follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## Unreleased

### Added
- **Language picker** in the ribbon's Advanced group (new **Languages...** button, keytip `L`). Choose which languages appear as pinned ribbon buttons (up to 14) and which appear under the new **More languages...** dropdown. Settings persist to `%APPDATA%\NoteHighlight2016\languages.json` and are preserved across uninstall / upgrade.

### Changed
- **Default pinned-languages list refreshed.** New defaults: C#, SQL, Python, JavaScript, HTML, XML, Java, CSS, PowerShell, Bash, JSON, Markdown. PHP, Perl, Ruby, and C++ have been dropped from the default pinned set. Existing users who previously had those visible via `visible="true"` in `ribbon.xml` can re-pin them via **Languages...**
- The ribbon now refreshes on the **next ribbon click** after Apply / OK in the Languages dialog (deferred invalidate), so changes appear without restarting OneNote.

### Fixed
- **MSI upgrade no longer ships a stale `ribbon.xml`.** `ribbon.xml` is now embedded inside the versioned `NoteHighlightAddin.dll` and read from there at runtime, so MSI upgrades over an existing install pick up new ribbon buttons / menus without needing `REINSTALLMODE=vamus` or a manual uninstall. A sibling `ribbon.xml` next to the DLL, if present, is still preferred (dev hot-edit convenience).

### Migration
If you previously hand-edited `ribbon.xml` in the installation folder to enable extra languages, open **Languages...** and tick them again - `ribbon.xml` is now the canonical declaration of available languages, but per-user visibility lives in `languages.json`.
