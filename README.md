# NoteHighlight2016
![Alt text](/img/menu.png?raw=true "Menu")

[![GitHub release](https://img.shields.io/github/release/elvirbrk/NoteHighlight2016.svg)](https://github.com/elvirbrk/NoteHighlight2016/releases/tag/v3.7)[![Github Releases](https://img.shields.io/github/downloads/elvirbrk/NoteHighlight2016/latest/total.svg)](https://github.com/elvirbrk/NoteHighlight2016/releases/tag/v3.7)
[![Github previous](https://img.shields.io/github/downloads/elvirbrk/NoteHighlight2016/v3.6/total.svg)](https://github.com/elvirbrk/NoteHighlight2016/releases/tag/v3.6)
[![Github All Releases](https://img.shields.io/github/downloads/elvirbrk/NoteHighlight2016/total.svg)](https://github.com/elvirbrk/NoteHighlight2016/releases)

[![Follow @NoteHighlight](https://img.shields.io/twitter/follow/NoteHighlight.svg?style=social&label=Follow%20@NoteHighlight)](https://twitter.com/NoteHighlight?ref_src=twsrc%5Etfw)<br >
Follow on Twitter for updates and general questions. For bug reports and feature requests please use [Issues](https://github.com/elvirbrk/NoteHighlight2016/issues) page.

Based on NoteHighlight 2013 (https://notehighlight2013.codeplex.com) and VanillaAddin (https://github.com/OneNoteDev/VanillaAddIn) to create working addin for 64-bit OneNote 2016 and OneNote for O365

Syntax highlighting performed using https://gitlab.com/saalen/highlight (currently v4.19)

# Install
To install just run NoteHighlight2016.msi from [releases](https://github.com/elvirbrk/NoteHighlight2016/releases). Only 64-bit Office is supported.
See [here](https://support.office.com/en-us/article/About-Office-What-version-of-Office-am-I-using-932788B8-A3CE-44BF-BB09-E334518B8B19?ui=en-US&rs=en-US&ad=US) how to check which Office version you have.

In case AddIn doesn't show after install, check if this helps [Not showing after install](https://github.com/elvirbrk/NoteHighlight2016/issues/7)

# Usage
## Add new code
1. Select language from menu
2. Enter source code in pop-up window and press OK
3. Highlighted source code will show up in page

![Alt text](/img/usage.png?raw=true "Usage")

## Format existing text or edit formatted text
1. Select text that you want to format or edit (you can select whole or only part of note). In case part on note is already formatted, you can select whole source code box.
2. From NoteHighlight menu select desired language
3. NoteHighlight form will open with selected text
4. Edit text or change formatting same as for new code

# Additional languages

From v3.8 onwards, the set of languages shown on the ribbon is controlled per-user from the **Languages...** button in the NoteHighlight ribbon's Advanced group. Open the dialog to pin languages as large ribbon buttons (up to 14) and enable additional languages under the **More languages...** dropdown. Settings live in `%APPDATA%\NoteHighlight2016\languages.json` and are preserved across uninstall / upgrade. The ribbon refreshes on the next ribbon click after **Apply** or **OK**.

> **Migration note:** If you previously hand-edited `ribbon.xml` to enable extra languages, open **Languages...** and tick them again - that file is now treated as the canonical declaration of available languages and is no longer the place to opt languages in or out.

Adding wholly new languages (supported by the `highlight` tool but not yet declared in `ribbon.xml`) still requires editing `ribbon.xml` in the installation folder and is for advanced users.

# Sample of Themes
samples directory [Theme Samples](./img/Theme%20Samples/ThemeSample.md)