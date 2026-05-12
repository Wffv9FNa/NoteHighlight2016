# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

NoteHighlight2016 is a COM add-in for 64-bit OneNote 2016 / OneNote for O365 that inserts syntax-highlighted source code into a OneNote page. Highlighting is performed by shelling out to the bundled `highlight.exe` (Andre Simon's `highlight` tool, currently v4.19, https://gitlab.com/saalen/highlight) and the resulting HTML is converted into the OneNote XML schema and pushed back into the active page via the OneNote COM API.

The codebase is C# targeting .NET Framework 4.5, built with Visual Studio 2015+ (solution format VS14). It is Windows-only.

## Build / Run / Test

Open `NoteHighlight2016.sln` in Visual Studio. The solution contains four projects:

- `NoteHighlightAddin` - the COM add-in DLL (the entry point; output is `Library`).
- `GenerateHighlightContent` - wraps invocation of `highlight.exe` and parses the `HighLightSection` from `App.config`.
- `Helper` - small utilities (`ProcessHelper`, `StringExtension.TemplateSubstitute`).
- `UnitTesting` - MSTest project.
- `Setup` - Visual Studio Installer (`.vdproj`) project producing the x64 MSI. **`.vdproj` requires the "Microsoft Visual Studio Installer Projects" extension to load/build.**

Important platform notes:

- Only the **x64** configuration is supported and shipped; the x86 MSI was dropped (the `NoteHighlightAddin` csproj retains its `x86` platform configurations for diagnostic use, but no x86 release artefact is built).
- `NoteHighlightAddin` registers as COM via `RegisterForComInterop` (only enabled in `Debug|AnyCPU`). For other configurations, COM registration is performed by the MSI.
- The `Release|Any CPU` mapping for `NoteHighlightAddin` actually builds `Release|x86` (see `.sln`); there is no AnyCPU output for the add-in itself.

There are no command-line `dotnet` build instructions - this is a classic `.csproj` solution. Use `msbuild NoteHighlight2016.sln /p:Configuration=Release /p:Platform=x64` from a Developer Command Prompt if building outside the IDE.

Tests live in `UnitTesting/UnitTesting.cs` (MSTest, `Microsoft.VisualStudio.TestTools.UnitTesting`). Run via Visual Studio Test Explorer or `vstest.console.exe UnitTesting\bin\Debug\UnitTesting.dll`.

## Architecture

### Execution flow when a user clicks a language button

1. **OneNote loads `NoteHighlightAddin.dll`** as a COM add-in (`AddIn` class implements `IDTExtensibility2` + `IRibbonExtensibility`, GUID `4C6B0362-F139-417F-9661-3663C268B9E9`, ProgID `NoteHighlight2016.AddIn`).
2. `AddIn.GetCustomUI` reads `ribbon.xml` (loaded from the assembly directory at runtime, **not** as an embedded resource - `ribbon.xml` must be deployed alongside the DLL). Each language button has a `tag` attribute that maps to a `highlight` syntax key (e.g. `cs`, `py`, `ps1`).
3. On button click, `AddIn.AddInButtonClicked` spawns a new STA thread running `ShowForm`, which:
   - Reads the currently viewed page via `OneNoteApplication.GetHierarchy(...)` and `GetPageContent(...)`.
   - Captures any selected text (handles plain selections, partial selections inside a `T` element, and selections inside a `Table` cell - see `GetSelectedText` and `IsSelectedTextInline`).
   - Opens `MainForm` (a WinForms editor backed by `ICSharpCode.TextEditor`) prefilled with the selection.
4. When the user clicks OK in `MainForm`, `GenerateHighlightContent.GenerateHighLight.GenerateHighLightCode` writes the source to a temp file, builds command-line args from `App.config`'s `<HighLightSection>`, and runs `highlight\highlight.exe` to produce HTML.
5. `AddIn.InsertHighLightCodeToCurrentSide` calls `PrepareFormatedContent` to translate the highlighter HTML into a OneNote `<Table>` (one cell for line numbers, one for code) under the OneNote XML namespace, then calls `OneNoteApplication.UpdatePageContent` to commit. If the user had a selection, the original `Outline` is mutated in place; otherwise a new `Outline` is created at the captured mouse position.

### `highlight.exe` invocation is config-driven

`NoteHighlightAddin/App.config` (and `GenerateHighlightContent/App.config`) define `<HighLightSection>` with `<GeneralArguments>` and `<OutputArguments>` collections. Each `<add>` row contributes a CLI flag; values can contain placeholder tokens (`{inputFileName}`, `{codeType}`, `{font}`, etc.) that `Helper.StringExtension.TemplateSubstitute` fills in at runtime. To change how `highlight.exe` is called, edit `App.config` rather than the C# code. The bundled `highlight\` folder (binary, `langDefs\`, `themes\`, runtime DLLs) is copied to the output via `<Content CopyToOutputDirectory="Always">` entries in `NoteHighlightAddin.csproj` - new themes or language definitions added there will be picked up automatically.

### Adding a new language to the ribbon

Add a `<button>` in `NoteHighlightAddin/ribbon.xml` with `onAction="AddInButtonClicked"`, `tag="<highlight syntax key>"`, and `image="<png in Resources\>"`. End users can also enable the many `visible="false"` buttons already declared by editing `ribbon.xml` in the install folder (no rebuild required - see README).

### Settings

User settings live under `NoteHighLightForm.Properties.Settings` (`ShowLineNumber`, `DarkMode`, `QuickStyle`, `Font`, `FontSize`, `HighLightStyle`, `BackgroundColor`, `ShowTableBorder`, `SaveOnClipboard`). They are read/written from both `AddIn.cs` ribbon callbacks and `SettingsForm`. `DarkMode` causes `PrepareFormatedContent` to strip the `background-color` style from the `<pre>` tag so OneNote dark theme renders correctly.

### Quirks worth knowing

- `PrepareFormatedContent` injects `lang=la` into each `<pre>` so OneNote disables spell-check on code.
- After updating page content, an `Indent` overflow workaround zeroes any `indent` attribute > 1,000,000 (OneNote occasionally emits absurd values on partial selections).
- `ContainsAsianCharacter` adds `spaceBefore`/`spaceAfter` (configurable via `AsianBeforeSpace` / `AsianAfterSpace` in `App.config`) when CJK characters are detected, to fix line-height issues.
- Forms are launched on a dedicated STA thread via `Application.Run` so they don't block OneNote's main thread - do not assume single-threaded UI ownership.
