# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

LECmd (Lnk Explorer Command line edition) is a single-project .NET console tool that parses Windows shortcut (`.lnk`) files and displays them or exports them to CSV, JSON, XML and XHTML. It's part of Eric Zimmerman's DFIR tool suite. The parsing itself lives in separate NuGet packages that come from sibling repos: `Lnk` (EricZimmerman/Lnk), `ExtensionBlocks` and `GuidMapping`. LECmd only handles the CLI, display and export. A fix to how a LNK structure is parsed usually belongs in the `Lnk` repo, followed by a package version bump in `LECmd/LECmd.csproj`.

## Commands

The SDK is pinned in `global.json` (.NET 10).

Debug info: `Directory.Build.props` sets `DebugType embedded` (the pdbs are inside the assemblies: stack traces and Exceptionless reports have file and line numbers, and no .pdb file ships) and, for Release, `ContinuousIntegrationBuild` (source paths in the pdbs are `/_/...`, not this machine's folders). Don't add `DebugType` or `DebugSymbols` to the project files.

```bash
dotnet build -c Release                               # builds LECmd.slnx, both net472 and net10.0
dotnet build LECmd/LECmd.csproj -f net10.0            # one target only
dotnet run --project LECmd -f net10.0 -- -f path/to/file.lnk
dotnet run --project LECmd -f net10.0 -- -d path/to/dir --csv out --json out --xml out/xml --html out -q
```

There is no test project or lint config. CI (`.github/workflows/build.yml`) only builds both targets. To verify a change, run the old and new binaries over a corpus of LNK files, such as `../lnk/ITA_LNKs` and `../lnk/Lnk.Test/TestFiles` if the Lnk repo is checked out next to this one, and diff the CSV and JSON output. Some files in TestFiles are deliberately malformed: they are listed under Failed files, and the run still exits 0. Exit codes are 0 or 1 only: 1 is a failure of the tool (invalid arguments, a missing `-f`/`-d`, a file it can't open or read (`_readErrors`: access denied, locked, IO errors), a failed export, or a fatal error); damaged lnk files exit 0.

## Architecture

Almost all of the code is in `LECmd/Program.cs`:

- **`Main`** builds a `System.CommandLine` `RootCommand` and awaits `InvokeAsync`, which calls **`DoWork`**. When you add an option, you need to (1) create the `Option<T>`, (2) add it to the `RootCommand` initializer and (3) pass `result.GetValue(...)` into `DoWork`. Missing step 2 has caused a bug before.
- **`DoWork`** validates the arguments, enumerates the files, calls `ProcessFile` on each one and collects successfully parsed `LnkFile`s in `_processedFiles`. It then runs a single export pass over all of them. Each record is converted with `GetCsvFormat` inside its own `try`, so one bad record doesn't abort the whole export. The resulting `CsvOut` row is written to CSV and XHTML, and to XML through `SaveXml`. JSON is the same `CsvOut` row with its times in `JsonDateFormat` (ISO 8601 UTC with a Z, as in JLECmd), serialized with ServiceStack.Text; `ConfigureJsonDates` at startup does the same for any `DateTime`/`DateTimeOffset` ServiceStack serializes.
- **`ProcessFile`** parses a file with `Lnk`. **`DisplayLnk`** does the console output. A file is counted as failed only when parsing fails. An exception thrown while displaying a file must not mark it as failed.
- **`CsvOut`** defines the CSV, XML and HTML columns. Keep existing columns in the same order and add new ones at the end, because Timeline Explorer layouts and users' scripts depend on the column positions.
- **Embedded resources** (`Properties/Resources.resx`): `MACs.txt` (the vendor lookup behind the tracker MAC address) and `normalize.css`/`style.css`, which are copied next to `index.xhtml` for HTML output.

### Dual-target conditional code

The project targets `net472` and `net10.0`. `#if NET9_0_OR_GREATER` blocks separate the two:
- **net472** uses AlphaFS (`Alphaleonis.Win32.Filesystem`) aliased as `File`, `Directory` and `Path`, and enumerates files with `DirectoryEnumerationFilters`.
- **net10.0** uses `System.IO` and `FileSystemEnumerable` with predicates that include only non-empty files and don't recurse into reparse points.

When you change file handling, keep the two branches behaving the same way, and build both targets.

### Build details

- Costura.Fody (`FodyWeavers.xml`) embeds the dependencies, so the result is a single exe.
- `LangVersion` is pinned to 10.
- Exceptionless telemetry starts at the top of `Main`.

## Versioning and dependencies

- Unlike the library repos (Lnk, ExtensionBlocks, GuidMapping), which use MinVer and publish to NuGet from a GitHub release, LECmd is not packable. Its version is the `<Version>` in `LECmd/LECmd.csproj` (calendar style, e.g. `2026.10.0`), bumped by hand for each release. The banner shows it.
- The dependency order is GuidMapping → ExtensionBlocks → Lnk → LECmd. LECmd references all three directly, so keep them in step with the versions Lnk uses.
- Since Lnk 2026.10.6 the shell items (`lnk.TargetIDs`) are ExtensionBlocks' classes (`IShellBag`, `ShellBag0X31`, ...; Lnk's own are gone, except `Lnk.ShellItems.ShellBagDamaged`). The `--all` details switch on those types (`ShellBag0X31`/`0X35`/`0X36`/`0X39`/`0XB1`/`0X3A` file entries, `ShellBagCFSF` for CFSF items, `ShellBag0X41`-`0X47` network items, ...). The target's MFT entry and sequence number come from the last 0xbeef0004 of `TargetBlocks` (a 0x00 property view's `DelegateExtensionBlocks`, then its own blocks). Checked when it moved: the CSV over 2,494 LNK files was the same.
- Dependabot (`.github/dependabot.yml`) opens monthly PRs for GitHub Actions and NuGet updates.
- Release builds are staged, signed and zipped by `Build-ZimmermanTools.ps1` in the get-zimmermantools repository (cloned next to this one). Its program table says what goes in this tool's download zip (the files taken from the build, folders copied from the repositories, files kept from the current release zip); a new output file, a content folder or a change in how the tool is built (single file, another framework) needs that entry updated.
