# LECmd

Parses Windows shortcut (`.lnk`) files: the target's paths and time stamps, the target ID list (shell items), the
volume and network share it was on, the machine and MAC address that created it (tracker data), and every extra data
block. Results are shown on the console and can be exported to CSV, JSON, XML and HTML.

| | |
|---|---|
| **Version** | 2026.10.1 |
| **Download** | [.NET 10](https://download.ericzimmermanstools.com/net/LECmd.zip) (recommended) &middot; [.NET Framework 4.7.2](https://download.ericzimmermanstools.com/LECmd.zip) |
| **All tools** | [ericzimmerman.github.io](https://ericzimmerman.github.io) or [Get-ZimmermanTools](https://download.ericzimmermanstools.com/Get-ZimmermanTools.zip), which downloads them all and keeps them current |
| **Changes** | [ChangeLog.txt](https://download.ericzimmermanstools.com/ChangeLog.txt) |
| **Issues** | [EricZimmerman/LECmd](https://github.com/EricZimmerman/LECmd/issues) or [EricZimmerman/Issues](https://github.com/EricZimmerman/Issues/issues) |

## Requirements

- **.NET 10 build**: the [.NET 10 runtime](https://dotnet.microsoft.com/download/dotnet/10.0). Run `LECmd.exe` on Windows,
  or `dotnet LECmd.dll` on any system with the runtime.
- **.NET Framework build**: .NET Framework 4.7.2 or later (Windows only).

## Command line

```text
Usage:
  LECmd [options]

Options:
  -f <f>          File to process. Either this or -d is required
  -d <d>          Directory to recursively process. Either this or -f is required
  -r              Only process lnk files pointing to removable drives
  --all           Process all files in directory vs. only files matching *.lnk
  --csv <csv>     Directory to save CSV formatted results to. Be sure to include the full path in double quotes
  --csvf <csvf>   File name to save CSV formatted results to. When present, overrides default name
  --json <json>   Directory to save JSON formatted results to. Be sure to include the full path in double quotes. Use --pretty for a more human readable layout
  --xml <xml>     Directory to save XML formatted results to. Be sure to include the full path in double quotes
  --html <html>   Directory to save xhtml formatted results to. Be sure to include the full path in double quotes
  --pretty        When exporting to json, use a more human readable layout
  -q              Do not dump full details about each file processed. Speeds up processing when using --json or --csv
  --nid           Suppress Target ID list details from being displayed
  --neb           Suppress Extra blocks information from being displayed
  --dt <dt>       The custom date/time format to use when displaying time stamps. See https://goo.gl/CNVq0k for options [default: yyyy-MM-dd HH:mm:ss]
  --mp            When true, display higher precision for timestamps
  --cp <cp>       Code page to parse strings [default: 1252]
  --debug         Show debug information during processing
  --trace         Show trace information during processing
  -?, -h, --help  Show help and usage information
  --version       Show version information
```

Short options (a single letter) start with one dash, long options with two.

## Examples

```text
LECmd.exe -f "C:\Temp\foobar.lnk"
LECmd.exe -f "C:\Temp\somelink.lnk" --json "D:\jsonOutput" --pretty
LECmd.exe -d "C:\Temp" --csv "c:\temp" --html c:\temp --xml c:\temp\xml -q
LECmd.exe -f "C:\Temp\some other link.lnk" --nid --neb
LECmd.exe -d "C:\Temp" --all
```

## Output

- **Console**: every field of each shortcut, including the target ID list and extra blocks (`--nid` and `--neb` hide
  them; `-q` shows only the file names).
- **CSV** (`--csv`): one row per shortcut. `TargetIDAbsolutePath` is the full path built from the target ID list, with
  folders' real names where Windows also stores a localized one. The last columns are `Name`, `IconLocation`,
  `ShowWindow`, `EnvironmentVariables`, the four tracker droids (`VolumeDroid`, `FileDroid`, `VolumeDroidBirth`,
  `FileDroidBirth`) and `Anomalies`.
- **JSON** (`--json`): one object per shortcut, time stamps in ISO 8601 UTC (`2024-01-02T03:04:05.6780000Z`).
- **XML** (`--xml`) and **HTML** (`--html`): one file per shortcut. XML file names include a hash of the source path, so
  same-named shortcuts from different folders don't overwrite each other.

Damaged shortcuts are parsed as far as possible: damaged shell items, parse problems and anomalies (data after the end
of the file, out of range time stamps, whitespace padded arguments, undefined show commands) are reported, and
shortcuts damaged by a text mode copy (a CR byte before every LF, as an ASCII mode FTP transfer does) are named as such.

## Exit codes

| Code | Meaning |
|---|---|
| `0` | Success. Damaged shortcuts are reported but are not a failure |
| `1` | The tool failed: invalid arguments, a missing `-f` or `-d` path, output that can't be written, or files or folders it can't read. Folders under `-d` that can't be read are reported (`Unable to read folder <path>: <reason>`) and the rest is still processed |

## What's new in 2026.10

- About 27% faster with less memory.
- New columns in every export: `Name`, `IconLocation`, `ShowWindow`, `EnvironmentVariables`, the tracker droids and
  `Anomalies`.
- Real folder names in `TargetIDAbsolutePath` instead of resource strings such as `@shell32.dll,-21781`; folders inside
  archives (Windows 11), document drive items and 0x3D folders are read; OneDrive and Creative Cloud folders are named
  whatever their account.
- Name and Arguments are no longer cut at 260 characters; Unicode volume labels and network share paths are correct.
- One bad shortcut no longer stops an export; JSON output replaces an existing file instead of leaving old data at its
  end; relative `--html` paths work.
- Exit code 1 when the tool fails (it was always 0). Error reports include source file and line numbers.

The full list is in the [ChangeLog](https://download.ericzimmermanstools.com/ChangeLog.txt).

## Building

```text
dotnet build LECmd.slnx -c Release
```

Builds `net10.0` and `net472`. Shortcut parsing is in the [Lnk](https://github.com/EricZimmerman/Lnk) library
([NuGet](https://www.nuget.org/packages/Lnk)).

## Further reading

- [Introducing LECmd!](https://binaryforay.blogspot.com/2016/02/introducing-lecmd.html)
- [LECmd v0.6.0.0 released!](https://binaryforay.blogspot.com/2016/02/lecmd-v0600-released.html)
- [PECmd, LECmd, and JLECmd updated!](https://binaryforay.blogspot.com/2016/03/pecmd-lecmd-and-jlecmd-updated.html)
- [LECmd and JLECmd updated](https://binaryforay.blogspot.com/2016/04/lecmd-and-jlecmd-updated.html)

## Special thanks

Open source development funding and support provided by the [SANS Institute](https://www.sans.org/) and
[SANS DFIR](https://www.sans.org/digital-forensics-incident-response/).

## License

[MIT](LICENSE), Copyright (c) Eric R. Zimmerman 2027.
