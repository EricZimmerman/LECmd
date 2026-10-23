using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Help;
using System.CommandLine.Invocation;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
#if NET9_0_OR_GREATER
using System.IO.Enumeration;
#endif
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
#if !NET9_0_OR_GREATER
using Alphaleonis.Win32.Filesystem;
using Directory = Alphaleonis.Win32.Filesystem.Directory;
using File = Alphaleonis.Win32.Filesystem.File;
using Path = Alphaleonis.Win32.Filesystem.Path;
#else
using Path = System.IO.Path;
using Directory = System.IO.Directory;
using File = System.IO.File;
#endif

using Exceptionless;
using ExtensionBlocks;
using Lnk;
using Lnk.ExtraData;
using Lnk.ShellItems;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using ServiceStack;
using ServiceStack.Text;
using CsvWriter = CsvHelper.CsvWriter;
using Resources = LECmd.Properties.Resources;
using VolumeInfo = Lnk.VolumeInfo;

namespace LECmd;

internal class Program
{
    private static readonly string PreciseTimeFormat = "yyyy-MM-dd HH:mm:ss.fffffff";

    private static string _activeDateTimeFormat;

    private static RootCommand _rootCommand;

    private static readonly string Header = $"LECmd version {Assembly.GetExecutingAssembly().GetName().Version.ToString(3)}" +
                                            "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
                                            "\r\nhttps://github.com/EricZimmerman/LECmd";

    private static readonly string Footer = @"Examples: LECmd.exe -f ""C:\Temp\foobar.lnk""" + "\r\n\t " +
                                            @"   LECmd.exe -f ""C:\Temp\somelink.lnk"" --json ""D:\jsonOutput"" --pretty" +
                                            "\r\n\t " +
                                            @"   LECmd.exe -d ""C:\Temp"" --csv ""c:\temp"" --html c:\temp --xml c:\temp\xml -q" +
                                            "\r\n\t " +
                                            @"   LECmd.exe -f ""C:\Temp\some other link.lnk"" --nid --neb " +
                                            "\r\n\t " +
                                            @"   LECmd.exe -d ""C:\Temp"" --all" + "\r\n\t" +
                                            "\r\n\t" +
                                            "    Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes\r\n";

    private static List<string> _failedFiles;

    //files that could not be opened or read (a damaged file is only listed in _failedFiles)
    private static int _readErrors;

    private static readonly Dictionary<string, string> MacList = new Dictionary<string, string>();

    private static List<LnkFile> _processedFiles;

    private static void LoadMacs()
    {
        var lines = Resources.MACs.Split(new[] {Environment.NewLine}, StringSplitOptions.None);

        foreach (var line in lines)
        {
            if (line.Trim().Length == 0)
            {
                continue;
            }

            var segs = line.ToUpperInvariant().Split('\t');
            var key = segs[0].Trim();
            var val = segs[1].Trim();

            if (MacList.ContainsKey(key) == false)
            {
                MacList.Add(key, val);
            }
        }
    }

    private static bool IsAdministrator()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return true;
        }

        var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    // JSON time stamps: ISO 8601 in UTC (2024-01-02T03:04:05.6780000Z), as JLECmd writes them, instead of ServiceStack's \/Date(ms)\/
    private const string JsonDateFormat = "yyyy-MM-ddTHH:mm:ss.fffffffZ";

    private static void ConfigureJsonDates()
    {
        JsConfig.DateHandler = DateHandler.ISO8601;
        JsConfig<DateTimeOffset>.SerializeFn = JsonDate;
        JsConfig<DateTimeOffset?>.SerializeFn = d => d.HasValue ? JsonDate(d.Value) : null;
        JsConfig<DateTime>.SerializeFn = JsonDate;
        JsConfig<DateTime?>.SerializeFn = d => d.HasValue ? JsonDate(d.Value) : null;
    }

    private static string JsonDate(DateTimeOffset d)
    {
        return d.UtcDateTime.ToString(JsonDateFormat, CultureInfo.InvariantCulture);
    }

    private static string JsonDate(DateTime d)
    {
        //timestamps without a kind are already UTC; ToUniversalTime would treat them as local time
        if (d.Kind == DateTimeKind.Unspecified)
        {
            d = DateTime.SpecifyKind(d, DateTimeKind.Utc);
        }

        return d.ToUniversalTime().ToString(JsonDateFormat, CultureInfo.InvariantCulture);
    }

    private static async Task<int> Main(string[] args)
    {
        ConfigureJsonDates();

        ExceptionlessClient.Default.Startup("FNWfyFuaUAPnVfofTZAhZOgeDG5lv7AnjYKNtsEJ");

        LoadMacs();
        
            var fOpt = new Option<string>("-f")
            {
                Description = "File to process. Either this or -d is required"
            };

            var dOpt = new Option<string>("-d")
            {
                Description = "Directory to recursively process. Either this or -f is required"
            };
            
            var rOpt = new Option<bool>("-r")
            {
                Description =
                    "Only process lnk files pointing to removable drives",
                DefaultValueFactory = _ => false
            };

            var allOpt = new Option<bool>("--all")
            {
                Description =
                    "Process all files in directory vs. only files matching *.lnk",
                DefaultValueFactory = _ => false
            };

            var csvOpt = new Option<string>(
                "--csv")
            {
                Description =
                    "Directory to save CSV formatted results to. Be sure to include the full path in double quotes"
            };

            var csvfOpt = new Option<string>(
                "--csvf")
            {
                Description = "File name to save CSV formatted results to. When present, overrides default name"
            };
            var jsonOpt = new Option<string>(
                "--json")
            {
                Description =
                    "Directory to save JSON formatted results to. Be sure to include the full path in double quotes. Use --pretty for a more human readable layout"
            };
            
            var xmlOpt = new Option<string>(
                "--xml")
            {
                Description =
                    "Directory to save XML formatted results to. Be sure to include the full path in double quotes"
            };

       

            var htmlOpt = new Option<string>(
                "--html")
            {
                Description =
                    "Directory to save xhtml formatted results to. Be sure to include the full path in double quotes"
            };

            var prettyOpt = new Option<bool>("--pretty")
            {
                Description = "When exporting to json, use a more human readable layout",
                DefaultValueFactory = _ => false
            };

            var qOpt = new Option<bool>("-q")
            {
                Description =
                    "Do not dump full details about each file processed. Speeds up processing when using --json or --csv",
                DefaultValueFactory = _ => false
            };
            
            var nidOpt = new Option<bool>("--nid")
            {
                Description =
                    "Suppress Target ID list details from being displayed",
                DefaultValueFactory = _ => false
            };
            
            var nebOpt = new Option<bool>("--neb")
            {
                Description =
                    "Suppress Extra blocks information from being displayed",
                DefaultValueFactory = _ => false
            };
        

            var dtOpt = new Option<string>(
                "--dt")
            {
                Description =
                    "The custom date/time format to use when displaying time stamps. See https://goo.gl/CNVq0k for options",
                DefaultValueFactory = _ => "yyyy-MM-dd HH:mm:ss"
            };

            var mpOpt = new Option<bool>("--mp")
            {
                Description = "When true, display higher precision for timestamps",
                DefaultValueFactory = _ => false
            };

            var cpOpt = new Option<int>("--cp")
            {
                Description = "Code page to parse strings",
                DefaultValueFactory = _ => 1252
            };

            var debugOpt = new Option<bool>("--debug")
            {
                Description = "Show debug information during processing",
                DefaultValueFactory = _ => false
            };
            var traceOpt = new Option<bool>("--trace")
            {
                Description = "Show trace information during processing",
                DefaultValueFactory = _ => false
            };

        _rootCommand = new RootCommand
        {
            fOpt,
            dOpt,
            rOpt,
            allOpt,
            csvOpt,
            csvfOpt,
            jsonOpt,
            xmlOpt,
            htmlOpt,
            prettyOpt,
            qOpt,
            nidOpt,
            nebOpt,
            dtOpt,
            mpOpt,
            cpOpt,
            debugOpt,
            traceOpt
        };

        _rootCommand.Description = Header + "\r\n\r\n" + Footer;

        _rootCommand.SetAction(result => DoWork(result.GetValue(fOpt), result.GetValue(dOpt), result.GetValue(rOpt),
            result.GetValue(qOpt), result.GetValue(allOpt), result.GetValue(csvOpt), result.GetValue(csvfOpt),
            result.GetValue(xmlOpt),result.GetValue(htmlOpt),result.GetValue(jsonOpt),result.GetValue(prettyOpt),result.GetValue(nidOpt),result.GetValue(nebOpt),result.GetValue(dtOpt),result.GetValue(mpOpt),result.GetValue(debugOpt),result.GetValue(traceOpt),result.GetValue(cpOpt)));


        // Exit code: 0 = success, 1 = bad arguments, a fatal error, or one or more files that could not be parsed.
        var exitCode = await _rootCommand.Parse(args).InvokeAsync();

        Log.CloseAndFlush();

        return exitCode;
    }

    class DateTimeOffsetFormatter : IFormatProvider, ICustomFormatter
    {
        private readonly IFormatProvider _innerFormatProvider;

        public DateTimeOffsetFormatter(IFormatProvider innerFormatProvider)
        {
            _innerFormatProvider = innerFormatProvider;
        }

        public object GetFormat(Type formatType)
        {
            return formatType == typeof(ICustomFormatter) ? this : _innerFormatProvider.GetFormat(formatType);
        }

        public string Format(string format, object arg, IFormatProvider formatProvider)
        {
            if (arg is DateTimeOffset)
            {
                var size = (DateTimeOffset) arg;
                return size.ToString(_activeDateTimeFormat);
            }

            var formattable = arg as IFormattable;
            if (formattable != null)
            {
                return formattable.ToString(format, _innerFormatProvider);
            }

            return arg.ToString();
        }
    }

    private static int DoWork(string f, string d, bool r, bool q, bool all, string csv, string csvf, string xml,
        string html, string json, bool pretty, bool nid, bool neb, string dt, bool mp, bool debug, bool trace, int cp)
    {
        _activeDateTimeFormat = dt;
        if (mp)
        {
            _activeDateTimeFormat = PreciseTimeFormat;
        }

        var formatter =
            new DateTimeOffsetFormatter(CultureInfo.CurrentCulture);

        var levelSwitch = new LoggingLevelSwitch();

        var template = "{Message:lj}{NewLine}{Exception}";

        if (debug)
        {
            levelSwitch.MinimumLevel = LogEventLevel.Debug;
            template = "[{Timestamp:HH:mm:ss.fff} {Level:u3}] {Message:lj}{NewLine}{Exception}";
        }

        if (trace)
        {
            levelSwitch.MinimumLevel = LogEventLevel.Verbose;
            template = "[{Timestamp:HH:mm:ss.fff} {Level:u3}] {Message:lj}{NewLine}{Exception}";
        }

        var conf = new LoggerConfiguration()
            .WriteTo.Console(outputTemplate: template, formatProvider: formatter)
            .MinimumLevel.ControlledBy(levelSwitch);

        Log.Logger = conf.CreateLogger();

        if (f.IsNullOrEmpty() && d.IsNullOrEmpty())
        {
            var aaa = new CustomHelpAction(new HelpAction());
            aaa.Invoke(_rootCommand.Parse("Either -f or -d is required. Exiting"));

            Console.WriteLine();
            return 1;
        }

        if (f.IsNullOrEmpty() == false && !File.Exists(f))
        {
            var aaa = new CustomHelpAction(new HelpAction());
            aaa.Invoke(_rootCommand.Parse($"File {f} not found. Exiting"));

            Console.WriteLine();
            return 1;
        }

        if (d.IsNullOrEmpty() == false &&
            !Directory.Exists(d))
        {
            var aaa = new CustomHelpAction(new HelpAction());
            aaa.Invoke(_rootCommand.Parse($"Directory {d} not found. Exiting"));

            Console.WriteLine();
            return 1;
        }


        Log.Information("{Header}",Header);
        Console.WriteLine();
        Log.Information("Command line: {Args}",string.Join(" ", Environment.GetCommandLineArgs().Skip(1)));
        Console.WriteLine();

        if (IsAdministrator() == false)
        {
            Log.Warning($"Warning: Administrator privileges not found!");
            Console.WriteLine();
            Console.WriteLine();
        }

        _processedFiles = new List<LnkFile>();

        _failedFiles = new List<string>();
        _readErrors = 0;

        var tsNow = DateTimeOffset.UtcNow;

        if (f?.Length > 0)
        {
            LnkFile lnk;

            try
            {
                f = Path.GetFullPath(f);

                lnk = ProcessFile(f, q, r, dt, nid, neb, cp);
                if (lnk != null)
                {
                    _processedFiles.Add(lnk);
                }
            }
            catch (UnauthorizedAccessException ua)
            {
                Log.Error(ua,
                    "Unable to access {F}. Are you running as an administrator? Error: {Message}",f,ua.Message);
                return 1;
            }
            catch (Exception ex)
            {
                Log.Error(ex,
                    "Error processing file {F} Please send it to saericzimmerman@gmail.com. Error: {Message}",f,ex.Message);
                return 1;
            }
        }
        else
        {
            Log.Information("Looking for lnk files in {D}",d);
            Console.WriteLine();

            d = Path.GetFullPath(d);

            string[] lnkFiles;

            try
            {
                var mask = "*.lnk";
                if (all)
                {
                    mask = "*";
                }

                IEnumerable<string> files2;

#if NET9_0_OR_GREATER
                // Files only (EnumerateFileSystemEntries also returned directories), skip empty files like the
                // legacy path does, and do not recurse into junctions/symlinks, which can loop on mounted images.
                // Reparse-point *files* (e.g. cloud placeholders) are still included.
                // Folders that can't be read are reported and counted (IgnoreInaccessible skipped them silently).
                files2 = FindLnkFiles(d, mask);

#else 
// Legacy implementation for previous frameworks

                        var directoryEnumerationFilters = new DirectoryEnumerationFilters();
                        directoryEnumerationFilters.InclusionFilter = fsei =>
                        {
                            if (fsei.FileSize == 0)
                            {
                                return false;
                            }
                        
                            mask = ".lnk".ToUpperInvariant();
                            if (all)
                            {
                                mask = "*";
                            }
                        
                            if (mask == "*")
                            {
                                return true;
                            }
                        
                            if (fsei.Extension.ToUpperInvariant() == mask)
                            {
                                return true;
                            }
                        
                            return false;
                        };
                        
                        directoryEnumerationFilters.RecursionFilter =
 entryInfo => !entryInfo.IsMountPoint && !entryInfo.IsSymbolicLink;
                        
                        directoryEnumerationFilters.ErrorFilter = UnreadableFolder;
                        
                        var dirEnumOptions =
                            DirectoryEnumerationOptions.Files | DirectoryEnumerationOptions.Recursive |
                            DirectoryEnumerationOptions.SkipReparsePoints |
                            DirectoryEnumerationOptions.BasicSearch ;

                        files2 =
                            Alphaleonis.Win32.Filesystem.Directory.EnumerateFileSystemEntries(d, dirEnumOptions, directoryEnumerationFilters);

                        // files2 =
                        //    Directory.EnumerateFileSystemEntries(d, mask,SearchOption.AllDirectories);


#endif


                lnkFiles = files2.ToArray();

//                    lnkFiles = Directory.GetFiles(_fluentCommandLineParser.Object.Directory, mask,
//                        SearchOption.AllDirectories);
            }
            catch (UnauthorizedAccessException ua)
            {
                Log.Error(ua,
                    "Unable to access {D}. Error message: {Message}",d,ua.Message);
                return 1;
            }
            catch (Exception ex)
            {
                Log.Error(ex,
                    "Error getting lnk files in {D}. Error: {Message}",d,ex.Message);
                return 1;
            }

            Log.Information("Found {Length:N0} files",lnkFiles.Length);
            Console.WriteLine();

            var sw = new Stopwatch();
            sw.Start();

            foreach (var file in lnkFiles)
            {
                var lnk = ProcessFile(file, q, r, dt, nid, neb, cp);
                if (lnk != null)
                {
                    _processedFiles.Add(lnk);
                }
            }

            sw.Stop();

            if (q)
            {
               Console.WriteLine();
            }

            Log.Information(
                "Processed {ProcessedCount:N0} out of {TotalFiles:N0} files in {TotalSeconds:N4} seconds",lnkFiles.Length - _failedFiles.Count,lnkFiles.Length,sw.Elapsed.TotalSeconds);
            if (_failedFiles.Count > 0)
            {
                Console.WriteLine();
                Log.Information("Failed files");
                foreach (var failedFile in _failedFiles)
                {
                    Log.Information("  {FailedFile}",failedFile);
                }
            }

            var withProblems = _processedFiles.Count(t => t.ParseErrors.Count > 0);
            if (withProblems > 0)
            {
                Console.WriteLine();
                Log.Warning(
                    "{Count:N0} file(s) were parsed but have problems or anomalies. See the Anomalies column or the per-file output for details",
                    withProblems);
            }
        }

        var exportErrors = 0;

        if (_processedFiles.Count > 0)
        {
            Console.WriteLine();

            try
            {
                CsvWriter csvWriter = null;
                StreamWriter sw = null;

                if (csv?.Length > 0)
                {
                    if (Directory.Exists(csv) == false)
                    {
                        Log.Information(
                           "{Csv} does not exist. Creating...",csv);
                        Directory.CreateDirectory(csv);
                    }

                    var outName = $"{tsNow:yyyyMMddHHmmss}_LECmd_Output.csv";

                    if (csvf.IsNullOrEmpty() == false)
                    {
                        outName = Path.GetFileName(csvf);
                    }

                    var outFile = Path.Combine(csv, outName);

                    csv = Path.GetFullPath(outFile);
                    Log.Information(
                        "CSV output will be saved to {Path}",Path.GetFullPath(outFile));

                    try
                    {
                        sw = new StreamWriter(outFile);
                        csvWriter = new CsvWriter(sw, CultureInfo.InvariantCulture);

                        csvWriter.WriteHeader(typeof(CsvOut));
                        csvWriter.NextRecord();
                    }
                    catch (Exception ex)
                    {
                        exportErrors += 1;
                        Log.Error(ex,
                            "Unable to open {OutFile} for writing. CSV export canceled. Error: {Message}",outFile,ex.Message);
                    }
                }

                if (json?.Length > 0)
                {
                    if (Directory.Exists(json) == false)
                    {
                        Log.Information(
                            "{Json} does not exist. Creating...",json);
                        Directory.CreateDirectory(json);
                    }

                    Log.Information("Saving json output to {Json}",json);
                }

                if (xml?.Length > 0)
                {
                    {
                        if (Directory.Exists(xml) == false)
                        {
                            Log.Information("{Xml} does not exist. Creating...",xml);
                            Directory.CreateDirectory(xml);
                        }
                    }
                    Log.Information("Saving XML output to {Xml}",xml);
                }

                XmlTextWriter xmlTextWriter = null;
                StreamWriter jsonOut = null;

                if (json?.Length > 0)
                {

                    if (Directory.Exists(json) == false)
                    {
                        Directory.CreateDirectory(json);
                    }

                    var outName =
                        $"{tsNow:yyyyMMddHHmmss}_LECmd_Output.json";
                    var outFile = Path.Combine(json, outName);


                    jsonOut = new StreamWriter(new FileStream(outFile, FileMode.Create, FileAccess.Write),
                        new UTF8Encoding(false));
                }

                if (html?.Length > 0)
                {
                    var outDir = Path.Combine(html,
                        $"{tsNow:yyyyMMddHHmmss}_LECmd_Output_for_{html.Replace(@":\", "_").Replace(@"\", "_")}");

                    if (Directory.Exists(outDir) == false)
                    {
                        Directory.CreateDirectory(outDir);
                    }

                    File.WriteAllText(Path.Combine(outDir, "normalize.css"), Resources.normalize);
                    File.WriteAllText(Path.Combine(outDir, "style.css"), Resources.style);

                    var outFile = Path.Combine(outDir, "index.xhtml");

                    Log.Information("Saving HTML output to {OutFile}",outFile);

                    xmlTextWriter = new XmlTextWriter(outFile, Encoding.UTF8)
                    {
                        Formatting = Formatting.Indented,
                        Indentation = 4
                    };

                    xmlTextWriter.WriteStartDocument();

                    xmlTextWriter.WriteProcessingInstruction("xml-stylesheet", "href=\"normalize.css\"");
                    xmlTextWriter.WriteProcessingInstruction("xml-stylesheet", "href=\"style.css\"");

                    xmlTextWriter.WriteStartElement("document");
                }

                foreach (var processedFile in _processedFiles)
                {
                    CsvOut o;
                    try
                    {
                        o = GetCsvFormat(processedFile, false, dt);
                    }
                    catch (Exception ex)
                    {
                        exportErrors += 1;
                        Log.Error(ex,
                            "Error building output record for {SourceFile}. It is missing from the exported files. Error: {Message}",
                            processedFile.SourceFile, ex.Message);
                        continue;
                    }

                    try
                    {
                        csvWriter?.WriteRecord(o);
                        csvWriter?.NextRecord();
                    }
                    catch (Exception ex)
                    {
                        exportErrors += 1;
                        Log.Error(ex,
                            "Error writing record for {SourceFile} to {Csv}. Error: {Message}",processedFile.SourceFile,csv,ex.Message);
                    }

                    if (jsonOut != null)
                    {
                        try
                        {
                            var cs = GetCsvFormat(processedFile, true, JsonDateFormat);

                            jsonOut.WriteLine(pretty ? cs.Dump() : cs.ToJson());
                        }
                        catch (Exception ex)
                        {
                            exportErrors += 1;
                            Log.Error(ex,
                                "Error writing JSON for {SourceFile}. Error: {Message}", processedFile.SourceFile,
                                ex.Message);
                        }
                    }


                    //XHTML
                    xmlTextWriter?.WriteStartElement("Container");
                    xmlTextWriter?.WriteElementString("SourceFile", o.SourceFile);
                    xmlTextWriter?.WriteElementString("SourceCreated", o.SourceCreated);
                    xmlTextWriter?.WriteElementString("SourceModified", o.SourceModified);
                    xmlTextWriter?.WriteElementString("SourceAccessed", o.SourceAccessed);

                    xmlTextWriter?.WriteElementString("TargetCreated", o.TargetCreated);
                    xmlTextWriter?.WriteElementString("TargetModified", o.TargetModified);
                    xmlTextWriter?.WriteElementString("TargetAccessed", o.TargetAccessed);

                    xmlTextWriter?.WriteElementString("FileSize", o.FileSize.ToString());
                    xmlTextWriter?.WriteElementString("RelativePath", o.RelativePath);
                    xmlTextWriter?.WriteElementString("WorkingDirectory", o.WorkingDirectory);
                    xmlTextWriter?.WriteElementString("FileAttributes", o.FileAttributes);
                    xmlTextWriter?.WriteElementString("HeaderFlags", o.HeaderFlags);
                    xmlTextWriter?.WriteElementString("DriveType", o.DriveType);
                    xmlTextWriter?.WriteElementString("VolumeSerialNumber", o.VolumeSerialNumber);
                    xmlTextWriter?.WriteElementString("VolumeLabel", o.VolumeLabel);
                    xmlTextWriter?.WriteElementString("LocalPath", o.LocalPath);
                    xmlTextWriter?.WriteElementString("NetworkPath", o.NetworkPath);
                    xmlTextWriter?.WriteElementString("CommonPath", o.CommonPath);
                    xmlTextWriter?.WriteElementString("Arguments", o.Arguments);

                    xmlTextWriter?.WriteElementString("TargetIDAbsolutePath", o.TargetIDAbsolutePath);

                    xmlTextWriter?.WriteElementString("TargetMFTEntryNumber", $"{o.TargetMFTEntryNumber}");
                    xmlTextWriter?.WriteElementString("TargetMFTSequenceNumber",
                        $"{o.TargetMFTSequenceNumber}");

                    xmlTextWriter?.WriteElementString("MachineID", o.MachineID);
                    xmlTextWriter?.WriteElementString("MachineMACAddress", o.MachineMACAddress);
                    xmlTextWriter?.WriteElementString("MACVendor", o.MACVendor);

                    xmlTextWriter?.WriteElementString("TrackerCreatedOn", o.TrackerCreatedOn);

                    xmlTextWriter?.WriteElementString("ExtraBlocksPresent", o.ExtraBlocksPresent);

                    xmlTextWriter?.WriteElementString("Name", o.Name);
                    xmlTextWriter?.WriteElementString("IconLocation", o.IconLocation);
                    xmlTextWriter?.WriteElementString("ShowWindow", o.ShowWindow);
                    xmlTextWriter?.WriteElementString("EnvironmentVariables", o.EnvironmentVariables);
                    xmlTextWriter?.WriteElementString("VolumeDroid", o.VolumeDroid);
                    xmlTextWriter?.WriteElementString("FileDroid", o.FileDroid);
                    xmlTextWriter?.WriteElementString("VolumeDroidBirth", o.VolumeDroidBirth);
                    xmlTextWriter?.WriteElementString("FileDroidBirth", o.FileDroidBirth);
                    xmlTextWriter?.WriteElementString("Anomalies", o.Anomalies);

                    xmlTextWriter?.WriteEndElement();

                    if (xml?.Length > 0)
                    {
                        if (SaveXml(o, xml) == false)
                        {
                            exportErrors += 1;
                        }
                    }
                }


                //Close CSV stuff
                sw?.Flush();
                sw?.Close();

                //close json
                jsonOut?.Flush();
                jsonOut?.Close();

                //Close XML
                xmlTextWriter?.WriteEndElement();
                xmlTextWriter?.WriteEndDocument();
                xmlTextWriter?.Flush();
                xmlTextWriter?.Close();
            }
            catch (Exception ex)
            {
                exportErrors += 1;
                Log.Error(ex,"Error exporting data! Error: {Message}",ex.Message);
            }
        }

        //damaged files are listed under Failed files but are not a failure of the run
        return _readErrors > 0 || exportErrors > 0 ? 1 : 0;
    }


    private static CsvOut GetCsvFormat(LnkFile lnk, bool nukeNulls, string datetimeFormat)
    {
        string netPath = String.Empty;

        Log.Debug("Processing {f}",lnk.SourceFile);

        if (lnk.NetworkShareInfo != null)
        {
            netPath = lnk.NetworkShareInfo.NetworkShareName;
        }

        var csOut = new CsvOut
        {
            SourceFile = lnk.SourceFile.Replace("\\\\?\\", ""),
            SourceCreated = lnk.SourceCreated?.ToString(datetimeFormat) ?? string.Empty,
            SourceModified = lnk.SourceModified?.ToString(datetimeFormat) ?? string.Empty,
            SourceAccessed = lnk.SourceAccessed?.ToString(datetimeFormat) ?? string.Empty,
            TargetCreated = FormatTimestamp(lnk.Header.TargetCreationDate, datetimeFormat),
            TargetModified = FormatTimestamp(lnk.Header.TargetModificationDate, datetimeFormat),
            TargetAccessed = FormatTimestamp(lnk.Header.TargetLastAccessedDate, datetimeFormat),
            CommonPath = lnk.CommonPath,
            VolumeLabel = lnk.VolumeInfo?.VolumeLabel,
            VolumeSerialNumber = lnk.VolumeInfo?.VolumeSerialNumber,
            DriveType = lnk.VolumeInfo == null ? "(None)" : GetDescriptionFromEnumValue(lnk.VolumeInfo.DriveType),
            FileAttributes = lnk.Header.FileAttributes.ToString(),
            FileSize = lnk.Header.FileSize,
            HeaderFlags = lnk.Header.DataFlags.ToString(),
            LocalPath = lnk.LocalPath,
            NetworkPath = netPath,
            RelativePath = lnk.RelativePath,
            Arguments = lnk.Arguments,
            Name = lnk.Name,
            IconLocation = lnk.IconLocation,
            ShowWindow = lnk.Header.ShowWindow.ToString(),
            EnvironmentVariables = lnk.ExtraBlocks.OfType<EnvironmentVariableDataBlock>().FirstOrDefault() is { } env
                ? (env.EnvironmentVariablesUnicode.IsNullOrEmpty() ? env.EnvironmentVariablesAscii : env.EnvironmentVariablesUnicode)
                : null,
            Anomalies = string.Join(" | ", GetAnomalies(lnk))
        };


        if (lnk.TargetIDs?.Count > 0)
        {
            csOut.TargetIDAbsolutePath = GetAbsolutePathFromTargetIDs(lnk.TargetIDs);
        }

        csOut.WorkingDirectory = lnk.WorkingDirectory;

        var ebPresent = string.Empty;

        if (lnk.ExtraBlocks.Count > 0)
        {
            var names = new List<string>();

            foreach (var extraDataBase in lnk.ExtraBlocks)
            {
                names.Add(extraDataBase.GetType().Name);
            }

            ebPresent = string.Join(", ", names);
        }

        csOut.ExtraBlocksPresent = ebPresent;

        // FirstOrDefault, not SingleOrDefault: a (crafted) file with two tracker blocks used to throw here and abort
        // the export of every file
        var tnbBlock = lnk.ExtraBlocks.OfType<TrackerDataBaseBlock>().FirstOrDefault();

        if (tnbBlock != null)
        {
            csOut.MachineID = tnbBlock.MachineId;
            csOut.VolumeDroid = tnbBlock.VolumeDroid.ToString();
            csOut.FileDroid = tnbBlock.FileDroid.ToString();
            csOut.VolumeDroidBirth = tnbBlock.VolumeDroidBirth.ToString();
            csOut.FileDroidBirth = tnbBlock.FileDroidBirth.ToString();

            // The MAC address and creation time are only encoded in version 1 (time-based) UUIDs
            if (IsTimeBasedGuid(tnbBlock.FileDroid))
            {
                csOut.TrackerCreatedOn = tnbBlock.CreationTime.ToString(datetimeFormat);
                csOut.MachineMACAddress = tnbBlock.MacAddress;
                csOut.MACVendor = GetVendorFromMac(tnbBlock.MacAddress);
            }
        }

        if (lnk.TargetIDs?.Count > 0)
        {
            var si = lnk.TargetIDs.Last();

            var targetBlocks = TargetBlocks(si);
            if (targetBlocks.Count > 0)
            {
                var eb = targetBlocks.LastOrDefault(t => t is Beef0004);
                if (eb is Beef0004)
                {
                    var eb4 = eb as Beef0004;
                    if (eb4.MFTInformation.MFTEntryNumber != null)
                    {
                        csOut.TargetMFTEntryNumber = $"0x{eb4.MFTInformation.MFTEntryNumber.Value:X}";
                    }

                    if (eb4.MFTInformation.MFTSequenceNumber != null)
                    {
                        csOut.TargetMFTSequenceNumber =
                            $"0x{eb4.MFTInformation.MFTSequenceNumber.Value:X}";
                    }
                }
            }
        }

        if (nukeNulls)
        {
            foreach (var prop in csOut.GetType().GetProperties())
            {
                if (prop.GetValue(csOut) as string == string.Empty)
                {
                    prop.SetValue(csOut, null);
                }
            }
        }


        return csOut;
    }

    private static bool SaveXml(CsvOut csout, string outDir)
    {
        try
        {
            if (Directory.Exists(outDir) == false)
            {
                Directory.CreateDirectory(outDir);
            }

            // Include a short hash of the full source path: shortcuts with the same name in different folders
            // (Command Prompt.lnk, Desktop.lnk, ...) used to overwrite each other within the same second.
            string pathHash;
            using (var sha = SHA1.Create())
            {
                pathHash = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(csout.SourceFile ?? string.Empty)), 0, 4)
                    .Replace("-", "");
            }

            var outName =
                $"{DateTimeOffset.UtcNow:yyyyMMddHHmmss}_{Path.GetFileName(csout.SourceFile)}_{pathHash}.xml";
            var outFile = Path.Combine(outDir, outName);


            File.WriteAllText(outFile, csout.ToXml());

            return true;
        }
        catch (Exception ex)
        {
            Log.Error(ex,"Error exporting XML for {SourceFile}. Error: {Message}",csout.SourceFile,ex.Message);

            return false;
        }
    }

    private static string GetDescriptionFromEnumValue(Enum value)
    {
        var attribute = value.GetType()
            .GetField(value.ToString())
            ?.GetCustomAttributes(typeof(DescriptionAttribute), false)
            .SingleOrDefault() as DescriptionAttribute;
        // Undefined values (e.g. ShowCommand 0x0D in LECmd#12) have no field, so fall back to the raw value
        return attribute == null ? value.ToString() : attribute.Description;
    }

    private static string GetAbsolutePathFromTargetIDs(List<IShellBag> ids)
    {
        var absPath = string.Empty;

        if (ids == null)
        {
            return absPath;
        }

        foreach (var shellBag in ids)
        {
            if (shellBag is ShellBagDamaged)
            {
                // The placeholder's Value is an error message, not a path segment
                absPath += @"(damaged item)\";
                continue;
            }

            absPath += shellBag.Value?.Trim('\\') + @"\";
        }

        return absPath.Length > 0 ? absPath.Substring(0, absPath.Length - 1) : absPath;
    }

    /// <summary>
    ///     True when a file could not be opened or read (access denied, locked, an IO error), as opposed to a file that was
    ///     read but is damaged
    /// </summary>
    private static bool IsReadError(Exception ex)
    {
        return ex is UnauthorizedAccessException || (ex is IOException && !(ex is EndOfStreamException));
    }

#if !NET9_0_OR_GREATER
    /// <summary>
    ///     AlphaFS ErrorFilter: a folder that can't be read is reported and counted as a read error (exit 1 at the end), and
    ///     the search goes on (it was skipped silently).
    ///     The options leave out ContinueOnException: with it, AlphaFS never calls ErrorFilter
    /// </summary>
    private static bool UnreadableFolder(int errorCode, string errorMessage, string pathProcessed)
    {
        Log.Error("Unable to read folder {Folder}: {Reason}", pathProcessed, errorMessage);
        _readErrors++;
        return true;
    }
#else
    /// <summary>
    ///     Files under a folder, recursively, matching a mask (case-insensitive): files only, empty files skipped, and no
    ///     recursion into junctions/symlinks (they can loop on mounted images); reparse-point files (cloud placeholders) are
    ///     still included. Folders that can't be read are skipped and kept in Unreadable (IgnoreInaccessible skipped them
    ///     silently)
    /// </summary>
    private sealed class LnkFinder : FileSystemEnumerator<string>
    {
        private readonly string _mask;

        //the subfolder whose handle is opened right after ShouldRecurseIntoEntry; an error then is about that folder
        private string _lastRecursed;

        //folder (null: the folder searched, as the base constructor can report it before this one runs) and reason
        public readonly List<(string Folder, string Reason)> Unreadable = new();

        public LnkFinder(string directory, string mask) : base(directory,
            new EnumerationOptions { IgnoreInaccessible = false, RecurseSubdirectories = true, AttributesToSkip = 0 })
        {
            _mask = mask;
        }

        protected override bool ShouldRecurseIntoEntry(ref FileSystemEntry entry)
        {
            if ((entry.Attributes & FileAttributes.ReparsePoint) != 0)
            {
                return false;
            }

            _lastRecursed = entry.ToFullPath();
            return true;
        }

        protected override bool ShouldIncludeEntry(ref FileSystemEntry entry)
        {
            _lastRecursed = null;
            return !entry.IsDirectory && entry.Length > 0 && FileSystemName.MatchesSimpleExpression(_mask, entry.FileName);
        }

        protected override string TransformEntry(ref FileSystemEntry entry)
        {
            return entry.ToFullPath();
        }

        protected override bool ContinueOnError(int error)
        {
            Unreadable.Add((_lastRecursed, new Win32Exception(error).Message));
            _lastRecursed = null;
            return true;
        }
    }

    /// <summary>
    ///     The lnk files found; each folder that can't be read is reported and counted as a read error (exit 1 at the end)
    /// </summary>
    private static List<string> FindLnkFiles(string directory, string mask)
    {
        using var finder = new LnkFinder(directory, mask);
        var files = new List<string>();
        while (finder.MoveNext())
        {
            files.Add(finder.Current);
        }

        foreach (var (folder, reason) in finder.Unreadable)
        {
            Log.Error("Unable to read folder {Folder}: {Reason}", folder ?? directory, reason);
            _readErrors++;
        }

        return files;
    }
#endif

    private static LnkFile ProcessFile(string lnkFile, bool quiet, bool removableOnly, string datetimeFormat, bool nid,
        bool neb, int codepage=1252)
    {
        if (quiet == false)
        {
            Log.Information("Processing {Name}",lnkFile.Replace("\\\\?\\", ""));
            Console.WriteLine();
        }

        var sw = new Stopwatch();
        sw.Start();

        LnkFile lnk;

        try
        {
            lnk = Lnk.Lnk.LoadFile(lnkFile, codepage);
        }
        catch (Exception ex)
        {
            var where = ex is LnkParseException { Offset: >= 0 } pe ? $" [{pe.Field} at offset {pe.Offset}]" : string.Empty;

            _failedFiles.Add($"{lnkFile.Replace("\\\\?\\", "")} ==> ({ex.Message}){where}");
            Log.Error(ex,"Error opening {LnkFile}. Message: {Message}{Where}",lnkFile,ex.Message,where);
            Console.WriteLine();

            if (IsReadError(ex))
            {
                _readErrors++;
            }

            return null;
        }

        if (removableOnly && lnk.VolumeInfo?.DriveType != VolumeInfo.DriveTypes.DriveRemovable)
        {
            return null;
        }

        if (quiet == false)
        {
            // A problem while displaying details must not turn a successfully parsed file into a failed one
            // (previously -q and non -q runs could disagree about which files failed)
            try
            {
                DisplayLnk(lnk, nid, neb);
            }
            catch (Exception ex)
            {
                Log.Warning(ex,
                    "Unable to display all details for {LnkFile}. The file was parsed and is included in exported output. Error: {Message}",
                    lnkFile, ex.Message);
            }
        }
        else if (lnk.ParseErrors.Count > 0)
        {
            Log.Warning("{Name} was parsed with {Count:N0} problem(s): {Problems}", lnk.SourceFile.Replace("\\\\?\\", ""),
                lnk.ParseErrors.Count, string.Join(" | ", lnk.ParseErrors));
        }

        sw.Stop();

        if (quiet == false)
        {
            Console.WriteLine();
        }

        Log.Information("---------- Processed {Name} in {TotalSeconds:N8} seconds ----------",lnk.SourceFile.Replace("\\\\?\\", ""),sw.Elapsed.TotalSeconds);

        if (quiet == false)
        {
            Log.Information("\r\n");
        }

        return lnk;
    }

    private static void DisplayLnk(LnkFile lnk, bool nid, bool neb)
    {
        Log.Information("Source file: {Name}",lnk.SourceFile.Replace("\\\\?\\", ""));
        Log.Information("  Source created:  {Date}",lnk.SourceCreated);
        Log.Information("  Source modified: {Date}",lnk.SourceModified);
        Log.Information("  Source accessed: {Date}",lnk.SourceAccessed);
        Console.WriteLine();

        Log.Information("--- Header ---");

        var tc = ToNullableTimestamp(lnk.Header.TargetCreationDate);
        var tm = ToNullableTimestamp(lnk.Header.TargetModificationDate);
        var ta = ToNullableTimestamp(lnk.Header.TargetLastAccessedDate);

        Log.Information("  Target created:  {Tc}",tc);
        Log.Information("  Target modified: {Tm}",tm);
        Log.Information("  Target accessed: {Ta}",ta);
        Console.WriteLine();
        Log.Information("  File size (bytes): {FileSize:N0}",lnk.Header.FileSize);
        Log.Information("  Flags: {DataFlags}",lnk.Header.DataFlags);
        Log.Information("  File attributes: {FileAttributes}",lnk.Header.FileAttributes);

        if (lnk.Header.HotKey.Length > 0)
        {
            Log.Information("  Hot key: {HotKey}",lnk.Header.HotKey);
        }

        Log.Information("  Icon index: {IconIndex}",lnk.Header.IconIndex);
        Log.Information(
            "  Show window: {ShowWindow} ({Desc})",lnk.Header.ShowWindow,GetDescriptionFromEnumValue(lnk.Header.ShowWindow));

        Console.WriteLine();

        if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasName) == Lnk.Header.DataFlag.HasName)
        {
            Log.Information("Name: {Name}",lnk.Name);
        }

        if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasRelativePath) == Lnk.Header.DataFlag.HasRelativePath)
        {
            Log.Information("Relative Path: {RelativePath}",lnk.RelativePath);
        }

        if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasWorkingDir) == Lnk.Header.DataFlag.HasWorkingDir)
        {
            Log.Information("Working Directory: {WorkingDirectory}",lnk.WorkingDirectory);
        }

        if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasArguments) == Lnk.Header.DataFlag.HasArguments)
        {
            Log.Information("Arguments: {Arguments}",lnk.Arguments);
        }

        if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasIconLocation) == Lnk.Header.DataFlag.HasIconLocation)
        {
            Log.Information("Icon Location: {IconLocation}",lnk.IconLocation);
        }

        if ((lnk.Header.DataFlags & Lnk.Header.DataFlag.HasLinkInfo) == Lnk.Header.DataFlag.HasLinkInfo)
        {
            Console.WriteLine();
            Log.Information("--- Link information ---");
            Log.Information("Flags: {Flags}",lnk.LocationFlags);

            if (lnk.VolumeInfo != null)
            {
                Console.WriteLine();
                Log.Information(">> Volume information");
                Log.Information("  Drive type: {Drive}",GetDescriptionFromEnumValue(lnk.VolumeInfo.DriveType));
                Log.Information("  Serial number: {VolumeSerialNumber}",lnk.VolumeInfo.VolumeSerialNumber);

                var label = lnk.VolumeInfo.VolumeLabel?.Length > 0
                    ? lnk.VolumeInfo.VolumeLabel
                    : "(No label)";

                Log.Information("  Label: {Label}",label);
            }

            if (lnk.NetworkShareInfo != null)
            {
                Console.WriteLine();
                Log.Information("  Network share information");

                if (lnk.NetworkShareInfo.DeviceName?.Length > 0)
                {
                    Log.Information("    Device name: {DeviceName}",lnk.NetworkShareInfo.DeviceName);
                }

                Log.Information("    Share name: {NetworkShareName}",lnk.NetworkShareInfo.NetworkShareName);

                // NetworkProviderType is only meaningful when the ValidNetType flag is set
                Log.Information("    Provider type: {NetworkProviderType}",
                    lnk.NetworkShareInfo.HasValidProviderType
                        ? lnk.NetworkShareInfo.NetworkProviderType.ToString()
                        : "(not set)");
                Log.Information("    Share flags: {ShareFlags}",lnk.NetworkShareInfo.ShareFlags);
                Console.WriteLine();
            }

            if (lnk.LocalPath?.Length > 0)
            {
                Log.Information("  Local path: {LocalPath}",lnk.LocalPath);
            }

            if (lnk.CommonPath?.Length > 0)
            {
                Log.Information("  Common path: {CommonPath}",lnk.CommonPath);
            }
        }

        if (nid)
        {
            Console.WriteLine();
            Log.Information("(Target ID information suppressed. Lnk TargetID count: {Count:N0})",lnk.TargetIDs.Count);
        }

        if (lnk.TargetIDs.Count > 0 && !nid)
        {
            Console.WriteLine();

            Log.Information("--- Target ID information (Format: Type ==> Value) ---");
            Console.WriteLine();
            Log.Information("  Absolute path: {AbsPath}",GetAbsolutePathFromTargetIDs(lnk.TargetIDs));
            Console.WriteLine();

            foreach (var shellBag in lnk.TargetIDs)
            {
                //HACK
                //This is a total hack until i can refactor some shellbag code to clean things up

                var val = shellBag.Value.IsNullOrEmpty() ? "(None)" : shellBag.Value;

                Log.Information("  -{FriendlyName} ==> {Value}", shellBag.FriendlyName, val);

                //the shell items are ExtensionBlocks' classes (shared with ShellBags Explorer, SBECmd and the RECmd plugins)
                switch (shellBag)
                {
                    case ShellBagDamaged damagedItem:
                        Log.Warning("    This shell item could not be parsed: {Error}", damagedItem.ErrorMessage);
                        break;

                    //files and folders: 0x31 (and 0x35, 0x36, 0x39, 0xB1, which are read as one), 0x32, and 0x3A file entries
                    case ShellBag0X31 b:
                        LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks);
                        break;
                    case ShellBag0X32 b:
                        LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks);
                        break;
                    case ShellBag0X35 b:
                        LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks);
                        break;
                    case ShellBag0X36 b:
                        LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks);
                        break;
                    case ShellBag0X39 b:
                        LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks);
                        break;
                    case ShellBag0XB1 b:
                        LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks);
                        break;
                    case ShellBag0X3A b when b.ShortName != null:
                        LogFileEntry(b.ShortName, b.LastModificationTime, b.ExtensionBlocks);
                        break;

                    //users files folder (a 0x74 item with CFSF)
                    case ShellBagCFSF b:
                        if (b.LastModificationTime.HasValue)
                        {
                            Log.Information(
                                "    Modified:    {LastModificationTime}",b.LastModificationTime.Value);
                        }
                        else
                        {
                            Log.Information($"    Modified:");
                        }

                        LogExtensionBlocks(b.ExtensionBlocks);
                        break;

                    case ShellBag0X00 b:
                        foreach (var store in b.PropertyStores)
                        {
                            LogPropertyStore(store, false);
                        }

                        break;
                    case ShellBag0X1F b:
                        LogPropertyStore(b.PropertyStore, true);
                        break;
                    case ShellBag0X01 b:
                        if (b.DriveLetter?.Length > 0)
                        {
                            Log.Information("  Drive letter: {DriveLetter}",b.DriveLetter);
                        }

                        break;
                    case ShellBag0X71 b:
                        if (b.PropertyStore?.Sheets.Count > 0)
                        {
                            Log.Fatal("Property stores found! Please email lnk file to saericzimmerman@gmail.com so support can be added!!");
                        }

                        break;

                    //the value above is all there is
                    case ShellBag0X2E:
                    case ShellBag0X2F:
                    case ShellBag0X2A:
                    case ShellBag0X3A:
                    case ShellBag0X41:
                    case ShellBag0X42:
                    case ShellBag0X43:
                    case ShellBag0X46:
                    case ShellBag0X47:
                    case ShellBag0X61:
                    case ShellBag0X23:
                    case ShellBag0XC3:
                    case ShellBag0X74:
                    case ShellBag0X79:
                    case ShellBagZipContents:
                    case ShellBag0X4C: // SharePoint directory; the name and URL are in Value, already shown above
                        break;

                    default:
                        Log.Warning(">> UNMAPPED Type! Please email lnk file to saericzimmerman@gmail.com so support can be added!");
                        Log.Warning(">> {ShellBag}",shellBag);
                        break;
                }

                Console.WriteLine();
            }

            Log.Information("--- End Target ID information ---");
        }

        if (neb)
        {
            Console.WriteLine();
            Log.Information("(Extra blocks information suppressed. Lnk Extra block count: {Count:N0})",lnk.ExtraBlocks.Count);
        }

        if (lnk.ExtraBlocks.Count > 0 && !neb)
        {
            Console.WriteLine();
            Log.Information("--- Extra blocks information ---");
            
            Console.WriteLine();

            foreach (var extraDataBase in lnk.ExtraBlocks)
            {
                switch (extraDataBase.GetType().Name)
                {
                    case "ConsoleDataBlock":
                        var cdb = extraDataBase as ConsoleDataBlock;
                        Log.Information(">> Console data block");
                        Log.Information("   Fill Attributes: {FillAttributes}",cdb.FillAttributes);
                        Log.Information("   Popup Attributes: {PopupFillAttributes}",cdb.PopupFillAttributes);
                        Log.Information(
                            "   Buffer Size (Width x Height): {ScreenWidthBufferSize} x {ScreenHeightBufferSize}",cdb.ScreenWidthBufferSize,cdb.ScreenHeightBufferSize);
                        Log.Information(
                            "   Window Size (Width x Height): {WindowWidth} x {WindowHeight}",cdb.WindowWidth,cdb.WindowHeight);
                        Log.Information("   Origin (X/Y): {WindowOriginX}/{WindowOriginY}",cdb.WindowOriginX,cdb.WindowOriginY);
                        Log.Information("   Font Size: {FontSize}",cdb.FontSize);
                        Log.Information("   Is Bold: {IsBold}",cdb.IsBold);
                        Log.Information("   Face Name: {FaceName}",cdb.FaceName);
                        Log.Information("   Cursor Size: {CursorSize}",cdb.CursorSize);
                        Log.Information("   Is Full Screen: {IsFullScreen}",cdb.IsFullScreen);
                        Log.Information("   Is Quick Edit: {IsQuickEdit}",cdb.IsQuickEdit);
                        Log.Information("   Is Insert Mode: {IsInsertMode}",cdb.IsInsertMode);
                        Log.Information("   Is Auto Positioned: {IsAutoPositioned}",cdb.IsAutoPositioned);
                        Log.Information("   History Buffer Size: {HistoryBufferSize}",cdb.HistoryBufferSize);
                        Log.Information("   History Buffer Count: {HistoryBufferCount}",cdb.HistoryBufferCount);
                        Log.Information("   History Duplicates Allowed: {HistoryDuplicatesAllowed}",cdb.HistoryDuplicatesAllowed);
                        Console.WriteLine();
                        break;
                    case "ConsoleFEDataBlock":
                        var cfedb = extraDataBase as ConsoleFeDataBlock;
                        Log.Information(">> Console FE data block");
                        Log.Information("   Code page: {CodePage}",cfedb.CodePage);
                        Console.WriteLine();
                        break;
                    case "DarwinDataBlock":
                        var ddb = extraDataBase as DarwinDataBlock;
                        Log.Information(">> Darwin data block");
                        Log.Information("   Application ID: {ApplicationIdentifierUnicode}",ddb.ApplicationIdentifierUnicode);
                        Log.Information("   Product code: {ProductCode}",ddb.ProductCode);
                        Log.Information("   Feature name: {FeatureName}",ddb.FeatureName);
                        Log.Information("   Component ID: {ComponentId}",ddb.ComponentId);
                        Console.WriteLine();
                        break;
                    case "EnvironmentVariableDataBlock":
                        var evdb = extraDataBase as EnvironmentVariableDataBlock;
                        Log.Information(">> Environment variable data block");
                        Log.Information("   Environment variables: {EnvironmentVariablesUnicode}",evdb.EnvironmentVariablesUnicode);
                        Console.WriteLine();
                        break;
                    case "IconEnvironmentDataBlock":
                        var iedb = extraDataBase as IconEnvironmentDataBlock;
                        Log.Information(">> Icon environment data block");
                        Log.Information("   Icon path: {IconPathUni}",iedb.IconPathUni);
                        Console.WriteLine();
                        break;
                    case "KnownFolderDataBlock":
                        var kfdb = extraDataBase as KnownFolderDataBlock;
                        Log.Information(">> Known folder data block");
                        Log.Information(
                            "   Known folder GUID: {KnownFolderId} ==> {KnownFolderName}",kfdb.KnownFolderId,kfdb.KnownFolderName);
                        Console.WriteLine();
                        break;
                    case "PropertyStoreDataBlock":
                        var psdb = extraDataBase as PropertyStoreDataBlock;

                        if (psdb.PropertyStore.Sheets.Count > 0)
                        {
                            Log.Information(">> Property store data block (Format: GUID\\ID Description ==> Value)");
                            var propCount = 0;

                            foreach (var prop in psdb.PropertyStore.Sheets)
                            {
                                foreach (var propertyName in prop.PropertyNames)
                                {
                                    propCount += 1;

                                    var prefix = $"{prop.GUID}\\{propertyName.Key}".PadRight(43);
                                    var intParsed = int.TryParse(propertyName.Key, out var propNameAsInt);

                                    var suffix = string.Empty;

                                    if (intParsed)
                                    {
                                        suffix =
                                            $"{Utils.GetDescriptionFromGuidAndKey(prop.GUID, int.Parse(propertyName.Key))}"
                                                .PadRight(35);
                                    }

//                                                var suffix =
//                                                    $"{Utils.GetDescriptionFromGuidAndKey(prop.GUID, int.Parse(propertyName.Key))}"
//                                                        .PadRight(35);

                                    Log.Information("   {Prefix} {Suffix} ==> {PropertyNameValue}",prefix,suffix,propertyName.Value);
                                }
                            }

                            if (propCount == 0)
                            {
                                Log.Warning("   (Property store is empty)");
                            }
                        }

                        Console.WriteLine();
                        break;
                    case "ShimDataBlock":
                        var sdb = extraDataBase as ShimDataBlock;
                        Log.Information(">> Shimcache data block");
                        Log.Information("   Layer name: {LayerName}",sdb.LayerName);
                        Console.WriteLine();
                        break;
                    case "SpecialFolderDataBlock":
                        var sfdb = extraDataBase as SpecialFolderDataBlock;
                        Log.Information(">> Special folder data block");
                        Log.Information("   Special Folder ID: {SpecialFolderId}",sfdb.SpecialFolderId);
                        Console.WriteLine();
                        break;
                    case "TrackerDataBaseBlock":
                        var tdb = extraDataBase as TrackerDataBaseBlock;
                        Log.Information(">> Tracker database block");
                        Log.Information("   Machine ID:  {MachineId}",tdb.MachineId);
                        if (IsTimeBasedGuid(tdb.FileDroid))
                        {
                            Log.Information("   MAC Address: {MacAddress}",tdb.MacAddress);
                            Log.Information("   MAC Vendor:  {Address}",GetVendorFromMac(tdb.MacAddress));
                            Log.Information("   Creation:    {Date}",tdb.CreationTime);
                        }
                        else
                        {
                            Log.Information("   MAC Address / Creation: (File Droid is not a time-based UUID)");
                        }
                        Console.WriteLine();
                        Log.Information("   Volume Droid:       {VolumeDroid}",tdb.VolumeDroid);
                        Log.Information("   Volume Droid Birth: {VolumeDroidBirth}",tdb.VolumeDroidBirth);
                        Log.Information("   File Droid:         {FileDroid}",tdb.FileDroid);
                        Log.Information("   File Droid birth:   {FileDroidBirth}",tdb.FileDroidBirth);
                        Console.WriteLine();
                        break;
                    case "VistaAndAboveIdListDataBlock":
                        var vdb = extraDataBase as VistaAndAboveIdListDataBlock;
                        Log.Information(">> Vista and above ID List data block");

                        foreach (var shellBag in vdb.TargetIDs)
                        {
                            var val = shellBag.Value.IsNullOrEmpty() ? "(None)" : shellBag.Value;
                            Log.Information("   {FriendlyName} ==> {Val}", shellBag.FriendlyName, val);
                        }

                        Console.WriteLine();
                        break;
                    case "DamagedDataBlock":
                        var dmg = extraDataBase as DamagedDataBlock;
                        Log.Information("Damaged data block");
                        Log.Information("   Original Signature: {OriginalSignature}",dmg.OriginalSignature);
                        Log.Information("   Error Message: {ErrorMessage}",dmg.ErrorMessage);

                        break;
                }
            }
        }

        var anomalies = GetAnomalies(lnk).ToList();
        if (anomalies.Count > 0)
        {
            Console.WriteLine();
            Log.Warning("--- Parse problems and anomalies ---");
            foreach (var anomaly in anomalies)
            {
                Log.Warning("  {Anomaly}", anomaly);
            }
        }
    }

    /// <summary>
    ///     The target's extension blocks: a property view's (0x00) are those of the items in its Delegate ID List and then
    ///     its own, so the last 0xbeef0004 is the deepest item's (as when Lnk had its own shell item classes)
    /// </summary>
    private static List<IExtensionBlock> TargetBlocks(IShellBag item)
    {
        var own = item.ExtensionBlocks ?? new List<IExtensionBlock>();
        return item is ShellBag0X00 propertyView ? propertyView.DelegateExtensionBlocks.Concat(own).ToList() : own;
    }

    /// <summary>A file or folder item: its short name, modified time and extension blocks</summary>
    private static void LogFileEntry(string shortName, DateTimeOffset? modified, List<IExtensionBlock> extensionBlocks)
    {
        Log.Information("    Short name: {ShortName}",shortName);
        if (modified.HasValue)
        {
            Log.Information(
                "    Modified:    {LastModificationTime}",modified.Value);
        }
        else
        {
            Log.Information($"    Modified:");
        }

        LogExtensionBlocks(extensionBlocks);
    }

    private static void LogPropertyStore(PropertyStore store, bool header)
    {
        if (store == null || store.Sheets.Count == 0)
        {
            return;
        }

        Log.Information("  >> Property store (Format: GUID\\ID Description ==> Value)");
        var propCount = 0;

        foreach (var prop in store.Sheets)
        {
            foreach (var propertyName in prop.PropertyNames)
            {
                propCount += 1;

                var prefix = $"{prop.GUID}\\{propertyName.Key}".PadRight(43);
                var suffix = string.Empty;

                if (int.TryParse(propertyName.Key, out var propNameAsInt))
                {
                    suffix = $"{Utils.GetDescriptionFromGuidAndKey(prop.GUID, propNameAsInt)}".PadRight(35);
                }

                Log.Information("     {Prefix} {Suffix} ==> {PropertyName}",prefix,suffix,propertyName.Value);
            }
        }

        if (propCount == 0)
        {
            Log.Warning("     (Property store is empty)");
        }
    }
    private static void LogExtensionBlocks(List<IExtensionBlock> extensionBlocks)
    {
        if (extensionBlocks == null || extensionBlocks.Count == 0)
        {
            return;
        }

        Log.Information("    Extension block count: {Count:N0}",extensionBlocks.Count);
        Console.WriteLine();

        var extensionNumber = 0;
        foreach (var extensionBlock in extensionBlocks)
        {
            Log.Information("    --------- Block {ExtensionNumber:N0} ({Name}) ---------",extensionNumber,extensionBlock.GetType().Name);

            switch (extensionBlock)
            {
                case Beef0004 b4:
                    Log.Information("    Long name: {LongName}",b4.LongName);
                    if (b4.LocalisedName?.Length > 0)
                    {
                        Log.Information("    Localized name: {LocalisedName}",b4.LocalisedName);
                    }

                    if (b4.CreatedOnTime.HasValue)
                    {
                        Log.Information("    Created:     {CreatedOnTime}",b4.CreatedOnTime.Value);
                    }
                    else
                    {
                        Log.Information("    Created:");
                    }

                    if (b4.LastAccessTime.HasValue)
                    {
                        Log.Information("    Last access: {LastAccessTime}",b4.LastAccessTime.Value);
                    }
                    else
                    {
                        Log.Information("    Last access: ");
                    }

                    if (b4.MFTInformation.MFTEntryNumber > 0)
                    {
                        Log.Information(
                            "    MFT entry/sequence #: {MftEntryNumber}/{MftSequenceNumber} (0x{MftEntryNumber2:X}/0x{MftSequenceNumber2:X})",b4.MFTInformation.MFTEntryNumber,b4.MFTInformation.MFTSequenceNumber,b4.MFTInformation.MFTEntryNumber,b4.MFTInformation.MFTSequenceNumber);
                    }

                    break;
                case Beef0025 b25:
                    // FileTime1/FileTime2 are nullable; .Value used to throw and fail the whole file
                    Log.Information("    Filetime 1: {FileTime1}, Filetime 2: {FileTime2}",
                        b25.FileTime1.HasValue ? (object) b25.FileTime1.Value : string.Empty,
                        b25.FileTime2.HasValue ? (object) b25.FileTime2.Value : string.Empty);
                    break;
                case Beef0003 b3:
                    Log.Information("    GUID: {Guid1} ({Guid1Folder})",b3.GUID1,b3.GUID1Folder);
                    break;
                case Beef001a b1a:
                    Log.Information("    File document type: {FileDocumentTypeString}",b1a.FileDocumentTypeString);
                    break;
                default:
                    Log.Information("    {ExtensionBlock}",extensionBlock);
                    break;
            }

            extensionNumber += 1;
        }
    }

    /// <summary>
    ///     1601-01-01 means "not set"; DateTimeOffset.MinValue is what Lnk returns for an out-of-range FILETIME
    ///     (the raw value is reported in the anomalies).
    /// </summary>
    private static bool IsUnsetTimestamp(DateTimeOffset value) => value.Year <= 1601;

    private static string FormatTimestamp(DateTimeOffset value, string format) =>
        IsUnsetTimestamp(value) ? string.Empty : value.ToString(format);

    private static DateTimeOffset? ToNullableTimestamp(DateTimeOffset value) =>
        IsUnsetTimestamp(value) ? null : value;

    /// <summary>
    ///     Only version 1 (time-based) UUIDs carry a timestamp and a MAC address.
    /// </summary>
    private static bool IsTimeBasedGuid(Guid guid) => (guid.ToByteArray()[7] >> 4) == 1;

    private static IEnumerable<string> GetAnomalies(LnkFile lnk)
    {
        // Parse problems, data appended after the terminal block, out-of-range timestamps, damaged items/blocks
        foreach (var error in lnk.ParseErrors)
        {
            yield return error;
        }

        var args = lnk.Arguments;
        if (string.IsNullOrEmpty(args))
        {
            yield break;
        }

        // Padding Arguments past what the shortcut properties dialog shows is a known way of hiding a payload
        if (args.Length > 260)
        {
            yield return $"Arguments are {args.Length:N0} characters long (the shortcut properties dialog shows 260)";
        }

        var longestRun = 0;
        var run = 0;
        foreach (var c in args)
        {
            run = char.IsWhiteSpace(c) ? run + 1 : 0;
            longestRun = Math.Max(longestRun, run);
        }

        if (longestRun >= 32)
        {
            yield return $"Arguments contain {longestRun:N0} consecutive whitespace characters";
        }
    }

    private class CustomHelpAction : SynchronousCommandLineAction
    {
        private readonly HelpAction _defaultHelp;

        public CustomHelpAction(HelpAction action)
        {
            _defaultHelp = action;
        }

        public override int Invoke(ParseResult parseResult)
        {
            var result = _defaultHelp.Invoke(parseResult);

            Log.Warning("{Msg}", string.Join(" ",parseResult.Tokens));

            return result;
        }
    }

    private static string GetVendorFromMac(string macAddress)
    {
        //00-00-00	XEROX CORPORATION
        //"00:14:22:0d:94:04"

        var mac = string.Join("", macAddress.Split(':').Take(3)).ToUpperInvariant();
        // .Replace(":", "-").ToUpper();

        var vendor = "(Unknown vendor)";

        if (MacList.TryGetValue(mac, out var value))
        {
            vendor = value;
        }

        return vendor;
    }

    private static readonly string BaseDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
}

public sealed class CsvOut
{
    public string SourceFile { get; set; }
    public string SourceCreated { get; set; }
    public string SourceModified { get; set; }
    public string SourceAccessed { get; set; }
    public string TargetCreated { get; set; }
    public string TargetModified { get; set; }
    public string TargetAccessed { get; set; }
    public uint FileSize { get; set; }
    public string RelativePath { get; set; }
    public string WorkingDirectory { get; set; }
    public string FileAttributes { get; set; }
    public string HeaderFlags { get; set; }
    public string DriveType { get; set; }
    public string VolumeSerialNumber { get; set; }
    public string VolumeLabel { get; set; }
    public string LocalPath { get; set; }
    public string NetworkPath { get; set; }
    public string CommonPath { get; set; }
    public string Arguments { get; set; }
    public string TargetIDAbsolutePath { get; set; }
    public string TargetMFTEntryNumber { get; set; }
    public string TargetMFTSequenceNumber { get; set; }
    public string MachineID { get; set; }
    public string MachineMACAddress { get; set; }
    public string MACVendor { get; set; }
    public string TrackerCreatedOn { get; set; }
    public string ExtraBlocksPresent { get; set; }

    // Added after the original columns so existing Timeline Explorer layouts and scripts keep working
    public string Name { get; set; }
    public string IconLocation { get; set; }
    public string ShowWindow { get; set; }
    public string EnvironmentVariables { get; set; }
    public string VolumeDroid { get; set; }
    public string FileDroid { get; set; }
    public string VolumeDroidBirth { get; set; }
    public string FileDroidBirth { get; set; }
    public string Anomalies { get; set; }
}

internal class ApplicationArguments
{
    public string File { get; set; }
    public string Directory { get; set; }

    public string JsonDirectory { get; set; }
    public bool JsonPretty { get; set; }
    public bool AllFiles { get; set; }
    public bool NoTargetIDList { get; set; }
    public bool NoExtraBlocks { get; set; }
    public string CsvDirectory { get; set; }
    public string CsvName { get; set; }
    public string XmlDirectory { get; set; }
    public string xHtmlDirectory { get; set; }

    public string DateTimeFormat { get; set; }

    public bool PreciseTimestamps { get; set; }

    public bool RemovableOnly { get; set; }

    public bool Quiet { get; set; }
}
