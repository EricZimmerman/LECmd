using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Help;
using System.CommandLine.NamingConventionBinder;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
#if NET462
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
using Spectre.Console;
using CsvWriter = CsvHelper.CsvWriter;
using Resources = LECmd.Properties.Resources;
using ShellBag = Lnk.ShellItems.ShellBag;
using VolumeInfo = Lnk.VolumeInfo;

namespace LECmd;

internal class Program
{
    private static readonly string PreciseTimeFormat = "yyyy-MM-dd HH:mm:ss.fffffff";

    private static string _activeDateTimeFormat;

    private static RootCommand _rootCommand;

    private static readonly string Header = $"LECmd version {Assembly.GetExecutingAssembly().GetName().Version}" +
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

    private static async Task Main(string[] args)
    {
        ExceptionlessClient.Default.Startup("FNWfyFuaUAPnVfofTZAhZOgeDG5lv7AnjYKNtsEJ");

        LoadMacs();

        _rootCommand = new RootCommand
        {
            new Option<string>(
                "-f",
                description: "File to process. Either this or -d is required"),
            new Option<string>(
                "-d",
                description: "Directory to recursively process. Either this or -f is required"),
            new Option<bool>(
                "-r",
                getDefaultValue: () => false,
                "Only process lnk files pointing to removable drives"),

            new Option<bool>(
                "-q",
                getDefaultValue: () => false,
                "Only show the filename being processed vs all output. Useful to speed up exporting to json and/or csv"),

            new Option<bool>(
                "--all",
                getDefaultValue: () => false,
                "Process all files in directory vs. only files matching *.lnk"),

            new Option<string>(
                "--csv",
                "Directory to save CSV formatted results to. Be sure to include the full path in double quotes"),

            new Option<string>(
                "--csvf",
                "File name to save CSV formatted results to. When present, overrides default name\r\n"),

            new Option<string>(
                "--xml",
                "Directory to save XML formatted results to. Be sure to include the full path in double quotes"),

            new Option<string>(
                "--html",
                "Directory to save xhtml formatted results to. Be sure to include the full path in double quotes"),

            new Option<string>(
                "--json",
                "Directory to save json representation to. Use --pretty for a more human readable layout"),

            new Option<bool>(
                "--pretty",
                getDefaultValue: () => false,
                "When exporting to json, use a more human readable layout"),

            new Option<bool>(
                "--nid",
                getDefaultValue: () => false,
                "Suppress Target ID list details from being displayed"),

            new Option<bool>(
                "--neb",
                getDefaultValue: () => false,
                "Suppress Extra blocks information from being displayed"),

            new Option<string>(
                "--dt",
                getDefaultValue: () => "yyyy-MM-dd HH:mm:ss",
                "The custom date/time format to use when displaying time stamps. See https://goo.gl/CNVq0k for options"),

            new Option<bool>(
                "--mp",
                getDefaultValue: () => false,
                "Display higher precision for time stamps"),

            new Option<bool>(
                "--debug",
                () => false,
                "Show debug information during processing"),

            new Option<bool>(
                "--trace",
                () => false,
                "Show trace information during processing")
        };

        _rootCommand.Description = Header + "\r\n\r\n" + Footer;

        _rootCommand.Handler = CommandHandler.Create(DoWork);

        await _rootCommand.InvokeAsync(args);

        Log.CloseAndFlush();
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

    private static void DoWork(string f, string d, bool r, bool q, bool all, string csv, string csvf, string xml,
        string html, string json, bool pretty, bool nid, bool neb, string dt, bool mp, bool debug, bool trace)
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
            var helpBld = new HelpBuilder(LocalizationResources.Instance, Console.WindowWidth);

            var hc = new HelpContext(helpBld, _rootCommand, Console.Out);

            helpBld.Write(hc);

            Log.Warning("Either -f or -d is required. Exiting");
            Console.WriteLine();
            return;
        }

        if (f.IsNullOrEmpty() == false && !File.Exists(f))
        {
            var helpBld = new HelpBuilder(LocalizationResources.Instance, Console.WindowWidth);
            var hc = new HelpContext(helpBld, _rootCommand, Console.Out);

            helpBld.Write(hc);

            Log.Warning("File {F} not found. Exiting",f);
            Console.WriteLine();
            return;
        }

        if (d.IsNullOrEmpty() == false &&
            !Directory.Exists(d))
        {
            var helpBld = new HelpBuilder(LocalizationResources.Instance, Console.WindowWidth);
            var hc = new HelpContext(helpBld, _rootCommand, Console.Out);

            helpBld.Write(hc);

            Log.Warning("Directory {D} not found. Exiting",d);
            Console.WriteLine();
            return;
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

        var tsNow = DateTimeOffset.UtcNow;

        if (f?.Length > 0)
        {
            LnkFile lnk;

            try
            {
                f = Path.GetFullPath(f);

                lnk = ProcessFile(f, q, r, dt, nid, neb);
                if (lnk != null)
                {
                    _processedFiles.Add(lnk);
                }
            }
            catch (UnauthorizedAccessException ua)
            {
                Log.Error(ua,
                    "Unable to access {F}. Are you running as an administrator? Error: {Message}",f,ua.Message);
                return;
            }
            catch (Exception ex)
            {
                Log.Error(ex,
                    "Error processing file {F} Please send it to saericzimmerman@gmail.com. Error: {Message}",f,ex.Message);
                return;
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

#if NET6_0
                        var enumerationOptions = new EnumerationOptions
                        {
                            IgnoreInaccessible = true,
                            MatchCasing = MatchCasing.CaseInsensitive,
                            RecurseSubdirectories = true,
                            AttributesToSkip = 0
                        };
                        
                       files2 =
                            Directory.EnumerateFileSystemEntries(d, mask,enumerationOptions);

#elif NET462
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
                        
                        directoryEnumerationFilters.ErrorFilter = (errorCode, errorMessage, pathProcessed) => true;
                        
                        var dirEnumOptions =
                            DirectoryEnumerationOptions.Files | DirectoryEnumerationOptions.Recursive |
                            DirectoryEnumerationOptions.SkipReparsePoints | DirectoryEnumerationOptions.ContinueOnException |
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
                return;
            }
            catch (Exception ex)
            {
                Log.Error(ex,
                    "Error getting lnk files in {D}. Error: {Message}",d,ex.Message);
                return;
            }

            Log.Information("Found {Length:N0} files",lnkFiles.Length);
            Console.WriteLine();

            var sw = new Stopwatch();
            sw.Start();

            foreach (var file in lnkFiles)
            {
                var lnk = ProcessFile(file, q, r, dt, nid, neb);
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
        }

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
                    JsConfig.DateHandler = DateHandler.ISO8601;

                    if (Directory.Exists(json) == false)
                    {
                        Directory.CreateDirectory(json);
                    }

                    var outName =
                        $"{tsNow:yyyyMMddHHmmss}_LECmd_Output.json";
                    var outFile = Path.Combine(json, outName);


                    jsonOut = new StreamWriter(new FileStream(outFile, FileMode.OpenOrCreate, FileAccess.Write),
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

                    var outFile = Path.Combine(html, outDir, "index.xhtml");

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
                    var o = GetCsvFormat(processedFile, false, dt);

                    try
                    {
                        csvWriter?.WriteRecord(o);
                        csvWriter?.NextRecord();
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex,
                            "Error writing record for {SourceFile} to {Csv}. Error: {Message}",processedFile.SourceFile,csv,ex.Message);
                    }

                    if (jsonOut != null)
                    {
                        var oldDt = dt;

                        dt = "o";

                        var cs = GetCsvFormat(processedFile, true, dt);

                        dt = oldDt;

                        if (pretty)
                        {
                            jsonOut.WriteLine(cs.Dump());
                        }
                        else
                        {
                            jsonOut.WriteLine(cs.ToJson());
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

                    xmlTextWriter?.WriteEndElement();

                    if (xml?.Length > 0)
                    {
                        SaveXml(o, xml);
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
            }
            catch (Exception ex)
            {
                Log.Error(ex,"Error exporting data! Error: {Message}",ex.Message);
            }
        }
    }


    private static CsvOut GetCsvFormat(LnkFile lnk, bool nukeNulls, string datetimeFormat)
    {
        string netPath = String.Empty;

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
            TargetCreated = lnk.Header.TargetCreationDate.Year == 1601
                ? string.Empty
                : lnk.Header.TargetCreationDate.ToString(datetimeFormat),
            TargetModified = lnk.Header.TargetModificationDate.Year == 1601
                ? string.Empty
                : lnk.Header.TargetModificationDate.ToString(datetimeFormat),
            TargetAccessed = lnk.Header.TargetLastAccessedDate.Year == 1601
                ? String.Empty
                : lnk.Header.TargetLastAccessedDate.ToString(datetimeFormat),
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
            Arguments = lnk.Arguments
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

        var tnb = lnk.ExtraBlocks.SingleOrDefault(t => t.GetType().Name.ToUpper() == "TRACKERDATABASEBLOCK");


        if (tnb != null)
        {
            var tnbBlock = tnb as TrackerDataBaseBlock;

            csOut.TrackerCreatedOn = tnbBlock?.CreationTime.ToString(datetimeFormat);

            csOut.MachineID = tnbBlock?.MachineId;
            csOut.MachineMACAddress = tnbBlock?.MacAddress;
            csOut.MACVendor = GetVendorFromMac(tnbBlock?.MacAddress);
        }

        if (lnk.TargetIDs?.Count > 0)
        {
            var si = lnk.TargetIDs.Last();

            if (si.ExtensionBlocks?.Count > 0)
            {
                var eb = si.ExtensionBlocks.LastOrDefault(t => t is Beef0004);
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

    private static void SaveXml(CsvOut csout, string outDir)
    {
        try
        {
            if (Directory.Exists(outDir) == false)
            {
                Directory.CreateDirectory(outDir);
            }

            var outName =
                $"{DateTimeOffset.UtcNow:yyyyMMddHHmmss}_{Path.GetFileName(csout.SourceFile)}.xml";
            var outFile = Path.Combine(outDir, outName);


            File.WriteAllText(outFile, csout.ToXml());
        }
        catch (Exception ex)
        {
            Log.Error(ex,"Error exporting XML for {SourceFile}. Error: {Message}",csout.SourceFile,ex.Message);
        }
    }

    private static string GetDescriptionFromEnumValue(Enum value)
    {
        var attribute = value.GetType()
            .GetField(value.ToString())
            .GetCustomAttributes(typeof(DescriptionAttribute), false)
            .SingleOrDefault() as DescriptionAttribute;
        return attribute?.Description;
    }

    private static string GetAbsolutePathFromTargetIDs(List<ShellBag> ids)
    {
        var absPath = string.Empty;

        foreach (var shellBag in ids)
        {
            absPath += shellBag.Value + @"\";
        }

        absPath = absPath.Substring(0, absPath.Length - 1);

        return absPath;
    }

    private static LnkFile ProcessFile(string lnkFile, bool quiet, bool removableOnly, string datetimeFormat, bool nid,
        bool neb)
    {
        if (quiet == false)
        {
            Log.Information("Processing {Name}",lnkFile.Replace("\\\\?\\", ""));
            Console.WriteLine();
        }

        var sw = new Stopwatch();
        sw.Start();

        try
        {
            var lnk = Lnk.Lnk.LoadFile(lnkFile);

            if (removableOnly && lnk.VolumeInfo?.DriveType != VolumeInfo.DriveTypes.DriveRemovable)
            {
                return null;
            }

            if (quiet == false)
            {
                Log.Information("Source file: {Name}",lnk.SourceFile.Replace("\\\\?\\", ""));
                Log.Information("  Source created:  {Date}",lnk.SourceCreated);
                Log.Information("  Source modified: {Date}",lnk.SourceModified);
                Log.Information("  Source accessed: {Date}",lnk.SourceAccessed);
                Console.WriteLine();

                Log.Information("--- Header ---");

                DateTimeOffset? tc = lnk.Header.TargetCreationDate.Year == 1601
                    ? null
                    : lnk.Header.TargetCreationDate;
                DateTimeOffset? tm = lnk.Header.TargetModificationDate.Year == 1601
                    ? null
                    : lnk.Header.TargetModificationDate;
                DateTimeOffset? ta = lnk.Header.TargetLastAccessedDate.Year == 1601
                    ? null
                    : lnk.Header.TargetLastAccessedDate;

                Log.Information("  Target created:  {Tc}",tc);
                Log.Information("  Target modified: {Tm}",tm);
                Log.Information("  Target accessed: {Ta}",ta);
                Console.WriteLine();
                Log.Information("  File size: {FileSize:N0}",lnk.Header.FileSize);
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

                        var label = lnk.VolumeInfo.VolumeLabel.Length > 0
                            ? lnk.VolumeInfo.VolumeLabel
                            : "(No label)";

                        Log.Information("  Label: {Label}",label);
                    }

                    if (lnk.NetworkShareInfo != null)
                    {
                        Console.WriteLine();
                        Log.Information("  Network share information");

                        if (lnk.NetworkShareInfo.DeviceName.Length > 0)
                        {
                            Log.Information("    Device name: {DeviceName}",lnk.NetworkShareInfo.DeviceName);
                        }

                        Log.Information("    Share name: {NetworkShareName}",lnk.NetworkShareInfo.NetworkShareName);

                        Log.Information("    Provider type: {NetworkProviderType}",lnk.NetworkShareInfo.NetworkProviderType);
                        Log.Information("    Share flags: {ShareFlags}",lnk.NetworkShareInfo.ShareFlags);
                        Console.WriteLine();
                    }

                    if (lnk.LocalPath?.Length > 0)
                    {
                        Log.Information("  Local path: {LocalPath}",lnk.LocalPath);
                    }

                    if (lnk.CommonPath.Length > 0)
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

                        Log.Information("  -{FriendlyName} ==> {val}", shellBag.FriendlyName, val);

                        switch (shellBag.GetType().Name.ToUpper())
                        {
                            case "SHELLBAG0X36":
                            case "SHELLBAG0X32":
                                var b32 = shellBag as ShellBag0X32;

                                Log.Information("    Short name: {ShortName}",b32.ShortName);
                                if (b32.LastModificationTime.HasValue)
                                {
                                    Log.Information(
                                        "    Modified:    {LastModificationTime}",b32.LastModificationTime.Value);
                                }
                                else
                                {
                                    Log.Information($"    Modified:");
                                }


                                var extensionNumber32 = 0;
                                if (b32.ExtensionBlocks.Count > 0)
                                {
                                    Log.Information("    Extension block count: {Count:N0}",b32.ExtensionBlocks.Count);
                                    Console.WriteLine();
                                    foreach (var extensionBlock in b32.ExtensionBlocks)
                                    {
                                        Log.Information("    --------- Block {ExtensionNumber32:N0} ({Name}) ---------",extensionNumber32,extensionBlock.GetType().Name);
                                        
                                        if (extensionBlock is Beef0004)
                                        {
                                            var b4 = extensionBlock as Beef0004;

                                            Log.Information("    Long name: {LongName}",b4.LongName);
                                            if (b4.LocalisedName.Length > 0)
                                            {
                                                Log.Information("    Localized name: {LocalisedName}",b4.LocalisedName);
                                            }

                                            if (b4.CreatedOnTime.HasValue)
                                            {
                                                Log.Information(
                                                    "    Created:     {CreatedOnTime}",b4.CreatedOnTime.Value);
                                            }
                                            else
                                            {
                                                Log.Information($"    Created:");
                                            }

                                            if (b4.LastAccessTime.HasValue)
                                            {
                                                Log.Information(
                                                    "    Last access: {LastAccessTime}",b4.LastAccessTime.Value);
                                            }
                                            else
                                            {
                                                Log.Information($"    Last access: ");
                                            }

                                            if (b4.MFTInformation.MFTEntryNumber > 0)
                                            {
                                                Log.Information(
                                                    "    MFT entry/sequence #: {MftEntryNumber}/{MftSequenceNumber} (0x{MftEntryNumber2:X}/0x{MftSequenceNumber2:X})",b4.MFTInformation.MFTEntryNumber,b4.MFTInformation.MFTSequenceNumber,b4.MFTInformation.MFTEntryNumber,b4.MFTInformation.MFTSequenceNumber);
                                            }
                                        }
                                        else if (extensionBlock is Beef0025)
                                        {
                                            var b25 = extensionBlock as Beef0025;
                                            Log.Information(
                                                "    Filetime 1: {FileTime1}, Filetime 2: {FileTime2}",b25.FileTime1.Value,b25.FileTime2.Value);
                                        }
                                        else if (extensionBlock is Beef0003)
                                        {
                                            var b3 = extensionBlock as Beef0003;
                                            Log.Information("    GUID: {Guid1} ({Guid1Folder})",b3.GUID1,b3.GUID1Folder);
                                        }
                                        else if (extensionBlock is Beef001a)
                                        {
                                            var b3 = extensionBlock as Beef001a;
                                            Log.Information("    File document type: {FileDocumentTypeString}",b3.FileDocumentTypeString);
                                        }

                                        else
                                        {
                                            Log.Information("    {ExtensionBlock}",extensionBlock);
                                        }

                                        extensionNumber32 += 1;
                                    }
                                }

                                break;
                            case "SHELLBAG0X31":

                                var b3x = shellBag as Lnk.ShellItems.ShellBag0X31;

                                Log.Information("    Short name: {ShortName}",b3x.ShortName);
                                if (b3x.LastModificationTime.HasValue)
                                {
                                    Log.Information(
                                        "    Modified:    {LastModificationTime}",b3x.LastModificationTime.Value);
                                }
                                else
                                {
                                    Log.Information($"    Modified:");
                                }

                                var extensionNumber = 0;
                                if (b3x.ExtensionBlocks.Count > 0)
                                {
                                    Log.Information("    Extension block count: {Count:N0}",b3x.ExtensionBlocks.Count);
                                    Console.WriteLine();
                                    foreach (var extensionBlock in b3x.ExtensionBlocks)
                                    {
                                        Log.Information("    --------- Block {ExtensionNumber:N0} ({Name}) ---------",extensionNumber,extensionBlock.GetType().Name);
                                    
                                        if (extensionBlock is Beef0004)
                                        {
                                            var b4 = extensionBlock as Beef0004;

                                            Log.Information("    Long name: {LongName}",b4.LongName);
                                            if (b4.LocalisedName.Length > 0)
                                            {
                                                Log.Information("    Localized name: {LocalisedName}",b4.LocalisedName);
                                            }

                                            if (b4.CreatedOnTime.HasValue)
                                            {
                                                Log.Information(
                                                    "    Created:     {CreatedOnTime}",b4.CreatedOnTime.Value);
                                            }
                                            else
                                            {
                                                Log.Information($"    Created:");
                                            }

                                            if (b4.LastAccessTime.HasValue)
                                            {
                                                Log.Information(
                                                    "    Last access: {LastAccessTime}",b4.LastAccessTime);
                                            }
                                            else
                                            {
                                                Log.Information($"    Last access: ");
                                            }

                                            if (b4.MFTInformation.MFTEntryNumber > 0)
                                            {
                                                Log.Information(
                                                    "    MFT entry/sequence #: {MftEntryNumber}/{MftSequenceNumber} (0x{MftEntryNumber2:X}/0x{MftSequenceNumber2:X})",b4.MFTInformation.MFTEntryNumber,b4.MFTInformation.MFTSequenceNumber,b4.MFTInformation.MFTEntryNumber,b4.MFTInformation.MFTSequenceNumber);
                                            }
                                        }
                                        else if (extensionBlock is Beef0025)
                                        {
                                            var b25 = extensionBlock as Beef0025;
                                            Log.Information(
                                                "    Filetime 1: {FileTime1}, Filetime 2: {FileTime2}",b25.FileTime1.Value,b25.FileTime2.Value);
                                        }
                                        else if (extensionBlock is Beef0003)
                                        {
                                            var b3 = extensionBlock as Beef0003;
                                            Log.Information("    GUID: {Guid1} ({Guid1Folder})",b3.GUID1,b3.GUID1Folder);
                                        }
                                        else if (extensionBlock is Beef001a)
                                        {
                                            var b3 = extensionBlock as Beef001a;
                                            Log.Information("    File document type: {FileDocumentTypeString}",b3.FileDocumentTypeString);
                                        }
                                        else
                                        {
                                            Log.Information("    {ExtensionBlock}",extensionBlock);
                                        }

                                        extensionNumber += 1;
                                    }
                                }

                                break;

                            case "SHELLBAG0X00":
                                var b00 = shellBag as ShellBag0X00;

                                if (b00.PropertyStore.Sheets.Count > 0)
                                {
                                    Log.Information("  >> Property store (Format: GUID\\ID Description ==> Value)");
                                    var propCount = 0;

                                    foreach (var prop in b00.PropertyStore.Sheets)
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
                                                    $"{Utils.GetDescriptionFromGuidAndKey(prop.GUID, propNameAsInt)}"
                                                        .PadRight(35);
                                            }

//                                                var suffix =
//                                                    $"{Utils.GetDescriptionFromGuidAndKey(prop.GUID, int.Parse(propertyName.Key))}"
//                                                        .PadRight(35);

                                            Log.Information("     {Prefix} {Suffix} ==> {PropertyName}",prefix,suffix,propertyName.Value);
                                        }
                                    }

                                    if (propCount == 0)
                                    {
                                        Log.Warning("     (Property store is empty)");
                                    }
                                }

                                break;
                            case "SHELLBAG0X01":
                                var baaaa1f = shellBag as ShellBag0X01;
                                if (baaaa1f.DriveLetter.Length > 0)
                                {
                                    Log.Information("  Drive letter: {DriveLetter}",baaaa1f.DriveLetter);
                                }

                                break;
                            case "SHELLBAG0X1F":

                                var b1f = shellBag as ShellBag0X1F;

                                if (b1f.PropertyStore.Sheets.Count > 0)
                                {
                                    Log.Information("  >> Property store (Format: GUID\\ID Description ==> Value)");
                                    var propCount = 0;

                                    foreach (var prop in b1f.PropertyStore.Sheets)
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

                                            Log.Information("     {Prefix} {Suffix} ==> {PropertyName}",prefix,suffix,propertyName.Value);
                                        }
                                    }

                                    if (propCount == 0)
                                    {
                                        Log.Warning("     (Property store is empty)");
                                    }
                                }

                                break;
                            case "SHELLBAG0X2E":
                                break;
                            case "SHELLBAG0X2F":
                                var b2f = shellBag as ShellBag0X2F;

                                break;
                            case "SHELLBAG0X40":
                                break;
                            case "SHELLBAG0X61":

                                break;
                            case "SHELLBAG0X71":
                                var b71 = shellBag as ShellBag0X71;
                                if (b71.PropertyStore?.Sheets.Count > 0)
                                {
                                    Log.Fatal("Property stores found! Please email lnk file to saericzimmerman@gmail.com so support can be added!!");
                                }

                                break;
                            case "SHELLBAG0X74":
                                var b74 = shellBag as ShellBag0X74;

                                if (b74.LastModificationTime.HasValue)
                                {
                                    Log.Information(
                                        "    Modified:    {LastModificationTime}",b74.LastModificationTime.Value);
                                }
                                else
                                {
                                    Log.Information($"    Modified:");
                                }

                                var extensionNumber74 = 0;
                                if (b74.ExtensionBlocks.Count > 0)
                                {
                                    Log.Information("    Extension block count: {Count:N0}",b74.ExtensionBlocks.Count);
                                    Console.WriteLine();
                                    foreach (var extensionBlock in b74.ExtensionBlocks)
                                    {
                                        Log.Information("    --------- Block {ExtNum:N0} ({Name}) ---------",extensionNumber74,extensionBlock.GetType().Name);
                                        
                                        if (extensionBlock is Beef0004)
                                        {
                                            var b4 = extensionBlock as Beef0004;

                                            Log.Information("    Long name: {LongName}",b4.LongName);
                                            if (b4.LocalisedName.Length > 0)
                                            {
                                                Log.Information("    Localized name: {LocalisedName}",b4.LocalisedName);
                                            }

                                            if (b4.CreatedOnTime.HasValue)
                                            {
                                                Log.Information("    Created:     {CreatedOnTime}",b4.CreatedOnTime.Value);
                                            }
                                            else
                                            {
                                                Log.Information($"    Created:");
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
                                                Log.Information("    MFT entry/sequence #: {MftEntryNumber}/{MftSequenceNumber} (0x{MftEntryNumber2:X}/0x{MftSequenceNumber2:X})",b4.MFTInformation.MFTEntryNumber,b4.MFTInformation.MFTSequenceNumber,b4.MFTInformation.MFTEntryNumber,b4.MFTInformation.MFTSequenceNumber);
                                            }
                                        }
                                        else if (extensionBlock is Beef0025)
                                        {
                                            var b25 = extensionBlock as Beef0025;
                                            Log.Information("    Filetime 1: {FileTime1}, Filetime 2: {FileTime2}",b25.FileTime1.Value,b25.FileTime2.Value);
                                        }
                                        else if (extensionBlock is Beef0003)
                                        {
                                            var b3 = extensionBlock as Beef0003;
                                            Log.Information("    GUID: {Guid1} ({Guid1Folder})",b3.GUID1,b3.GUID1Folder);
                                        }
                                        else if (extensionBlock is Beef001a)
                                        {
                                            var b3 = extensionBlock as Beef001a;
                                            Log.Information("    File document type: {FileDocumentTypeString}",b3.FileDocumentTypeString);
                                        }
                                        else
                                        {
                                            Log.Information("    {ExtensionBlock}",extensionBlock);
                                        }

                                        extensionNumber74 += 1;
                                    }
                                }

                                break;
                            case "SHELLBAG0X23":
                            case "SHELLBAG0XC3":
                                break;
                            case "SHELLBAGZIPCONTENTS":
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
                                Log.Information("   MAC Address: {MacAddress}",tdb.MacAddress);
                                Log.Information("   MAC Vendor:  {Address}",GetVendorFromMac(tdb.MacAddress));
                                Log.Information("   Creation:    {Date}",tdb.CreationTime);
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

        catch (Exception ex)
        {
            _failedFiles.Add($"{lnkFile.Replace("\\\\?\\", "")} ==> ({ex.Message})");
            Log.Fatal(ex,"Error opening {LnkFile}. Message: {Message}",lnkFile,ex.Message);
            Console.WriteLine();
        }

        return null;
    }


    private static string GetVendorFromMac(string macAddress)
    {
        //00-00-00	XEROX CORPORATION
        //"00:14:22:0d:94:04"

        var mac = string.Join("", macAddress.Split(':').Take(3)).ToUpperInvariant();
        // .Replace(":", "-").ToUpper();

        var vendor = "(Unknown vendor)";

        if (MacList.ContainsKey(mac))
        {
            vendor = MacList[mac];
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
