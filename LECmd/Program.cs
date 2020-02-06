using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Xml;
using Alphaleonis.Win32.Filesystem;
using Exceptionless;
using ExtensionBlocks;
using Fclp;
using Fclp.Internals.Extensions;
using LECmd.Properties;
using Lnk;
using Lnk.ExtraData;
using Lnk.ShellItems;
using NLog;
using NLog.Config;
using NLog.Targets;
using ServiceStack;
using ServiceStack.Text;
using CsvWriter = CsvHelper.CsvWriter;
using Directory = Alphaleonis.Win32.Filesystem.Directory;
using File = Alphaleonis.Win32.Filesystem.File;
using Path = Alphaleonis.Win32.Filesystem.Path;
using ShellBag = Lnk.ShellItems.ShellBag;
using VolumeInfo = Lnk.VolumeInfo;

namespace LECmd
{
    internal class Program
    {
        private static Logger _logger;

        private static string _preciseTimeFormat = "yyyy-MM-dd HH:mm:ss.fffffff";

        private static FluentCommandLineParser<ApplicationArguments> _fluentCommandLineParser;

        private static List<string> _failedFiles;

        private static readonly Dictionary<string, string> MacList = new Dictionary<string, string>();

        private static List<LnkFile> _processedFiles;

        private static void LoadMACs()
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

        public static bool IsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static void Main(string[] args)
        {
            ExceptionlessClient.Default.Startup("FNWfyFuaUAPnVfofTZAhZOgeDG5lv7AnjYKNtsEJ");

            LoadMACs();

            SetupNLog();
            
            _logger = LogManager.GetCurrentClassLogger();
      

            _fluentCommandLineParser = new FluentCommandLineParser<ApplicationArguments>
            {
                IsCaseSensitive = false
            };

            _fluentCommandLineParser.Setup(arg => arg.File)
                .As('f')
                .WithDescription("File to process. Either this or -d is required");

            _fluentCommandLineParser.Setup(arg => arg.Directory)
                .As('d')
                .WithDescription("Directory to recursively process. Either this or -f is required");

            _fluentCommandLineParser.Setup(arg => arg.AllFiles)
                .As("all")
                .WithDescription(
                    "Process all files in directory vs. only files matching *.lnk. Default is FALSE\r\n").SetDefault(false);

            _fluentCommandLineParser.Setup(arg => arg.CsvDirectory)
                .As("csv")
                .WithDescription(
                    "Directory to save CSV formatted results to. Be sure to include the full path in double quotes");
            _fluentCommandLineParser.Setup(arg => arg.CsvName)
                .As("csvf")
                .WithDescription("File name to save CSV formatted results to. When present, overrides default name\r\n");


            _fluentCommandLineParser.Setup(arg => arg.XmlDirectory)
                .As("xml")
                .WithDescription(
                    "Directory to save XML formatted results to. Be sure to include the full path in double quotes");

            _fluentCommandLineParser.Setup(arg => arg.xHtmlDirectory)
                .As("html")
                .WithDescription(
                    "Directory to save xhtml formatted results to. Be sure to include the full path in double quotes");

            _fluentCommandLineParser.Setup(arg => arg.JsonDirectory)
                .As("json")
                .WithDescription(
                    "Directory to save json representation to. Use --pretty for a more human readable layout");

            _fluentCommandLineParser.Setup(arg => arg.JsonPretty)
                .As("pretty")
                .WithDescription(
                    "When exporting to json, use a more human readable layout. Default is FALSE\r\n").SetDefault(false);

            _fluentCommandLineParser.Setup(arg => arg.RemovableOnly)
                .As('r')
                .WithDescription(
                    "Only process lnk files pointing to removable drives. Default is FALSE")
                .SetDefault(false);

            _fluentCommandLineParser.Setup(arg => arg.Quiet)
                .As('q')
                .WithDescription(
                    "Only show the filename being processed vs all output. Useful to speed up exporting to json and/or csv. Default is FALSE\r\n")
                .SetDefault(false);


            _fluentCommandLineParser.Setup(arg => arg.NoTargetIDList)
                .As("nid")
                .WithDescription(
                    "Suppress Target ID list details from being displayed. Default is FALSE").SetDefault(false);


            _fluentCommandLineParser.Setup(arg => arg.NoExtraBlocks)
                .As("neb")
                .WithDescription(
                    "Suppress Extra blocks information from being displayed. Default is FALSE\r\n").SetDefault(false);

         

            _fluentCommandLineParser.Setup(arg => arg.DateTimeFormat)
    .As("dt")
    .WithDescription(
        "The custom date/time format to use when displaying time stamps. See https://goo.gl/CNVq0k for options. Default is: yyyy-MM-dd HH:mm:ss").SetDefault("yyyy-MM-dd HH:mm:ss");

            _fluentCommandLineParser.Setup(arg => arg.PreciseTimestamps)
   .As("mp")
   .WithDescription(
       $"Display higher precision for time stamps. Default is FALSE").SetDefault(false);


            var header =
                $"LECmd version {Assembly.GetExecutingAssembly().GetName().Version}" +
                "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
                "\r\nhttps://github.com/EricZimmerman/LECmd";
                

            var footer = @"Examples: LECmd.exe -f ""C:\Temp\foobar.lnk""" + "\r\n\t " +
                         @" LECmd.exe -f ""C:\Temp\somelink.lnk"" --json ""D:\jsonOutput"" --jsonpretty" + "\r\n\t " +
                         @" LECmd.exe -d ""C:\Temp"" --csv ""c:\temp"" --html c:\temp --xml c:\temp\xml -q" + "\r\n\t " +
                         @" LECmd.exe -f ""C:\Temp\some other link.lnk"" --nid --neb " + "\r\n\t " +
                         @" LECmd.exe -d ""C:\Temp"" --all" + "\r\n\t" + 
                         "\r\n\t"+
                         "  Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes\r\n";

            _fluentCommandLineParser.SetupHelp("?", "help")
                .WithHeader(header)
                .Callback(text => _logger.Info(text + "\r\n" + footer));

            var result = _fluentCommandLineParser.Parse(args);

            if (result.HelpCalled)
            {
                return;
            }

            if (result.HasErrors)
            {
                _logger.Error("");
                _logger.Error(result.ErrorText);

                _fluentCommandLineParser.HelpOption.ShowHelp(_fluentCommandLineParser.Options);

                return;
            }

            if (UsefulExtension.IsNullOrEmpty(_fluentCommandLineParser.Object.File) &&
                UsefulExtension.IsNullOrEmpty(_fluentCommandLineParser.Object.Directory))
            {
                _fluentCommandLineParser.HelpOption.ShowHelp(_fluentCommandLineParser.Options);

                _logger.Warn("Either -f or -d is required. Exiting");
                return;
            }

            if (UsefulExtension.IsNullOrEmpty(_fluentCommandLineParser.Object.File) == false &&
                !File.Exists(_fluentCommandLineParser.Object.File))
            {
                _logger.Warn($"File '{_fluentCommandLineParser.Object.File}' not found. Exiting");
                return;
            }

            if (UsefulExtension.IsNullOrEmpty(_fluentCommandLineParser.Object.Directory) == false &&
                !Directory.Exists(_fluentCommandLineParser.Object.Directory))
            {
                _logger.Warn($"Directory '{_fluentCommandLineParser.Object.Directory}' not found. Exiting");
                return;
            }

            _logger.Info(header);
            _logger.Info("");
            _logger.Info($"Command line: {string.Join(" ", Environment.GetCommandLineArgs().Skip(1))}\r\n");

            if (IsAdministrator() == false)
            {
                _logger.Fatal($"Warning: Administrator privileges not found!\r\n");
            }

            if (_fluentCommandLineParser.Object.PreciseTimestamps)
            {
                _fluentCommandLineParser.Object.DateTimeFormat = _preciseTimeFormat;
            }

   
            _processedFiles = new List<LnkFile>();


            _failedFiles = new List<string>();

            if (_fluentCommandLineParser.Object.File?.Length > 0)
            {
                LnkFile lnk = null;

                try
                {
                    lnk = ProcessFile(_fluentCommandLineParser.Object.File);
                    if (lnk != null)
                    {
                        _processedFiles.Add(lnk);
                    }
                }
                catch (UnauthorizedAccessException ua)
                {
                    _logger.Error(
                        $"Unable to access '{_fluentCommandLineParser.Object.File}'. Are you running as an administrator? Error: {ua.Message}");
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Error(
                        $"Error processing file '{_fluentCommandLineParser.Object.File}' Please send it to saericzimmerman@gmail.com. Error: {ex.Message}");
                    return;
                }
            }
            else
            {
                _logger.Info($"Looking for lnk files in '{_fluentCommandLineParser.Object.Directory}'");
                _logger.Info("");

                string[] lnkFiles = null;

                try
                {
                 

                    var f = new DirectoryEnumerationFilters();
                f.InclusionFilter = fsei =>
                {
                    var mask = ".lnk".ToUpperInvariant();
                    if (_fluentCommandLineParser.Object.AllFiles)
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

                f.RecursionFilter = entryInfo => !entryInfo.IsMountPoint && !entryInfo.IsSymbolicLink;

                f.ErrorFilter = (errorCode, errorMessage, pathProcessed) => true;

                var dirEnumOptions =
                    DirectoryEnumerationOptions.Files | DirectoryEnumerationOptions.Recursive |
                    DirectoryEnumerationOptions.SkipReparsePoints | DirectoryEnumerationOptions.ContinueOnException |
                    DirectoryEnumerationOptions.BasicSearch;

                var files2 =
                    Directory.EnumerateFileSystemEntries(_fluentCommandLineParser.Object.Directory, dirEnumOptions, f);

                lnkFiles = files2.ToArray();

//                    lnkFiles = Directory.GetFiles(_fluentCommandLineParser.Object.Directory, mask,
//                        SearchOption.AllDirectories);
                }
                catch (UnauthorizedAccessException ua)
                {
                    _logger.Error(
                        $"Unable to access '{_fluentCommandLineParser.Object.Directory}'. Error message: {ua.Message}");
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Error(
                        $"Error getting lnk files in '{_fluentCommandLineParser.Object.Directory}'. Error: {ex.Message}");
                    return;
                }

                _logger.Info($"Found {lnkFiles.Length:N0} files");
                _logger.Info("");

                var sw = new Stopwatch();
                sw.Start();

                foreach (var file in lnkFiles)
                {
                    var lnk = ProcessFile(file);
                    if (lnk != null)
                    {
                        _processedFiles.Add(lnk);
                    }
                }

                sw.Stop();

                if (_fluentCommandLineParser.Object.Quiet)
                {
                    _logger.Info("");
                }

                _logger.Info(
                    $"Processed {lnkFiles.Length - _failedFiles.Count:N0} out of {lnkFiles.Length:N0} files in {sw.Elapsed.TotalSeconds:N4} seconds");
                if (_failedFiles.Count > 0)
                {
                    _logger.Info("");
                    _logger.Warn("Failed files");
                    foreach (var failedFile in _failedFiles)
                    {
                        _logger.Info($"  {failedFile}");
                    }
                }
            }

            if (_processedFiles.Count > 0)
            {
                _logger.Info("");

                try
                {
                    CsvWriter csv = null;
                    StreamWriter sw = null;

                    if (_fluentCommandLineParser.Object.CsvDirectory?.Length > 0)
                    {
                        if (Directory.Exists(_fluentCommandLineParser.Object.CsvDirectory) == false)
                        {
                            _logger.Warn(
                                $"'{_fluentCommandLineParser.Object.CsvDirectory} does not exist. Creating...'");
                            Directory.CreateDirectory(_fluentCommandLineParser.Object.CsvDirectory);
                        }

                        var outName = $"{DateTimeOffset.Now.ToString("yyyyMMddHHmmss")}_LECmd_Output.csv";

                        if (_fluentCommandLineParser.Object.CsvName.IsNullOrEmpty() == false)
                        {
                            outName = Path.GetFileName(_fluentCommandLineParser.Object.CsvName);
                        }

                        var outFile = Path.Combine(_fluentCommandLineParser.Object.CsvDirectory, outName);

                        _fluentCommandLineParser.Object.CsvDirectory =
                            Path.GetFullPath(outFile);
                        _logger.Warn(
                            $"CSV output will be saved to '{Path.GetFullPath(outFile)}'");

                        try
                        {
                            sw = new StreamWriter(outFile);
                            csv = new CsvWriter(sw,CultureInfo.InvariantCulture);

                            csv.WriteHeader(typeof(CsvOut));
                            csv.NextRecord();
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(
                                $"Unable to open '{outFile}' for writing. CSV export canceled. Error: {ex.Message}");
                        }
                    }

                    if (_fluentCommandLineParser.Object.JsonDirectory?.Length > 0)
                    {
                        if (Directory.Exists(_fluentCommandLineParser.Object.JsonDirectory) == false)
                        {
                            _logger.Warn(
                                $"'{_fluentCommandLineParser.Object.JsonDirectory} does not exist. Creating...'");
                            Directory.CreateDirectory(_fluentCommandLineParser.Object.JsonDirectory);
                        }

                        _logger.Warn($"Saving json output to '{_fluentCommandLineParser.Object.JsonDirectory}'");
                    }

                    if (_fluentCommandLineParser.Object.XmlDirectory?.Length > 0)
                    {
                        {
                            if (Directory.Exists(_fluentCommandLineParser.Object.XmlDirectory) == false)
                            {
                                _logger.Warn(
                                    $"'{_fluentCommandLineParser.Object.XmlDirectory} does not exist. Creating...'");
                                Directory.CreateDirectory(_fluentCommandLineParser.Object.XmlDirectory);
                            }

                        }
                        _logger.Warn($"Saving XML output to '{_fluentCommandLineParser.Object.XmlDirectory}'");
                    }

                    XmlTextWriter xml = null;
                    StreamWriter jsonOut = null;

                    if (_fluentCommandLineParser.Object.JsonDirectory?.Length > 0) 
                    {
                        JsConfig.DateHandler = DateHandler.ISO8601;

                        if (Directory.Exists(_fluentCommandLineParser.Object.JsonDirectory) == false)
                        {
                            Directory.CreateDirectory(_fluentCommandLineParser.Object.JsonDirectory);
                        }

                        var outName =
                            $"{DateTimeOffset.UtcNow:yyyyMMddHHmmss}_LECmd_Output.json";
                        var outFile = Path.Combine(_fluentCommandLineParser.Object.JsonDirectory, outName);


                        jsonOut = new StreamWriter(new FileStream(outFile,FileMode.OpenOrCreate,FileAccess.Write),Encoding.UTF8);
                    }

                    if (_fluentCommandLineParser.Object.xHtmlDirectory?.Length > 0)
                    {

                        var outDir = Path.Combine(_fluentCommandLineParser.Object.xHtmlDirectory,
                            $"{DateTimeOffset.UtcNow:yyyyMMddHHmmss}_LECmd_Output_for_{_fluentCommandLineParser.Object.xHtmlDirectory.Replace(@":\", "_").Replace(@"\", "_")}");

                        if (Directory.Exists(outDir) == false)
                        {
                            Directory.CreateDirectory(outDir);
                        }

                        File.WriteAllText(Path.Combine(outDir, "normalize.css"), Resources.normalize);
                        File.WriteAllText(Path.Combine(outDir, "style.css"), Resources.style);

                        var outFile = Path.Combine(_fluentCommandLineParser.Object.xHtmlDirectory, outDir, "index.xhtml");

                        _logger.Warn($"Saving HTML output to '{outFile}'");

                        xml = new XmlTextWriter(outFile, Encoding.UTF8)
                        {
                            Formatting = Formatting.Indented,
                            Indentation = 4
                        };

                        xml.WriteStartDocument();

                        xml.WriteProcessingInstruction("xml-stylesheet", "href=\"normalize.css\"");
                        xml.WriteProcessingInstruction("xml-stylesheet", "href=\"style.css\"");

                        xml.WriteStartElement("document");
                    }

                    foreach (var processedFile in _processedFiles)
                    {
                        var o = GetCsvFormat(processedFile,false);

                        try
                        {
                            csv?.WriteRecord(o);
                            csv?.NextRecord();
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(
                                $"Error writing record for '{processedFile.SourceFile}' to '{_fluentCommandLineParser.Object.CsvDirectory}'. Error: {ex.Message}");
                        }

                        if (jsonOut != null)
                        {
                            var oldDt = _fluentCommandLineParser.Object.DateTimeFormat;

                            _fluentCommandLineParser.Object.DateTimeFormat = "o";

                            var cs = GetCsvFormat(processedFile,true);

                            _fluentCommandLineParser.Object.DateTimeFormat = oldDt;

                            if (_fluentCommandLineParser.Object.JsonPretty)
                            {
                                jsonOut.WriteLine(cs.Dump());
                            }
                            else
                            {
                                jsonOut.WriteLine(cs.ToJson());
                                
                            }
                        }


                        //XHTML
                        xml?.WriteStartElement("Container");
                        xml?.WriteElementString("SourceFile", o.SourceFile);
                        xml?.WriteElementString("SourceCreated", o.SourceCreated);
                        xml?.WriteElementString("SourceModified", o.SourceModified);
                        xml?.WriteElementString("SourceAccessed", o.SourceAccessed);

                        xml?.WriteElementString("TargetCreated", o.TargetCreated);
                        xml?.WriteElementString("TargetModified", o.TargetModified);
                        xml?.WriteElementString("TargetAccessed", o.TargetModified);

                        xml?.WriteElementString("FileSize", o.FileSize.ToString());
                        xml?.WriteElementString("RelativePath", o.RelativePath);
                        xml?.WriteElementString("WorkingDirectory", o.WorkingDirectory);
                        xml?.WriteElementString("FileAttributes", o.FileAttributes);
                        xml?.WriteElementString("HeaderFlags", o.HeaderFlags);
                        xml?.WriteElementString("DriveType", o.DriveType);
                        xml?.WriteElementString("VolumeSerialNumber", o.VolumeSerialNumber);
                        xml?.WriteElementString("VolumeLabel", o.VolumeLabel);
                        xml?.WriteElementString("LocalPath", o.LocalPath);
                        xml?.WriteElementString("NetworkPath", o.NetworkPath);
                        xml?.WriteElementString("CommonPath", o.CommonPath);
                        xml?.WriteElementString("Arguments", o.Arguments);

                        xml?.WriteElementString("TargetIDAbsolutePath", o.TargetIDAbsolutePath);

                        xml?.WriteElementString("TargetMFTEntryNumber", $"{o.TargetMFTEntryNumber}");
                        xml?.WriteElementString("TargetMFTSequenceNumber", $"{o.TargetMFTSequenceNumber}");

                        xml?.WriteElementString("MachineID", o.MachineID);
                        xml?.WriteElementString("MachineMACAddress", o.MachineMACAddress);
                        xml?.WriteElementString("MACVendor", o.MACVendor);

                        xml?.WriteElementString("TrackerCreatedOn", o.TrackerCreatedOn);

                        xml?.WriteElementString("ExtraBlocksPresent", o.ExtraBlocksPresent);

                        xml?.WriteEndElement();

                        if (_fluentCommandLineParser.Object.XmlDirectory?.Length > 0)
                        {
                            SaveXML(o, _fluentCommandLineParser.Object.XmlDirectory);
                        }
                    }


                    //Close CSV stuff
                    sw?.Flush();
                    sw?.Close();

                    //close json
                    jsonOut?.Flush();
                    jsonOut?.Close();

                    //Close XML
                    xml?.WriteEndElement();
                    xml?.WriteEndDocument();
                    xml?.Flush();
                }
                catch (Exception ex)
                {
                    _logger.Error(
                        $"Error exporting data! Error: {ex.Message}");
                }
            }
        }

        private static CsvOut GetCsvFormat(LnkFile lnk, bool nukeNulls)
        {

            string netPath = String.Empty;
            
            if (lnk.NetworkShareInfo != null)
            {
                netPath = lnk.NetworkShareInfo.NetworkShareName;
            }

            var csOut = new CsvOut
            {
                SourceFile = lnk.SourceFile,
                SourceCreated =  lnk.SourceCreated?.ToString(_fluentCommandLineParser.Object.DateTimeFormat) ?? string.Empty,
                SourceModified = lnk.SourceModified?.ToString(_fluentCommandLineParser.Object.DateTimeFormat) ?? string.Empty,
                SourceAccessed = lnk.SourceAccessed?.ToString(_fluentCommandLineParser.Object.DateTimeFormat) ?? string.Empty,
                TargetCreated = lnk.Header.TargetCreationDate.Year == 1601 ?  string.Empty:lnk.Header.TargetCreationDate.ToString(_fluentCommandLineParser.Object.DateTimeFormat),
                TargetModified = lnk.Header.TargetModificationDate.Year == 1601 ? string.Empty : lnk.Header.TargetModificationDate.ToString(_fluentCommandLineParser.Object.DateTimeFormat),
                TargetAccessed = lnk.Header.TargetLastAccessedDate.Year == 1601 ? String.Empty : lnk.Header.TargetLastAccessedDate.ToString(_fluentCommandLineParser.Object.DateTimeFormat),
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

                csOut.TrackerCreatedOn = tnbBlock?.CreationTime.ToString(_fluentCommandLineParser.Object.DateTimeFormat);

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
                            csOut.TargetMFTEntryNumber = $"0x{eb4.MFTInformation.MFTEntryNumber.Value.ToString("X")}";
                        }

                        if (eb4.MFTInformation.MFTSequenceNumber != null)
                        {
                            csOut.TargetMFTSequenceNumber =
                                $"0x{eb4.MFTInformation.MFTSequenceNumber.Value.ToString("X")}";
                        }
                    }
                }
            }

            if (nukeNulls)
            {
                foreach(var prop in csOut.GetType().GetProperties())
                {
                    
                    if ( prop.GetValue(csOut) as string == string.Empty)
                    {
                        prop.SetValue(csOut,null);
                    }
                }
            }


            return csOut;
        }

        private static void SaveXML(CsvOut csout, string outDir)
        {
            try
            {
                if (Directory.Exists(outDir) == false)
                {
                    Directory.CreateDirectory(outDir);
                }

                var outName =
                    $"{DateTimeOffset.UtcNow.ToString("yyyyMMddHHmmss")}_{Path.GetFileName(csout.SourceFile)}.xml";
                var outFile = Path.Combine(outDir, outName);


                File.WriteAllText(outFile, csout.ToXml());
            }
            catch (Exception ex)
            {
                _logger.Error($"Error exporting XML for '{csout.SourceFile}'. Error: {ex.Message}");
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

        private static LnkFile ProcessFile(string lnkFile)
        {
            if (_fluentCommandLineParser.Object.Quiet == false)
            {
                _logger.Warn($"Processing '{lnkFile}'");
                _logger.Info("");
            }

            var sw = new Stopwatch();
            sw.Start();

            try
            {
                var lnk = Lnk.Lnk.LoadFile(lnkFile);

                if (_fluentCommandLineParser.Object.RemovableOnly && lnk.VolumeInfo?.DriveType != VolumeInfo.DriveTypes.DriveRemovable)
                {
                    return null;
                }

                if (_fluentCommandLineParser.Object.Quiet == false)
                {
                    _logger.Error($"Source file: {lnk.SourceFile}");
                    _logger.Info($"  Source created:  {lnk.SourceCreated?.ToString(_fluentCommandLineParser.Object.DateTimeFormat) ?? string.Empty}");
                    _logger.Info($"  Source modified: {lnk.SourceModified?.ToString(_fluentCommandLineParser.Object.DateTimeFormat) ?? string.Empty}");
                    _logger.Info($"  Source accessed: {lnk.SourceAccessed?.ToString(_fluentCommandLineParser.Object.DateTimeFormat) ?? string.Empty}");
                    _logger.Info("");

                    _logger.Warn("--- Header ---");

                    var tc = lnk.Header.TargetCreationDate.Year == 1601 ? "" : lnk.Header.TargetCreationDate.ToString(_fluentCommandLineParser.Object.DateTimeFormat);
                    var tm = lnk.Header.TargetModificationDate.Year == 1601 ? "" : lnk.Header.TargetModificationDate.ToString(_fluentCommandLineParser.Object.DateTimeFormat);
                    var ta = lnk.Header.TargetLastAccessedDate.Year == 1601 ? "" : lnk.Header.TargetLastAccessedDate.ToString(_fluentCommandLineParser.Object.DateTimeFormat);

                    _logger.Info($"  Target created:  {tc}");
                    _logger.Info($"  Target modified: {tm}");
                    _logger.Info($"  Target accessed: {ta}");
                    _logger.Info("");
                    _logger.Info($"  File size: {lnk.Header.FileSize:N0}");
                    _logger.Info($"  Flags: {lnk.Header.DataFlags}");
                    _logger.Info($"  File attributes: {lnk.Header.FileAttributes}");

                    if (lnk.Header.HotKey.Length > 0)
                    {
                        _logger.Info($"  Hot key: {lnk.Header.HotKey}");
                    }

                    _logger.Info($"  Icon index: {lnk.Header.IconIndex}");
                    _logger.Info(
                        $"  Show window: {lnk.Header.ShowWindow} ({GetDescriptionFromEnumValue(lnk.Header.ShowWindow)})");

                    _logger.Info("");

                    if ((lnk.Header.DataFlags & Header.DataFlag.HasName) == Header.DataFlag.HasName)
                    {
                        _logger.Info($"Name: {lnk.Name}");
                    }

                    if ((lnk.Header.DataFlags & Header.DataFlag.HasRelativePath) == Header.DataFlag.HasRelativePath)
                    {
                        _logger.Info($"Relative Path: {lnk.RelativePath}");
                    }

                    if ((lnk.Header.DataFlags & Header.DataFlag.HasWorkingDir) == Header.DataFlag.HasWorkingDir)
                    {
                        _logger.Info($"Working Directory: {lnk.WorkingDirectory}");
                    }

                    if ((lnk.Header.DataFlags & Header.DataFlag.HasArguments) == Header.DataFlag.HasArguments)
                    {
                        _logger.Info($"Arguments: {lnk.Arguments}");
                    }

                    if ((lnk.Header.DataFlags & Header.DataFlag.HasIconLocation) == Header.DataFlag.HasIconLocation)
                    {
                        _logger.Info($"Icon Location: {lnk.IconLocation}");
                    }

                    if ((lnk.Header.DataFlags & Header.DataFlag.HasLinkInfo) == Header.DataFlag.HasLinkInfo)
                    {
                        _logger.Info("");
                        _logger.Error("--- Link information ---");
                        _logger.Info($"Flags: {lnk.LocationFlags}");

                        if (lnk.VolumeInfo != null)
                        {
                            _logger.Info("");
                            _logger.Warn(">>Volume information");
                            _logger.Info($"  Drive type: {GetDescriptionFromEnumValue(lnk.VolumeInfo.DriveType)}");
                            _logger.Info($"  Serial number: {lnk.VolumeInfo.VolumeSerialNumber}");

                            var label = lnk.VolumeInfo.VolumeLabel.Length > 0
                                ? lnk.VolumeInfo.VolumeLabel
                                : "(No label)";

                            _logger.Info($"  Label: {label}");
                        }

                        if (lnk.NetworkShareInfo != null)
                        {
                            _logger.Info("");
                            _logger.Warn("  Network share information");

                            if (lnk.NetworkShareInfo.DeviceName.Length > 0)
                            {
                                _logger.Info($"    Device name: {lnk.NetworkShareInfo.DeviceName}");
                            }

                            _logger.Info($"    Share name: {lnk.NetworkShareInfo.NetworkShareName}");

                            _logger.Info($"    Provider type: {lnk.NetworkShareInfo.NetworkProviderType}");
                            _logger.Info($"    Share flags: {lnk.NetworkShareInfo.ShareFlags}");
                            _logger.Info("");
                        }

                        if (lnk.LocalPath?.Length > 0)
                        {
                            _logger.Info($"  Local path: {lnk.LocalPath}");
                        }

                        if (lnk.CommonPath.Length > 0)
                        {
                            _logger.Info($"  Common path: {lnk.CommonPath}");
                        }
                    }

                    if (_fluentCommandLineParser.Object.NoTargetIDList)
                    {
                        _logger.Info("");
                        _logger.Warn($"(Target ID information suppressed. Lnk TargetID count: {lnk.TargetIDs.Count:N0})");
                    }

                    if (lnk.TargetIDs.Count > 0 && !_fluentCommandLineParser.Object.NoTargetIDList)
                    {
                        _logger.Info("");

                        var absPath = string.Empty;

                        foreach (var shellBag in lnk.TargetIDs)
                        {
                            absPath += shellBag.Value + @"\";
                        }

                        _logger.Error("--- Target ID information (Format: Type ==> Value) ---");
                        _logger.Info("");
                        _logger.Info($"  Absolute path: {GetAbsolutePathFromTargetIDs(lnk.TargetIDs)}");
                        _logger.Info("");

                        foreach (var shellBag in lnk.TargetIDs)
                        {
                            //HACK
                            //This is a total hack until i can refactor some shellbag code to clean things up

                            var val = shellBag.Value.IsNullOrEmpty() ? "(None)" : shellBag.Value;

                            _logger.Info($"  -{shellBag.FriendlyName} ==> {val}");

                            switch (shellBag.GetType().Name.ToUpper())
                            {
                                case "SHELLBAG0X36":
                                case "SHELLBAG0X32":
                                    var b32 = shellBag as ShellBag0X32;

                                    _logger.Info($"    Short name: {b32.ShortName}");
                                    if (b32.LastModificationTime.HasValue)
                                    {
                                        _logger.Info($"    Modified: {b32.LastModificationTime.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                    }
                                    else
                                    {
                                        _logger.Info($"    Modified:");
                                    }
                                    

                                    var extensionNumber32 = 0;
                                    if (b32.ExtensionBlocks.Count > 0)
                                    {
                                        _logger.Info($"    Extension block count: {b32.ExtensionBlocks.Count:N0}");
                                        _logger.Info("");
                                        foreach (var extensionBlock in b32.ExtensionBlocks)
                                        {
                                            _logger.Info(
                                                $"    --------- Block {extensionNumber32:N0} ({extensionBlock.GetType().Name}) ---------");
                                            if (extensionBlock is Beef0004)
                                            {
                                                var b4 = extensionBlock as Beef0004;

                                                _logger.Info($"    Long name: {b4.LongName}");
                                                if (b4.LocalisedName.Length > 0)
                                                {
                                                    _logger.Info($"    Localized name: {b4.LocalisedName}");
                                                }

                                                if (b4.CreatedOnTime.HasValue)
                                                {
                                                    _logger.Info($"    Created: {b4.CreatedOnTime.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                                }
                                                else
                                                {
                                                    _logger.Info($"    Created:");
                                                }

                                                if (b4.LastAccessTime.HasValue)
                                                {
                                                    _logger.Info($"    Last access: {b4.LastAccessTime.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                                }
                                                else
                                                {
                                                    _logger.Info($"    Last access: ");
                                                }
                                                
                                                if (b4.MFTInformation.MFTEntryNumber > 0)
                                                {
                                                    _logger.Info(
                                                        $"    MFT entry/sequence #: {b4.MFTInformation.MFTEntryNumber}/{b4.MFTInformation.MFTSequenceNumber} (0x{b4.MFTInformation.MFTEntryNumber:X}/0x{b4.MFTInformation.MFTSequenceNumber:X})");
                                                }
                                            }
                                            else if (extensionBlock is Beef0025)
                                            {
                                                var b25 = extensionBlock as Beef0025;
                                                _logger.Info(
                                                    $"    Filetime 1: {b25.FileTime1.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}, Filetime 2: {b25.FileTime2.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                            }
                                            else if (extensionBlock is Beef0003)
                                            {
                                                var b3 = extensionBlock as Beef0003;
                                                _logger.Info($"    GUID: {b3.GUID1} ({b3.GUID1Folder})");
                                            }
                                            else if (extensionBlock is Beef001a)
                                            {
                                                var b3 = extensionBlock as Beef001a;
                                                _logger.Info($"    File document type: {b3.FileDocumentTypeString}");
                                            }

                                            else
                                            {
                                                _logger.Info($"    {extensionBlock}");
                                            }

                                            extensionNumber32 += 1;
                                        }
                                    }

                                    break;
                                case "SHELLBAG0X31":

                                    var b3x = shellBag as Lnk.ShellItems.ShellBag0X31;

                                    _logger.Info($"    Short name: {b3x.ShortName}");
                                    if (b3x.LastModificationTime.HasValue)
                                    {
                                        _logger.Info(
                                            $"    Modified: {b3x.LastModificationTime.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                    }
                                    else
                                    {
                                        _logger.Info($"    Modified:");
                                    }

                                    var extensionNumber = 0;
                                    if (b3x.ExtensionBlocks.Count > 0)
                                    {
                                        _logger.Info($"    Extension block count: {b3x.ExtensionBlocks.Count:N0}");
                                        _logger.Info("");
                                        foreach (var extensionBlock in b3x.ExtensionBlocks)
                                        {
                                            _logger.Info(
                                                $"    --------- Block {extensionNumber:N0} ({extensionBlock.GetType().Name}) ---------");
                                            if (extensionBlock is Beef0004)
                                            {
                                                var b4 = extensionBlock as Beef0004;

                                                _logger.Info($"    Long name: {b4.LongName}");
                                                if (b4.LocalisedName.Length > 0)
                                                {
                                                    _logger.Info($"    Localized name: {b4.LocalisedName}");
                                                }

                                                if (b4.CreatedOnTime.HasValue)
                                                {
                                                    _logger.Info($"    Created: {b4.CreatedOnTime.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                                }
                                                else
                                                {
                                                    _logger.Info($"    Created:");
                                                }

                                                if (b4.LastAccessTime.HasValue)
                                                {
                                                    _logger.Info($"    Last access: {b4.LastAccessTime.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                                }
                                                else
                                                {
                                                    _logger.Info($"    Last access: ");
                                                }
                                                
                                                if (b4.MFTInformation.MFTEntryNumber > 0)
                                                {
                                                    _logger.Info(
                                                        $"    MFT entry/sequence #: {b4.MFTInformation.MFTEntryNumber}/{b4.MFTInformation.MFTSequenceNumber} (0x{b4.MFTInformation.MFTEntryNumber:X}/0x{b4.MFTInformation.MFTSequenceNumber:X})");
                                                }
                                            }
                                            else if (extensionBlock is Beef0025)
                                            {
                                                var b25 = extensionBlock as Beef0025;
                                                _logger.Info(
                                                    $"    Filetime 1: {b25.FileTime1.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}, Filetime 2: {b25.FileTime2.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                            }
                                            else if (extensionBlock is Beef0003)
                                            {
                                                var b3 = extensionBlock as Beef0003;
                                                _logger.Info($"    GUID: {b3.GUID1} ({b3.GUID1Folder})");
                                            }
                                            else if (extensionBlock is Beef001a)
                                            {
                                                var b3 = extensionBlock as Beef001a;
                                                _logger.Info($"    File document type: {b3.FileDocumentTypeString}");
                                            }
                                            else
                                            {
                                                _logger.Info($"    {extensionBlock}");
                                            }

                                            extensionNumber += 1;
                                        }
                                    }
                                    break;

                                case "SHELLBAG0X00":
                                    var b00 = shellBag as ShellBag0X00;

                                    if (b00.PropertyStore.Sheets.Count > 0)
                                    {
                                        _logger.Warn("  >> Property store (Format: GUID\\ID Description ==> Value)");
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
                                                        $"{Utils.GetDescriptionFromGuidAndKey(prop.GUID, int.Parse(propertyName.Key))}"
                                                            .PadRight(35);
                                                }

//                                                var suffix =
//                                                    $"{Utils.GetDescriptionFromGuidAndKey(prop.GUID, int.Parse(propertyName.Key))}"
//                                                        .PadRight(35);

                                                _logger.Info($"     {prefix} {suffix} ==> {propertyName.Value}");
                                            }
                                        }

                                        if (propCount == 0)
                                        {
                                            _logger.Warn("     (Property store is empty)");
                                        }
                                    }

                                    break;
                                case "SHELLBAG0X01":
                                    var baaaa1f = shellBag as ShellBag0X01;
                                    if (baaaa1f.DriveLetter.Length > 0)
                                    {
                                        _logger.Info($"  Drive letter: {baaaa1f.DriveLetter}");
                                    }
                                    break;
                                case "SHELLBAG0X1F":

                                    var b1f = shellBag as ShellBag0X1F;

                                    if (b1f.PropertyStore.Sheets.Count > 0)
                                    {
                                        _logger.Warn("  >> Property store (Format: GUID\\ID Description ==> Value)");
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
                                                
                                                _logger.Info($"     {prefix} {suffix} ==> {propertyName.Value}");
                                            }
                                        }

                                        if (propCount == 0)
                                        {
                                            _logger.Warn("     (Property store is empty)");
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
                                        _logger.Fatal(
                                            "Property stores found! Please email lnk file to saericzimmerman@gmail.com so support can be added!!");
                                    }

                                    break;
                                case "SHELLBAG0X74":
                                    var b74 = shellBag as ShellBag0X74;

                                    if (b74.LastModificationTime.HasValue)
                                    {
                                        _logger.Info($"    Modified: {b74.LastModificationTime.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                    }
                                    else
                                    {
                                        _logger.Info($"    Modified:");
                                    }

                                    var extensionNumber74 = 0;
                                    if (b74.ExtensionBlocks.Count > 0)
                                    {
                                        _logger.Info($"    Extension block count: {b74.ExtensionBlocks.Count:N0}");
                                        _logger.Info("");
                                        foreach (var extensionBlock in b74.ExtensionBlocks)
                                        {
                                            _logger.Info(
                                                $"    --------- Block {extensionNumber74:N0} ({extensionBlock.GetType().Name}) ---------");
                                            if (extensionBlock is Beef0004)
                                            {
                                                var b4 = extensionBlock as Beef0004;

                                                _logger.Info($"    Long name: {b4.LongName}");
                                                if (b4.LocalisedName.Length > 0)
                                                {
                                                    _logger.Info($"    Localized name: {b4.LocalisedName}");
                                                }

                                                if (b4.CreatedOnTime.HasValue)
                                                {
                                                    _logger.Info($"    Created: {b4.CreatedOnTime.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                                }
                                                else
                                                {
                                                    _logger.Info($"    Created:");
                                                }

                                                if (b4.LastAccessTime.HasValue)
                                                {
                                                    _logger.Info($"    Last access: {b4.LastAccessTime.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                                }
                                                else
                                                {
                                                    _logger.Info($"    Last access: ");
                                                }
                                                if (b4.MFTInformation.MFTEntryNumber > 0)
                                                {
                                                    _logger.Info(
                                                        $"    MFT entry/sequence #: {b4.MFTInformation.MFTEntryNumber}/{b4.MFTInformation.MFTSequenceNumber} (0x{b4.MFTInformation.MFTEntryNumber:X}/0x{b4.MFTInformation.MFTSequenceNumber:X})");
                                                }
                                            }
                                            else if (extensionBlock is Beef0025)
                                            {
                                                var b25 = extensionBlock as Beef0025;
                                                _logger.Info(
                                                    $"    Filetime 1: {b25.FileTime1.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}, Filetime 2: {b25.FileTime2.Value.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                            }
                                            else if (extensionBlock is Beef0003)
                                            {
                                                var b3 = extensionBlock as Beef0003;
                                                _logger.Info($"    GUID: {b3.GUID1} ({b3.GUID1Folder})");
                                            }
                                            else if (extensionBlock is Beef001a)
                                            {
                                                var b3 = extensionBlock as Beef001a;
                                                _logger.Info($"    File document type: {b3.FileDocumentTypeString}");
                                            }
                                            else
                                            {
                                                _logger.Info($"    {extensionBlock}");
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
                                    _logger.Fatal(
                                        $">> UNMAPPED Type! Please email lnk file to saericzimmerman@gmail.com so support can be added!");
                                    _logger.Fatal($">>{shellBag}");
                                    break;
                            }

                            _logger.Info("");
                        }
                        _logger.Error("--- End Target ID information ---");
                    }

                    if (_fluentCommandLineParser.Object.NoExtraBlocks)
                    {
                        _logger.Info("");
                        _logger.Warn(
                            $"(Extra blocks information suppressed. Lnk Extra block count: {lnk.ExtraBlocks.Count:N0})");
                    }

                    if (lnk.ExtraBlocks.Count > 0 && !_fluentCommandLineParser.Object.NoExtraBlocks)
                    {
                        _logger.Info("");
                        _logger.Error("--- Extra blocks information ---");
                        _logger.Info("");

                        foreach (var extraDataBase in lnk.ExtraBlocks)
                        {
                            switch (extraDataBase.GetType().Name)
                            {
                                case "ConsoleDataBlock":
                                    var cdb = extraDataBase as ConsoleDataBlock;
                                    _logger.Warn(">> Console data block");
                                    _logger.Info($"   Fill Attributes: {cdb.FillAttributes}");
                                    _logger.Info($"   Popup Attributes: {cdb.PopupFillAttributes}");
                                    _logger.Info(
                                        $"   Buffer Size (Width x Height): {cdb.ScreenWidthBufferSize} x {cdb.ScreenHeightBufferSize}");
                                    _logger.Info(
                                        $"   Window Size (Width x Height): {cdb.WindowWidth} x {cdb.WindowHeight}");
                                    _logger.Info($"   Origin (X/Y): {cdb.WindowOriginX}/{cdb.WindowOriginY}");
                                    _logger.Info($"   Font Size: {cdb.FontSize}");
                                    _logger.Info($"   Is Bold: {cdb.IsBold}");
                                    _logger.Info($"   Face Name: {cdb.FaceName}");
                                    _logger.Info($"   Cursor Size: {cdb.CursorSize}");
                                    _logger.Info($"   Is Full Screen: {cdb.IsFullScreen}");
                                    _logger.Info($"   Is Quick Edit: {cdb.IsQuickEdit}");
                                    _logger.Info($"   Is Insert Mode: {cdb.IsInsertMode}");
                                    _logger.Info($"   Is Auto Positioned: {cdb.IsAutoPositioned}");
                                    _logger.Info($"   History Buffer Size: {cdb.HistoryBufferSize}");
                                    _logger.Info($"   History Buffer Count: {cdb.HistoryBufferCount}");
                                    _logger.Info($"   History Duplicates Allowed: {cdb.HistoryDuplicatesAllowed}");
                                    _logger.Info("");
                                    break;
                                case "ConsoleFEDataBlock":
                                    var cfedb = extraDataBase as ConsoleFeDataBlock;
                                    _logger.Warn(">> Console FE data block");
                                    _logger.Info($"   Code page: {cfedb.CodePage}");
                                    _logger.Info("");
                                    break;
                                case "DarwinDataBlock":
                                    var ddb = extraDataBase as DarwinDataBlock;
                                    _logger.Warn(">> Darwin data block");
                                    _logger.Info($"   Application ID: {ddb.ApplicationIdentifierUnicode}");
                                    _logger.Info($"   Product code: {ddb.ProductCode}");
                                    _logger.Info($"   Feature name: {ddb.FeatureName}");
                                    _logger.Info($"   Component ID: {ddb.ComponentId}");
                                    _logger.Info("");
                                    break;
                                case "EnvironmentVariableDataBlock":
                                    var evdb = extraDataBase as EnvironmentVariableDataBlock;
                                    _logger.Warn(">> Environment variable data block");
                                    _logger.Info($"   Environment variables: {evdb.EnvironmentVariablesUnicode}");
                                    _logger.Info("");
                                    break;
                                case "IconEnvironmentDataBlock":
                                    var iedb = extraDataBase as IconEnvironmentDataBlock;
                                    _logger.Warn(">> Icon environment data block");
                                    _logger.Info($"   Icon path: {iedb.IconPathUni}");
                                    _logger.Info("");
                                    break;
                                case "KnownFolderDataBlock":
                                    var kfdb = extraDataBase as KnownFolderDataBlock;
                                    _logger.Warn(">> Known folder data block");
                                    _logger.Info(
                                        $"   Known folder GUID: {kfdb.KnownFolderId} ==> {kfdb.KnownFolderName}");
                                    _logger.Info("");
                                    break;
                                case "PropertyStoreDataBlock":
                                    var psdb = extraDataBase as PropertyStoreDataBlock;

                                    if (psdb.PropertyStore.Sheets.Count > 0)
                                    {
                                        _logger.Warn(
                                            ">> Property store data block (Format: GUID\\ID Description ==> Value)");
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

                                                _logger.Info($"   {prefix} {suffix} ==> {propertyName.Value}");
                                            }
                                        }

                                        if (propCount == 0)
                                        {
                                            _logger.Warn("   (Property store is empty)");
                                        }
                                    }
                                    _logger.Info("");
                                    break;
                                case "ShimDataBlock":
                                    var sdb = extraDataBase as ShimDataBlock;
                                    _logger.Warn(">> Shimcache data block");
                                    _logger.Info($"   Layer name: {sdb.LayerName}");
                                    _logger.Info("");
                                    break;
                                case "SpecialFolderDataBlock":
                                    var sfdb = extraDataBase as SpecialFolderDataBlock;
                                    _logger.Warn(">> Special folder data block");
                                    _logger.Info($"   Special Folder ID: {sfdb.SpecialFolderId}");
                                    _logger.Info("");
                                    break;
                                case "TrackerDataBaseBlock":
                                    var tdb = extraDataBase as TrackerDataBaseBlock;
                                    _logger.Warn(">> Tracker database block");
                                    _logger.Info($"   Machine ID: {tdb.MachineId}");
                                    _logger.Info($"   MAC Address: {tdb.MacAddress}");
                                    _logger.Info($"   MAC Vendor: {GetVendorFromMac(tdb.MacAddress)}");
                                    _logger.Info($"   Creation: {tdb.CreationTime.ToString(_fluentCommandLineParser.Object.DateTimeFormat)}");
                                    _logger.Info("");
                                    _logger.Info($"   Volume Droid: {tdb.VolumeDroid}");
                                    _logger.Info($"   Volume Droid Birth: {tdb.VolumeDroidBirth}");
                                    _logger.Info($"   File Droid: {tdb.FileDroid}");
                                    _logger.Info($"   File Droid birth: {tdb.FileDroidBirth}");
                                    _logger.Info("");
                                    break;
                                case "VistaAndAboveIdListDataBlock":
                                    var vdb = extraDataBase as VistaAndAboveIdListDataBlock;
                                    _logger.Warn(">> Vista and above ID List data block");

                                    foreach (var shellBag in vdb.TargetIDs)
                                    {
                                        var val = shellBag.Value.IsNullOrEmpty() ? "(None)" : shellBag.Value;
                                        _logger.Info($"   {shellBag.FriendlyName} ==> {val}");
                                    }

                                    _logger.Info("");
                                    break;
                                case "DamagedDataBlock":
                                    var dmg = extraDataBase as DamagedDataBlock;
                                    _logger.Warn("Damaged data block");
                                    _logger.Info($"   Original Signature: {dmg.OriginalSignature}");
                                    _logger.Info($"   Error Message: {dmg.ErrorMessage}");
                                
                                    break;
                            }
                        }
                    }
                }

                sw.Stop();

                if (_fluentCommandLineParser.Object.Quiet == false)
                {
                    _logger.Info("");
                }

                _logger.Info(
                    $"---------- Processed '{lnk.SourceFile}' in {sw.Elapsed.TotalSeconds:N8} seconds ----------");

                if (_fluentCommandLineParser.Object.Quiet == false)
                {
                    _logger.Info("\r\n");
                }

                return lnk;
            }

            catch (Exception ex)
            {
                _failedFiles.Add($"{lnkFile} ==> ({ex.Message})");
                _logger.Fatal($"Error opening '{lnkFile}'. Message: {ex.Message}");
                _logger.Info("");
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
        private static void SetupNLog()
        {
            if (File.Exists( Path.Combine(BaseDirectory,"Nlog.config")))
            {
                return;
            }
            var config = new LoggingConfiguration();
            var loglevel = LogLevel.Info;

            var layout = @"${message}";

            var consoleTarget = new ColoredConsoleTarget();

            config.AddTarget("console", consoleTarget);

            consoleTarget.Layout = layout;

            var rule1 = new LoggingRule("*", loglevel, consoleTarget);
            config.LoggingRules.Add(rule1);

            LogManager.Configuration = config;
        }
    }

    public sealed class CsvOut
    {
        public string SourceFile { get; set; }
        public string SourceCreated { get; set; }
        public string SourceModified { get; set; }
        public string SourceAccessed { get; set; }
        public string TargetCreated { get; set; }
        public string TargetModified { get; set; }
        public string  TargetAccessed { get; set; }
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
}