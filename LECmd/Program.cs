using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Fclp;
using Fclp.Internals.Extensions;
using Lnk;
using Microsoft.Win32;
using NLog;
using NLog.Config;
using NLog.Fluent;
using NLog.Targets;
using ServiceStack;
using ServiceStack.Text;

namespace LnkCmd
{
    class Program
    {
        private static Logger _logger;

        private static FluentCommandLineParser<ApplicationArguments> _fluentCommandLineParser;

        private static bool CheckForDotnet46()
        {
            using (
                var ndpKey =
                    RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                        .OpenSubKey("SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full\\"))
            {
                var releaseKey = Convert.ToInt32(ndpKey?.GetValue("Release"));

                return releaseKey >= 393295;
            }
        }

        static void Main(string[] args)
        {
            SetupNLog();

            _logger = LogManager.GetCurrentClassLogger();

            if (!CheckForDotnet46())
            {
                _logger.Warn(".net 4.6 not detected. Please install .net 4.6 and try again.");
                return;
            }

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

            _fluentCommandLineParser.Setup(arg => arg.JsonDirectory)
              .As("json")
              .WithDescription(
                  "Directory to save json representation to. Use --pretty for a more human readable layout");



            _fluentCommandLineParser.Setup(arg => arg.JsonPretty)
                .As("pretty")
                .WithDescription(
                    "When exporting to json, use a more human readable layout").SetDefault(false);



            var header =
               $"LECmd version {Assembly.GetExecutingAssembly().GetName().Version}" +
               "\r\n\r\nAuthor: Eric Zimmerman (saericzimmerman@gmail.com)" +
               "\r\nhttps://github.com/EricZimmerman/LECmd";

            var footer = @"Examples: LECmd.exe -f ""C:\Temp\foobar.lnk""" + "\r\n\t " +
                         @" LECmd.exe -f ""C:\Temp\somelink.lnk"" --json ""D:\jsonOutput"" --jsonpretty" +
                         //"\r\n\t " +
                         @" LECmd.exe -d ""C:\Temp"" " + "\r\n\t ";
            //@" LECmd.exe -d ""C:\Temp"" --csv ""c:\temp\prefetch_out.csv"" --local" + "\r\n\t " +
            //                         @" PECmd.exe -f ""C:\Temp\someOtherFile.txt"" --lr cc -sa" + "\r\n\t " +
            //                         @" PECmd.exe -f ""C:\Temp\someOtherFile.txt"" --lr cc -sa -m 15 -x 22" + "\r\n\t " +
            //@" LECmd.exe -d ""C:\Windows\Prefetch""" + "\r\n\t ";

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
            _logger.Info($"Command line: {string.Join(" ", Environment.GetCommandLineArgs().Skip(1))}");

            if (_fluentCommandLineParser.Object.File?.Length > 0)
            {
                Lnk.LnkFile lnk = null;

                try
                {
                    lnk = LoadFile(_fluentCommandLineParser.Object.File);
                }
                catch (UnauthorizedAccessException ex)
                {
                    _logger.Error(
                        $"Unable to access '{_fluentCommandLineParser.Object.File}'. Are you running as an administrator?");
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Error(
                        $"Error getting lnk files in '{_fluentCommandLineParser.Object.Directory}'. Error: {ex.Message}");
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
                    lnkFiles = Directory.GetFiles(_fluentCommandLineParser.Object.Directory, "*.lnk",
                    SearchOption.AllDirectories);
                }
                catch (UnauthorizedAccessException)
                {
                    _logger.Error($"Unable to access '{_fluentCommandLineParser.Object.Directory}'. Are you running as an administrator?");
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Error($"Error getting lnk files in '{_fluentCommandLineParser.Object.Directory}'. Error: {ex.Message}");
                    return;
                }

                _logger.Info($"Found {lnkFiles.Length} lnk files");
                _logger.Info("");

                var sw = new Stopwatch();
                sw.Start();

                foreach (var file in lnkFiles)
                {
                    var lnk = LoadFile(file);
                }

                sw.Stop();

    


                _logger.Info($"Processed {lnkFiles.Length:N0} files in {sw.Elapsed.TotalSeconds:N4} seconds");
            }

        }

        private static void DumpToJson(LnkFile lnk, bool pretty, string outFile)
        {
            if (pretty)
            {
                File.WriteAllText(outFile, lnk.Dump());
            }
            else
            {
                File.WriteAllText(outFile, lnk.ToJson());
            }
        }

        private static void SaveJson(LnkFile lnk, bool pretty, string outDir)
        {
            try
            {
                if (Directory.Exists(outDir) == false)
                {
                    Directory.CreateDirectory(outDir);
                }

                var outName =
                    $"{DateTimeOffset.UtcNow.ToString("yyyyMMddHHmmss")}_{Path.GetFileName(lnk.SourceFile)}.json";
                var outFile = Path.Combine(outDir, outName);

                _logger.Info("");
                _logger.Info($"Saving json output to '{outFile}'");

                DumpToJson(lnk, pretty, outFile);
            }
            catch (Exception ex)
            {
                _logger.Error($"Error exporting json. Error: {ex.Message}");
            }
        }

        private static  string GetDescriptionFromEnumValue(Enum value)
        {
            var attribute = value.GetType()
                .GetField(value.ToString())
                .GetCustomAttributes(typeof(DescriptionAttribute), false)
                .SingleOrDefault() as DescriptionAttribute;
            return attribute == null ? value.ToString() : attribute.Description;
        }

        private static LnkFile LoadFile(string lnkFile)
        {
            _logger.Warn($"Processing '{lnkFile}'");
            _logger.Info("");

            var sw = new Stopwatch();
            sw.Start();

            try
            {
                var lnk = Lnk.Lnk.LoadFile(lnkFile);

                _logger.Error($"Source file: {lnk.SourceFile}");
                _logger.Info($"Source created: {lnk.SourceCreated}");
                _logger.Info($"Source modified: {lnk.SourceModified}");
                _logger.Info($"Source accessed: {lnk.SourceAccessed}");
                _logger.Info("");
                _logger.Warn("--- Header ---");
                _logger.Info($"  File size: {lnk.Header.FileSize:N0}");
                _logger.Info($"  Flags: {lnk.Header.DataFlags}");
                _logger.Info($"  File attributes: {lnk.Header.FileAttributes}");

                if (lnk.Header.HotKey.Length > 0)
                {
                    _logger.Info($"  Hot key: {lnk.Header.HotKey}");
                }

                _logger.Info($"  Icon index: {lnk.Header.IconIndex}");
                _logger.Info($"  Show window: {lnk.Header.ShowWindow} ({GetDescriptionFromEnumValue(lnk.Header.ShowWindow)})");
                _logger.Info($"  Target created: {lnk.Header.TargetCreationDate}");
                _logger.Info($"  Target modified: {lnk.Header.TargetLastAccessedDate}");
                _logger.Info($"  Target accessed: {lnk.Header.TargetModificationDate}");
                _logger.Info("");

                if (lnk.TargetIDs.Count > 0)
                {
                    _logger.Info("");
                    _logger.Warn("--- Target ID information ---");
                    foreach (var shellBag in lnk.TargetIDs)
                    {
                        _logger.Info($">>{shellBag}");
                    }
                }

                if ((lnk.Header.DataFlags & Header.DataFlag.HasLinkInfo) == Header.DataFlag.HasLinkInfo)
                {
                    _logger.Info("");
                    _logger.Warn("--- Link information ---");
                    _logger.Info($"Location flags: {lnk.LocationFlags}");

                    if (lnk.VolumeInfo != null)
                    {
                        _logger.Info("");
                        _logger.Info(">>Volume information");
                        _logger.Info($"Drive type: {lnk.VolumeInfo.DriveType}");
                        _logger.Info($"Serial number: {lnk.VolumeInfo.DriveSerialNumber}");

                        var label = lnk.VolumeInfo.VolumeLabel.Length > 0 ? lnk.VolumeInfo.VolumeLabel : "(No label)";

                        _logger.Info($"Label: {label}");
                    }

                    if (lnk.LocalPath?.Length > 0)
                    {
                        _logger.Info($"Local path: {lnk.LocalPath}");
                    }

                    if (lnk.NetworkShareInfo != null)
                    {
                        _logger.Info("");
                        _logger.Warn("Network share information");

                        if (lnk.NetworkShareInfo.DeviceName.Length > 0)
                        {
                            _logger.Info($"Device name: {lnk.NetworkShareInfo.DeviceName}");
                        }

                        _logger.Info($"Share name: {lnk.NetworkShareInfo.NetworkShareName}");

                        _logger.Info($"Provider type: {lnk.NetworkShareInfo.NetworkProviderType}");
                        _logger.Info($"Share flags: {lnk.NetworkShareInfo.ShareFlags}");

                    }

                    if (lnk.CommonPath.Length > 0)
                    {
                        _logger.Info($"Common path: {lnk.CommonPath}");
                    }
                }

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

                if (lnk.ExtraBlocks.Count > 0)
                {
                    _logger.Info("");
                    _logger.Warn("--- Extra blocks information ---");
                    foreach (var extraDataBase in lnk.ExtraBlocks)
                    {
                        _logger.Info($">>{extraDataBase}");
                        _logger.Info("");
                    }
                }

                sw.Stop();

                if (_fluentCommandLineParser.Object.JsonDirectory?.Length > 0)
                {
                    SaveJson(lnk, _fluentCommandLineParser.Object.JsonPretty,
                        _fluentCommandLineParser.Object.JsonDirectory);
                }

                _logger.Info("");
                _logger.Info(
                    $"---------- Processed '{lnk.SourceFile}' in {sw.Elapsed.TotalSeconds:N4} seconds ----------");
                _logger.Info("\r\n");
                return lnk;
            }

            catch (Exception ex)
            {
                _logger.Error($"Error opening '{lnkFile}'. Message: {ex.Message}");
                _logger.Info("");
            }


            return null;
        }

        private static void SetupNLog()
        {
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

    internal class ApplicationArguments
    {
        public string File { get; set; }
        public string Directory { get; set; }
        //   public string Keywords { get; set; }
         public string JsonDirectory { get; set; }
          public bool JsonPretty { get; set; }
        //  public bool LocalTime { get; set; }
        //  public string CsvFile { get; set; }
        //  public bool Quiet { get; set; }
    }
}
