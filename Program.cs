using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace Excel2TextDiff
{
    internal class CommandLineOptions
    {
        [Option('t', SetName = "transform", HelpText = "transform excel to text file")]
        public bool IsTransform { get; set; }

        [Option('d', SetName = "diff", HelpText = "transform and diff file")]
        public bool IsDiff { get; set; }

        [Option('p', SetName = "diff", Required = false, HelpText = "3rd diff program. default TortoiseMerge")]
        public string DiffProgram { get; set; }

        [Option('f', SetName = "diff", Required = false,
            HelpText = "3rd diff program argument format. default is TortoiseMerge format:'/base:{0} /mine:{1}'")]
        public string DiffProgramArgumentFormat { get; set; }

        [Option('s', HelpText = "row cell separator for diff text file, default is ','. usage: -s \"|\"")]
        public string Separator { get; set; } = ",";

        [Option('a', HelpText = "align from which row, default value 0 means disable align. usage: -a 1")]
        public int AlignBeginRow { get; set; }

        [Value(0)] public IList<string> Files { get; set; }

        [Usage()]
        public static IEnumerable<Example> Examples => new List<Example>
        {
            new("transform to text",
                new CommandLineOptions { IsTransform = true, Files = new List<string> { "a.xlsx", "a.txt" } }),

            new("diff two excel file",
                new CommandLineOptions { IsDiff = true, Files = new List<string> { "a.xlsx", "b.xlsx" } }),

            new("diff two excel file with separator '|' and align from 3rd row",
                new CommandLineOptions
                {
                    IsDiff = true, Files = new List<string> { "a.xlsx", "b.xlsx" },
                    Separator = "|", AlignBeginRow = 3
                }),

            new("diff two excel file with TortoiseMerge, separator '|' and align from 3rd row",
                new CommandLineOptions
                {
                    IsDiff = true, DiffProgram = "TortoiseMerge", DiffProgramArgumentFormat = "/base:{0} /mine:{1}",
                    Files = new List<string> { "a.xlsx", "b.xlsx" },
                    Separator = "|", AlignBeginRow = 3
                }),
        };
    }

    internal static class Program
    {
        private static CommandLineOptions _options;

        private static void Main(string[] args)
        {
            _options = ParseOptions(args);

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            if (_options.IsTransform)
            {
                if (_options.Files.Count != 2)
                {
                    Console.WriteLine("Usage: Excel2TextDiff -t <excel file> <text file>");
                    Environment.Exit(1);
                }

                TransformToTextAndSave(_options.Files[0], _options.Files[1]);
            }
            else
            {
                if (_options.Files.Count != 2)
                {
                    Console.WriteLine("Usage: Excel2TextDiff -d <excel file 1> <excel file 2> ");
                    Environment.Exit(1);
                }

                var diffProgram = _options.DiffProgram ?? "TortoiseMerge.exe";

                var tempTxt1 = Path.GetTempFileName();
                TransformToTextAndSave(_options.Files[0], tempTxt1);

                var tempTxt2 = Path.GetTempFileName();
                TransformToTextAndSave(_options.Files[1], tempTxt2);

                var startInfo = new ProcessStartInfo
                {
                    FileName = diffProgram
                };
                var argsFormation = _options.DiffProgramArgumentFormat ?? "/base:{0} /mine:{1}";
                startInfo.Arguments = string.Format(argsFormation, tempTxt1, tempTxt2);
                Process.Start(startInfo);
            }
        }

        private static CommandLineOptions ParseOptions(string[] args)
        {
            var helpWriter = new StringWriter();
            var parser = new Parser(ps => { ps.HelpWriter = helpWriter; });

            var result = parser.ParseArguments<CommandLineOptions>(args);
            if (result.Tag == ParserResultType.NotParsed)
            {
                Console.Error.WriteLine(helpWriter.ToString());
                throw new Exception("parse command line failed");
                // Environment.Exit(1);
            }

            return ((Parsed<CommandLineOptions>)result).Value;
        }


        private static void TransformToTextAndSave(string excelFile, string outputTextFile)
        {
            var writer = new TextWriter
            {
                AlignBeginRow = _options.AlignBeginRow,
                Separator = _options.Separator
            };

            var ext = Path.GetExtension(excelFile);
            IReader reader = ext switch
            {
                ".xml" => (new XmlReader(excelFile)),
                _ => new ExcelReader(excelFile)
            };
            reader.Read(writer);

            writer.Save(outputTextFile);
        }
    }
}