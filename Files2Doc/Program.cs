using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Files2Doc
{
    class Program
    {
        public class Options
        {
            [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
            public bool Verbose { get; set; }
            [Option('x', "extensions", Required = false, HelpText = "Extensions to process", Separator=',', Default = new[] { "*.cs", "*.aspx", "*.ascx", "*.ashx", "*.js", "*.html", "*.resx", "*.csproj", "*.xsd", "*.xss", "*.xss" })]
            public IEnumerable<string> Extensions { get; set; }
            [Option('f', "folders", Required = false, HelpText = "Folders to process", Default = new[] { "." })]
            public IEnumerable<string> Folders { get; set; }
            [Option('r', "remove-comments", Required = false, HelpText = "Remove comments from source files", Default = true)]
            public bool RemoveComments { get; set; }
            [Option('m', "remove-xml-comments", Required = false, HelpText = "Remove comments from XML files", Default = true)] 
            public bool RemoveXMLComments { get; set; }
            [Option('o', "output", Required = false, HelpText = "Output file name", Default = "output.docx")] 
            public string Output { get; set; }
        }

        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(o =>
                {
                    var cr = new CommentRemover.CommentRemover(o.RemoveXMLComments);
                    foreach (var path in o.Folders)
                    {
                        cr.Remove(path);
                    }

                    var doc = new Files2Doc(o.Extensions, o.Folders, o.Output);
                    doc.process();
                    doc.save();
                });

        }
    }
}
