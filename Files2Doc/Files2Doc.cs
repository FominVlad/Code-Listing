using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Text;

namespace Files2Doc
{
    class Files2Doc
    {
        List<string> _extensions, _folders;
        string _outputFn;
        FileStream _fs = null;
        WordprocessingDocument _doc = null;
        MainDocumentPart _mainPart = null;
        Body _body = null;
        public Files2Doc(IEnumerable<string> extensions, IEnumerable<string> folders, string outputFn)
        {
            _extensions = extensions.ToList();
            _folders = folders.ToList();
            _outputFn = outputFn;
            File.Copy("Template.docx", _outputFn, true);
            _fs = new FileStream(_outputFn, FileMode.Open);
            _doc = WordprocessingDocument.Open(_fs, true);

            _mainPart = _doc.MainDocumentPart;
                        
            _body = _mainPart.Document.Body;
            

        }

        public void process()
        {
            foreach (var f in _folders) process(f);
        }

        public void save()
        {
            //_mainPart.Document.Save();
            //_doc.Package.Flush();
            _doc.Close();
            //_fs.Flush(true);
            _fs.Close();
            //_fs.Dispose();
        }

        public void process(string path)
        {
            Console.WriteLine("[Dir]: {0}", path);
            StyleDefinitionsPart part = _doc.MainDocumentPart.StyleDefinitionsPart;


            foreach (string file in GetFilesList(path))
            {
                var conentStr = File.ReadAllText(file).Trim();
                if (string.IsNullOrWhiteSpace(conentStr)) break;
                string[] content = conentStr.Split('\n');

                Console.WriteLine("[processing]: {0}", Path.GetFileName(file));
                Paragraph fileNameParagraph = new Paragraph(
                            new Run(
                                new Text(Path.GetFileName(file))));
                fileNameParagraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = "Heading2" });
                _body.Append(fileNameParagraph);
                
                var fileParagraph = new Paragraph();
                var fileParagraphRun = fileParagraph.AppendChild(new Run());
                foreach (var t in content)
                {
                    fileParagraphRun.AppendChild(new Text() { Text = t.TrimEnd('\r'), Space = SpaceProcessingModeValues.Preserve});
                    fileParagraphRun.AppendChild(new Break());
                }
                _body.AppendChild(fileParagraph);
            }
        }
        private List<string> GetFilesList(string rootDir)
        {
            var filesList = new List<string>();
            foreach (string extension in _extensions)
            {
                string[] files = Directory.GetFiles(rootDir, extension, SearchOption.AllDirectories);
                filesList.AddRange(files);
            }
            return filesList;
        }
    }
}
