using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocxTranslator
{
    internal class WordDocumentClient : IDisposable
    {
        private WordprocessingDocument wordDoc;

        public WordDocumentClient(string baseDocxFilePath, string outputDocxFilePath)
        {
            // Copy a base docx file as a putput docx file. 
            File.Copy(baseDocxFilePath, outputDocxFilePath, true);

            // Open a docx file.
            var openSettings = new OpenSettings()
            {
                AutoSave = false,
                MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.NoProcess, FileFormatVersions.Office2013),
                MaxCharactersInPart = 0,
            };

            wordDoc = WordprocessingDocument.Open(outputDocxFilePath, true, openSettings);
        }

        public void Dispose()
        {
            wordDoc.Dispose();
        }

        public void Save()
        {
            wordDoc.MainDocumentPart.Document.Save();
            wordDoc.Save();
        }

        public IEnumerable<Paragraph> GetParagraphs()
        {
            MainDocumentPart mainDocPart = wordDoc.MainDocumentPart;
            return mainDocPart.Document.Body.Descendants<Paragraph>();
        }
    }
}
