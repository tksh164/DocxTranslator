using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxTranslator
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("command <Docx Path>");
                return;
            }

            var docxFilePath = args[0];
            var outputDocxFilePath = Path.Combine(Path.GetDirectoryName(docxFilePath), Path.GetFileNameWithoutExtension(docxFilePath) + ".ja.docx");

            DoTranslate(docxFilePath, outputDocxFilePath);
        }

        private static void DoTranslate(string docxFilePath, string outputDocxFilePath)
        {
            const string apiKey = "<API Key>";

            using (var wordClient = new WordDocumentClient(docxFilePath, outputDocxFilePath))
            {
                foreach (var paragraph in wordClient.GetParagraphs())
                {
                    var paragraphStyle = paragraph.ParagraphProperties?.ParagraphStyleId;
                    if (!IsExcludedStyleParagraph(paragraphStyle))
                    {
                        // Get the text for translate.
                        var text = GetConcatenatedTextInParagraph(paragraph);

                        using (var translatorClient = TranslatorClient.Create(apiKey))
                        {
                            Console.WriteLine(text);

                            // Translate.
                            var result = translatorClient.TranslateTextArray("en", "ja", new string[] { text });
                            result.Wait();

                            //
                            paragraph.RemoveAllChildren<Run>();

                            //
                            paragraph.AppendChild(new Run(new Text(result.Result[0])));

                            Console.WriteLine(result.Result[0]);

                            //text.Text = result.Result[0];
                            //Console.WriteLine(text.Text);
                        }


                        //foreach (var text in paragraph.Descendants<Text>())
                        //{
                        //    // Translate.
                        //    using (var translatorClient = TranslatorClient.Create(apiKey))
                        //    {
                        //        Console.WriteLine(text.Text);

                        //        var result = translatorClient.TranslateTextArray("en", "ja", new string[] { text.Text });
                        //        result.Wait();
                        //        text.Text = result.Result[0];

                        //        Console.WriteLine(text.Text);
                        //    }
                        //}
                    }
                }

                wordClient.Save();
            }
        }

        private static bool IsExcludedStyleParagraph(ParagraphStyleId paragraphStyle)
        {
            // Paragraph style is not specified. e.g. "Normal" style in Word.
            if (paragraphStyle == null) return false;

            if (paragraphStyle.Val.HasValue)
            {
                string styleName = paragraphStyle.Val.Value;

                string[] excludedStyleNames = new string[] { "Figure", "Code" };
                foreach (var excludedStyleName in excludedStyleNames)
                {
                    if (styleName.Equals(excludedStyleName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static string GetConcatenatedTextInParagraph(Paragraph paragraph)
        {
            StringBuilder builder = new StringBuilder();

            foreach (var text in paragraph.Descendants<Text>())
            {
                builder.Append(text.Text);
            }

            return builder.ToString();
        }
    }
}
