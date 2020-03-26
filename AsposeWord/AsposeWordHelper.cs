using DevExpress.XtraPrinting;
using DevExpress.XtraRichEdit;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Aspose.Words.Replacing;
using Fortuna.Services.TemplateStorage.Interface.Dtos;
using Document = Aspose.Words.Document;

namespace AsposeWord
{
    public static class AsposeWordHelper
    {
        /// <summary>Placeholder elejét jelző string.</summary>
        private static string placeholderStart = @"@@";
        /// <summary>Placeholder végét jelző string.</summary>
        private static string placeholderEnd = @"##";
        /// <summary>Placeholdert leíró reguláris kifejezés.</summary>
        private static string placeholderRegexString = $"{placeholderStart}" + @"{1}\w*" + $"{placeholderEnd}" + @"{1}";
        /// <summary>Placeholder regex.</summary>
        private static Regex placeholderRegex = new Regex(placeholderRegexString);

        /// <summary>
        /// Ellenőrzi, hogy docx formátumú-e.
        /// </summary>
        /// <param name="stream">Word dokumentum stream.</param>
        /// <returns>Visszaadja, hogy docx-e.</returns>
        public static bool IsWordDocument(Stream stream)
        {
            try
            {
                var doc = new Document(stream);

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Ellenőrzi, hogy docx formátumú-e.
        /// </summary>
        /// <param name="byteArray">Word dokumentum.</param>
        /// <returns>Visszaadja, hogy docx-e.</returns>
        public static bool IsWordDocument(byte[] byteArray)
        {
            try
            {
                using (var stream = new MemoryStream(byteArray, false))
                {
                    return IsWordDocument(stream);
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Placeholder cseréje.
        /// </summary>
        /// <param name="stream">Word dokumentum stream.</param>
        /// <param name="property">Placeholder neve.</param>
        /// <param name="value">Placeholder értéke.</param>
        public static void SetPlaceHolder(Stream stream, string property, string value)
        {
            var doc = new Document(stream);
            {
                var placeHolder = placeholderStart + property + placeholderEnd;
                doc.Range.Replace(placeHolder, value);
            }
        }

        /// <summary>
        /// Placeholder cseréje listával.
        /// </summary>
        /// <param name="stream">Word dokumentum stream.</param>
        /// <param name="property">Placeholder neve.</param>
        /// <param name="valueList">Placeholder értékei.</param>
        public static void SetPlaceHolder(Stream stream, string property, List<string> valueList)
        {
            using (var document = WordprocessingDocument.Open(stream, true))
            {
                var placeHolder = placeholderStart + property + placeholderEnd;

                var texts = document.MainDocumentPart.Document.Body.Descendants<Text>().Where(x => x.Text.Contains(placeHolder)).ToList();
                texts.ForEach(text =>
                {
                    var appendParts = valueList.Select(x => new Text(x)).ToList();
                    var first = text.Parent.Parent;
                    var afterThis = text.Parent.Parent;

                    text.Text = string.Empty;
                    appendParts.ForEach(part =>
                    {
                        var cloned = text.Parent.Parent.CloneNode(true);
                        var clonedRun = cloned.ChildElements.OfType<Run>().Single();
                        var clonedText = clonedRun.ChildElements.OfType<Text>().Single();
                        clonedRun.RemoveChild(clonedText);
                        clonedRun.AppendChild(part);
                        text.Parent.Parent.Parent.InsertAfter(cloned, afterThis);
                        afterThis = cloned;
                    });

                    text.Parent.Parent.Parent.RemoveChild(first);
                });

                document.Save();
                document.Close();
            }
        }

        /// <summary>
        /// DOCX konvertálása PDF-be.
        /// </summary>
        /// <param name="docx"></param>
        /// <param name="pdf"></param>
        public static byte[] ExportToPdf(Stream docx)
        {
            using (RichEditDocumentServer server = new RichEditDocumentServer())
            {
                using (var pdfStream = new MemoryStream())
                {
                    docx.Position = 0;
                    server.LoadDocument(docx, DevExpress.XtraRichEdit.DocumentFormat.OpenXml);

                    //Specify export options:
                    PdfExportOptions options = new PdfExportOptions();
                    options.Compressed = false;
                    options.ImageQuality = PdfJpegImageQuality.Highest;

                    //Export the document to the stream: 
                    server.ExportToPdf(pdfStream, options);

                    return pdfStream.ToArray();
                }
            }
        }

        /// <summary>
        /// DOCX dokumentumban lecseréli a megadott propertyket, és PDF-fé konvertálja.
        /// </summary>
        /// <param name="template">A dokumentum.</param>
        /// <param name="nameValueList">Property név-érték lista.</param>
        /// <returns>A generált PDF.</returns>
        public static byte[] ReplacePlaceholdersAndExportToPdf(byte[] template, List<PlaceHolderNameValueDtoBase> nameValueList)
        {
            if (!ValidatePlaceholders(template, nameValueList))
            {
                throw new ArgumentException("A kitöltendő adatok nem egyeznek meg a sablonban definiáltakkal!", nameof(nameValueList));
            }

            using (var templateStream = new MemoryStream())
            {
                templateStream.Write(template, 0, template.Length);

                foreach (var nameValue in nameValueList)
                {
                    if (nameValue is ListPlaceHolderNameValueDto listDto)
                    {
                        SetPlaceHolder(templateStream, listDto.Name, listDto.Value);
                    }
                    else if (nameValue is SimplePlaceHolderNameValueDto simpleDto)
                    {
                        SetPlaceHolder(templateStream, simpleDto.Name, simpleDto.Value);
                    }
                }

                return ExportToPdf(templateStream);
            }
        }

        /// <summary>
        /// Ellenőrzi, hogy a PDF generálás kérelmében az összes változó adat meg van adva.
        /// </summary>
        /// <param name="template">A dokumentum.</param>
        /// <param name="nameValueList">Property név-érték lista.</param>
        /// <returns>Igaz, ha a név-érték lista a sablon összes változó adatát tartalmazza.</returns>
        public static bool ValidatePlaceholders(byte[] template, List<PlaceHolderNameValueDtoBase> nameValueList)
        {
            var placeholders = GetPlaceholdersWithRegexp(template);
            var names = nameValueList.Select(x => x.Name).Distinct().ToList();
            return placeholders.OrderBy(x => x).SequenceEqual(names.OrderBy(x => x));
        }

        /// <summary>Kikeresi a megadott dokumentumban található placeholder azonosítókat, egy reguláris kifejezés alapján.</summary>
        /// <param name="template">A dokumentum.</param>
        /// <returns>Placeholder azonosítók listája.</returns>
        private static List<string> GetPlaceholdersWithRegexp(byte[] template)
        {
            var ret = new List<string>();

            using (var templateStream = new MemoryStream())
            {
                templateStream.Write(template, 0, template.Length);

                using (var document = WordprocessingDocument.Open(templateStream, false))
                {
                    var placeholderMatches = document.MainDocumentPart.Document.Body.Descendants<Text>().Select(t => placeholderRegex.Matches(t.Text)).ToList();

                    document.MainDocumentPart.HeaderParts.ToList().ForEach(x =>
                    {
                        placeholderMatches.AddRange(x.Header.Descendants<Text>().Select(t => placeholderRegex.Matches(t.Text)).ToList());
                    });

                    ret = GetPlaceholdersFromRegexpMatches(placeholderMatches);
                }
            }

            return ret.Distinct().ToList();
        }

        /// <summary>Megkeresi a placeholdereket a regex találatokból.</summary>
        /// <param name="matches">Regex találatok listája.</param>
        /// <returns>Placeholderek listája.</returns>
        private static List<string> GetPlaceholdersFromRegexpMatches(List<MatchCollection> matches)
        {
            var placeholderTexts = new List<string>();

            foreach (var match in matches)
            {
                placeholderTexts.AddRange(match.Cast<Match>().Select(m => m.Value).ToArray());
            }

            var ret = placeholderTexts.Select(t => t.Substring(placeholderStart.Length, t.Length - placeholderStart.Length - placeholderEnd.Length)).ToList();
            return ret;
        }
    }
}
