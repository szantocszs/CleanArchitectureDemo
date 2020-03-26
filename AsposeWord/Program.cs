using System;
using Aspose.Words;

namespace AsposeWord
{
    class Program
    {
        static void Main(string[] args)
        {


            AsposeTest();
        }

        private static void AsposeTest()
        {
            Console.WriteLine("Hello World!");
            var wrdf = new Document("rectification2.docx");

            // save in different formats
            // wrdf.Save("output.docx", Aspose.Words.SaveFormat.Docx);
            wrdf.Save("output.pdf", Aspose.Words.SaveFormat.Pdf);
            // wrdf.Save("output.html", Aspose.Words.SaveFormat.Html);
        }

        /// <summary>
        /// This code activates Aspose.Words license.
        /// If you don't specify a license, Aspose.Words will work in evaluation mode.
        /// </summary>
        private static void LicenseAsposeWords(string licenseFile)
        {
            Aspose.Words.License licenseWords = new Aspose.Words.License();
            licenseWords.SetLicense(licenseFile);
        }

    }
}
