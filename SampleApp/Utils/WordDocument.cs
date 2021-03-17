using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SampleApp.Utils
{
    public class WordDocument : IWordDocument, IDisposable
    {
        private readonly WordprocessingDocument _document;
        private string _documentText;

        public WordDocument(string templateFile)
        {
            if (!File.Exists(templateFile))
                throw new FileNotFoundException("Template file not found.");

            try
            {
                _document = WordprocessingDocument.Open(templateFile, true);
                using var reader = new StreamReader(_document.MainDocumentPart.GetStream());
                _documentText = reader.ReadToEnd();
            }
            catch (InvalidDataException)
            {
                throw new ArgumentException("Provided file is not a word template.", templateFile);
            }
            catch (FileFormatException)
            {
                throw new ArgumentException("Provided file is not a word template.", templateFile);
            }
            catch (ArgumentException)
            {
                throw new ArgumentException("Provided file is not a word template.", templateFile);
            }
        }


        public WordDocument ReplaceValue(string initialValue, string newValue)
        {
            _documentText =   Regex.Replace(_documentText,initialValue ,newValue);
            return this;
        }

        public void UpdateDocument()
        {
            using var writer = new StreamWriter(_document.MainDocumentPart.GetStream(FileMode.Create));
            writer.Write(_documentText);
        }

        public Stream GetDocumentStream()
        {
            var outStream = new MemoryStream();
            _document.Clone(outStream);
             return outStream;
        }

        private void Dispose(bool disposing)
        {
            if (!disposing) return;
            _document?.Dispose();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}