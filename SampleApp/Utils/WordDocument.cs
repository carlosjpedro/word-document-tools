using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SampleApp.Utils
{
    public class WordDocument : IWordDocument, IDisposable
    {
        private readonly WordprocessingDocument _document;

        public WordDocument(string templateFile)
        {
            if (!File.Exists(templateFile))
                throw new FileNotFoundException("Template file not found.");

            try
            {
                _document = WordprocessingDocument.CreateFromTemplate(templateFile);
            }
            catch (ArgumentException)
            {
                throw new ArgumentException("Provided file is not a word template.", templateFile);
            }
        }

        public void Dispose()
        {
            _document.Dispose();
        }

        public void SetField(string fieldName, string newValue)
        {

            var field = _document.CustomFilePropertiesPart
                .Properties
                .Cast<CustomDocumentProperty>()
                .FirstOrDefault(x => x.Name == fieldName);

            var newProp = new CustomDocumentProperty();
            newProp.VTBString = new VTBString(newValue);
            newProp.FormatId = $"{{{Guid.NewGuid().ToString().ToUpper()}}}";
            newProp.Name = field.Name;
            field.Remove();
            
            _document.CustomFilePropertiesPart.Properties.AppendChild(newProp);
            _document.CustomFilePropertiesPart.Properties.Save();
            _document.SaveAs("Good.docx");
        }
    }
}