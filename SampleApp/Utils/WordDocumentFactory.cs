using System;
using System.IO;

namespace SampleApp.Utils
{
    public class WordDocumentFactory : IWordDocumentFactory
    {
        public IWordDocument CreateDocument(string templateFile)
        {
            if (!File.Exists(templateFile))
                throw new ArgumentException("Template file not found", nameof(templateFile));
            
            return new WordDocument(templateFile);
        }
    }
}