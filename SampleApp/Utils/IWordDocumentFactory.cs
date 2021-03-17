using System.IO;

namespace SampleApp.Utils
{
    public interface IWordDocument
    {
        WordDocument ReplaceValue(string initialValue, string newValue);
        void UpdateDocument();
        Stream GetDocumentStream();
    }

    public interface IWordDocumentFactory
    {
        IWordDocument CreateDocument(string templateFile);
    }
}