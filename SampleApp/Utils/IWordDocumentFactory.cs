using System.IO;

namespace SampleApp.Utils
{
    public interface IWordDocument {}

    public interface IWordDocumentFactory
    {
        IWordDocument CreateDocument(string templateFile);
    }
}