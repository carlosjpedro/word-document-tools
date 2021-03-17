using System;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using SampleApp.Utils;
using Xunit;
using static WordUtils.Tests.DocumentPaths;

namespace WordUtils.Tests
{
    [Collection("Integration")]
    public class WordDocumentTests
    {
        [Fact]
        public void When_FileDoesNotExist_Then_ThrowException()
        {
            var exception = Assert
                .Throws<FileNotFoundException>(() =>
                    new WordDocument(NotFoundFilePath));

            exception.Message
                .Should()
                .Be("Template file not found.");
        }

        [Fact]
        public void When_FileIsNotTemplate_Then_ThrowException()
        {
            var exception = Assert.Throws<ArgumentException>(
                () => new WordDocument(NonTemplateFilePath));

            exception.Message
                .Should()
                .Be("Provided file is not a word template. (Parameter './resource/sample.txt')");
        }

        [Fact]
        public void When_ValueIsReplaced_Then_ExportedDocumentHasExpectedValue()
        {
            var value = "New Updated Value";
            using var document = new WordDocument(TemplateFilePath);
            document
                .ReplaceValue("Test", value)
                .UpdateDocument();

            var documentText = GetDocumentTextFromStream(document.GetDocumentStream());
            documentText.Should().Contain(value);
        }

        [Fact]
        public void When_ValueDoesNotExists_DontMakeChanges()
        {
            var value = "Sample Text";
            var initialText = GetDocumentTextFromDocxFile(TemplateFilePath);
            
            using var document = new WordDocument(TemplateFilePath);
            document
                .ReplaceValue("NotFoundText", value)
                .UpdateDocument();
            var updatedText =  GetDocumentTextFromStream(document.GetDocumentStream());

            updatedText.Should().Be(initialText);
        }

        private string GetDocumentTextFromStream(Stream stream)
        {
            using var openXmlWordDocument = WordprocessingDocument.Open(stream, false);
            using var streamReader = new StreamReader(openXmlWordDocument.MainDocumentPart.GetStream());
            return streamReader.ReadToEnd();
        }

        private string GetDocumentTextFromDocxFile(string file)
        {
            using var openXmlWordDocument = WordprocessingDocument.Open(file, false);
            using var streamReader = new StreamReader(openXmlWordDocument.MainDocumentPart.GetStream());
            return streamReader.ReadToEnd();
        }
    }
}