using System;
using FluentAssertions;
using SampleApp.Utils;
using Xunit;

namespace WordUtils.Tests
{
    [Collection("Integration")]
    public class WordDocumentFactoryTests
    {
        private readonly IWordDocumentFactory _factory = new WordDocumentFactory();
        private const string TemplateFile = "./resource/template.dotx";
        private const string DocumentFile = "./resource/document.dotx";
        private const string InvalidFilePath = "./resource/no.file";

        [Fact]
        public void CanCreateDocumentFromStream()
        {
            var document = _factory.CreateDocument(TemplateFile);
            document.Should().NotBeNull();
        }

        [Fact]
        public void When_FileNotFound_ThrowArgException()
        {
            var exception = Assert.Throws<ArgumentException>(()=> _factory.CreateDocument(InvalidFilePath));
            exception.Message.Should().Be("Template file not found (Parameter 'templateFile')");
        }
    }
}