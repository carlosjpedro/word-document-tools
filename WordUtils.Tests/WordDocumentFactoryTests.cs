using System;
using FluentAssertions;
using SampleApp.Utils;
using Xunit;

using static WordUtils.Tests.DocumentPaths;
namespace WordUtils.Tests
{
    [Collection("Integration")]
    public class WordDocumentFactoryTests
    {
        private readonly IWordDocumentFactory _factory = new WordDocumentFactory();
        
        [Fact]
        public void CanCreateDocumentFromStream()
        {
            var document = _factory.CreateDocument(TemplateFilePath);
            document.Should().NotBeNull();
        }

        [Fact]
        public void When_FileNotFound_ThrowArgException()
        {
            var exception = Assert.Throws<ArgumentException>(()=> _factory.CreateDocument(NotFoundFilePath));
            exception.Message.Should().Be("Template file not found (Parameter 'templateFile')");
        }
    }
}