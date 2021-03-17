using System;
using System.IO;
using FluentAssertions;
using SampleApp.Utils;
using Xunit;

namespace WordUtils.Tests
{
    [Collection("Integration")]
    public class WordDocumentTests
    {
        private const string TemplateFilePath = "./resource/template.dotx";
        private const string NonTemplateFilePath = "./resource/sample.txt";
        private const string NotFoundFilePath = "./resource/no.file";
        
        [Fact]
        public void When_FileDoesNotExist_ThrowException()
        {
            var exception =  Assert
                .Throws<FileNotFoundException>(()=>
                 new WordDocument(NotFoundFilePath));

            exception.Message
                .Should()
                .Be("Template file not found.");
        }

        [Fact]
        public void When_FileIsNotTemplate_ThrowException()
        { 
            var exception = Assert.Throws<ArgumentException>(
                () => new WordDocument(NonTemplateFilePath));
            
            exception.Message
                .Should()
                .Be("Provided file is not a word template. (Parameter './resource/sample.txt')");
        }

        [Fact]
        public void When_FieldIsSet_ExportedDocumentHasExpectedValue()
        {
            var value = "test"; 
            var document = new WordDocument(TemplateFilePath);
             document.SetField("TestField", value);
            //initialValue.Should().Be("InitialValue");
        }
    }
}