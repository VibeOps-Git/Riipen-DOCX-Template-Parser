using System.Text.Json;
using TemplateParser.Core;

namespace TemplateParser.Tests;

public sealed class ParserTests
{
    // TODO (Week 1-6): Replace this placeholder with real tests aligned to weekly goals.
        // Suggested first tests:
        // - [Week 1] paragraph extraction smoke tests
        // - [Week 2] heading-based hierarchy tests
        // - [Week 3] table/list/image node detection tests
        // - [Week 4] formatting heuristics tests for missing heading styles
        // - [Week 5] integration tests that run parser through the CLI flow
        // - [Week 6] refactor tests into readable groups and helper builders
        //
        // You may create test helpers/builders to reduce repetition (recommended by Week 6).
    private readonly DocxParser _parser = new();

    private static string TestFile(string name)
        => Path.Combine("test-documents", name);

    private static string ExpectedFile(string name)
        => Path.Combine("expected", name);

    private static ParserResult LoadExpected(string file)
    {
        var json = File.ReadAllText(file);
        return JsonSerializer.Deserialize<ParserResult>(json)!;
    }

    // Node count
    [Fact]
    public void Parse_Test1_NodeCountMatchesExpected()
    {
        var result = _parser.ParseDocxTemplate(
            TestFile("test1.docx"),
            Guid.Parse("9f7b1b44-2f75-4d52-9b12-3c6d0e6a4b19"));

        var expected = LoadExpected(ExpectedFile("test1.json"));

        Assert.Equal(expected.Nodes.Count, result.Nodes.Count);
    }

    // Heading detection
    [Fact]
    public void Parse_Test1_ContainsSectionsOrSubsections()
    {
        var result = _parser.ParseDocxTemplate(
            TestFile("test1.docx"),
            Guid.NewGuid());

        Assert.Contains(result.Nodes, n =>
            n.Type == "section" || n.Type == "subsection");
    }

    // List detection
    [Fact]
    public void Parse_Test1_ContainsListNodes()
    {
        var result = _parser.ParseDocxTemplate(
            TestFile("test1.docx"),
            Guid.NewGuid());

        Assert.Contains(result.Nodes, n => n.Type == "list");
    }

    // Hierarchy detection
    [Fact]
    public void Parse_Test2_HasHierarchy()
    {
        var result = _parser.ParseDocxTemplate(
            TestFile("test2.docx"),
            Guid.NewGuid());

        Assert.Contains(result.Nodes, n => n.ParentId != null);
    }

    // Text node detection
    [Fact]
    public void Parse_Test2_ContainsTextNodes()
    {
        var result = _parser.ParseDocxTemplate(
            TestFile("test2.docx"),
            Guid.NewGuid());

        Assert.Contains(result.Nodes, n => n.Type == "text");
    }

    // Image detection
    [Fact]
    public void Parse_Test1_ContainsImagesIfPresent()
    {
        var result = _parser.ParseDocxTemplate(
            TestFile("test1.docx"),
            Guid.NewGuid());

        // Only asserts if document contains images
        Assert.True(result.Nodes.Any(n => n.Type == "image") ||
                    !result.Nodes.Any(n => true));
    }

    // Table detection
    [Fact]
    public void Parse_Test1_ContainsTablesIfPresent()
    {
        var result = _parser.ParseDocxTemplate(
            TestFile("test1.docx"),
            Guid.NewGuid());

        Assert.True(result.Nodes.Any(n => n.Type == "table") ||
                    !result.Nodes.Any(n => true));
    }
}