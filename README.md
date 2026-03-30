# TemplateParser.NET

TemplateParser.NET is a beginner-friendly .NET 10 starter repository for a 6-week practicum where you will build a DOCX template parser from scratch.

This repo is intentionally scaffolded with almost no parser logic so you can practice designing and implementing it yourself.

## What You Should Edit

If you are new, follow this rule:

- **Main implementation work goes in** `TemplateParser.Core`
- **Command-line wiring lives in** `TemplateParser.Cli`
- **Tests go in** `TemplateParser.Tests`

Most students should start in `TemplateParser.Core/DocxParser.cs`.

## Project Structure

- `TemplateParser.Core`
  - Domain models (`Node`, `ParserResult`)
  - `DocxParser` class where parsing behavior should be implemented
- `TemplateParser.Cli`
  - Console entry point (`Program.cs`)
  - Validates input and calls `DocxParser`
- `TemplateParser.Tests`
  - xUnit tests
  - Start by replacing placeholder tests with real parser behavior tests
- `sample-documents`
  - Put your input DOCX files here
- `expected`
  - Put expected output JSON files or fixtures here

## The `Node` Object (Most Important Model)

Your parser should ultimately transform DOCX content into a list of `Node` objects.

Each `Node` represents one meaningful piece of parsed content (for example: section, paragraph group, table, list, or image reference).

Fields on `Node`:

- `Id`: unique identifier for this node
- `TemplateId`: links every node back to the template being parsed
- `ParentId`: supports hierarchy (null for root nodes, set for child nodes)
- `Type`: category of content (for example: `section`, `table`, `list`, `image`)
- `Title`: human-readable label for the node
- `OrderIndex`: position in document order
- `MetadataJson`: extra structured details you want to store

Keep this mental model as you build:

- DOCX input -> parse content blocks -> map blocks to `Node` -> return `ParserResult.Nodes`

## Can You Create Helper Classes?

Yes. You are encouraged to create helpers when code becomes hard to read.

- Good place for parser helpers: inside `TemplateParser.Core` (for example: `TemplateParser.Core/Parsing/` or `TemplateParser.Core/Utilities/`)
- Keep helper names specific (example: `HeadingExtractor`, `NodeFactory`, `DocxTraversalService`)
- Avoid putting parsing logic in `TemplateParser.Cli`; keep parsing logic in `TemplateParser.Core`
- Keep each helper focused on one responsibility

## Suggested 6-Week Path

- **Week 1:** Understand the `.docx` file structure and print paragraphs from the document
- **Week 2:** Build the basic section hierarchy from Word heading styles
- **Week 3:** Detect tables, lists, and images as structured content nodes
- **Week 4:** Implement formatting heuristics when heading styles are missing
- **Week 5:** Wire the parser into a CLI tool and add integration tests
- **Week 6:** Documentation, refactoring, and final delivery

In every week, ask: "How does this step improve the quality of the `Node` objects I produce?"

## Run the CLI

From the repo root:

```bash
dotnet run --project TemplateParser.Cli -- parse <filePath> <templateId>
```

Shorter option from the repo root (recommended):

```bash
./parse <filePath> <templateId>
```

From inside `TemplateParser.Cli`:

```bash
dotnet run -- parse <filePath> <templateId>
```

## Learning Workflow (Recommended)

1. Add a sample DOCX in `sample-documents`
2. Write or update tests in `TemplateParser.Tests`
3. Implement parser behavior in `TemplateParser.Core`
4. Run:
   - `dotnet build`
   - `dotnet test`
5. Use CLI to inspect JSON output manually
6. Save expected outputs in `expected` and compare

## Important Note

Parser implementation is intentionally incomplete in this starter repository. You are expected to implement:

- DOCX reading/parsing
- node extraction
- relationship building
- metadata handling
- test coverage
