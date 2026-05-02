# DOCX Template Parser

## Overview

This project parses `.docx` files and converts them into a structured JSON format composed of hierarchical nodes (sections, subsections, text, lists, tables, and images). The parser is designed to work with both well-structured Word documents (using heading styles) and less structured documents by applying heuristic-based inference.

---

# Parsing Strategy

The parser uses a **single-pass traversal** of the document body via the OpenXML SDK.

### Core Approach

1. Open the `.docx` file using `WordprocessingDocument`
2. Traverse all top-level elements in order (paragraphs, tables)
3. Classify each element into a node type:

   * Heading -> section / subsection
   * Paragraph -> text
   * Numbered paragraph -> list
   * Table -> table
   * Drawing -> image
4. Maintain hierarchy using a **stack**
5. Assign for each node:

   * `Id` (deterministic GUID)
   * `TemplateId`
   * `ParentId` (based on hierarchy)
   * `OrderIndex` (relative to siblings)
   * `MetadataJson`

### Hierarchy Handling

* A stack tracks the current heading structure
* When a new heading is encountered:

  * Pop stack until correct parent level is found
  * Attach node to parent
* Non-heading content is assigned to the most recent heading

---

# Heading Detection Heuristics

The parser uses two complementary methods:

## 1. Style-Based Detection (Primary)

Uses Word’s built-in heading styles:

* `Heading1` -> Section
* `Heading2` -> Subsection

### Example

"Heading1: 1. Executive Summary" -> Section

---

## 2. Heuristic-Based Detection (Fallback)

Used when heading styles are missing.

### Signals Used

#### 1. Numbering Pattern

Detects structured numbering:

* `1.` -> Section
* `1.1` -> Subsection
* `1.1.1` -> Sub-subsection

Example:
"4. Findings & Recommendations"
"4.1 Structural Integrity"

---

#### 2. Text Length Filter

* Text longer than ~120 characters is unlikely to be a heading

---

#### 3. Sentence Filter

* Text ending in a period is treated as body text, not a heading

---

#### 4. Font Size

Larger text increases heading likelihood:

* >=32 -> strong signal
* >=28 -> moderate signal

---

#### 5. Bold Formatting

* Bold text increases heading confidence

---

#### 6. Spacing

* Extra spacing before/after paragraph indicates heading

---

### Scoring System

| Score | Classification |
| ----- | -------------- |
| >=4    | Section        |
| >=3    | Subsection     |
| >=2    | Sub-subsection |

---

# How to Run the CLI

### Step 1: Navigate to CLI project

```bash
cd TemplateParser.Cli
```

### Step 2: Run the program

```bash
dotnet run
```

### Step 3: Example usage in Program.cs

```csharp
var parser = new DocxParser();

var result = parser.ParseDocxTemplate(
    "test-documents/test1.docx",
    Guid.Parse("9f7b1b44-2f75-4d52-9b12-3c6d0e6a4b19")
);

var json = JsonSerializer.Serialize(result, new JsonSerializerOptions
{
    WriteIndented = true
});

File.WriteAllText("output.json", json);
```

---

### Output Location

The generated file will appear in:

```
TemplateParser.Cli/bin/Debug/netX.X/output.json
```

---

### Example Output

```json
{
  "nodes": [
    {
      "id": "00000000-0000-4000-8000-000000000001",
      "parentId": null,
      "type": "Section",
      "title": "1. Executive Summary",
      "orderIndex": 0,
      "metadataJson": "{}"
    }
  ]
}
```

---

# Integration Instructions

To use the parser in another project:

### 1. Add reference to Core project

Reference:

```
TemplateParser.Core
```

---

### 2. Call the parser

```csharp
using TemplateParser.Core;

var parser = new DocxParser();

ParserResult result = parser.ParseDocxTemplate(
    "path/to/file.docx",
    Guid.NewGuid()
);
```

---

### 3. Serialize result

```csharp
var json = JsonSerializer.Serialize(result, new JsonSerializerOptions
{
    WriteIndented = true
});
```

---

# Known Limitations

### 1. Inconsistent Heading Styles

* Documents without proper styles rely on heuristics
* May misclassify:

  * short bold text
  * numbered lists

---

### 2. List Detection

* Limited ability to distinguish bullet vs numbered lists
* May default to numbered classification

---

### 3. Image Handling

* Detects images but does not extract:

  * actual image files
  * captions
  * alt text

---

### 4. Table Parsing

* Tables are flattened into text arrays
* Does not preserve:

  * merged cells
  * formatting
  * header semantics

---

### 5. Metadata Structure

* Stored as serialized JSON string
* Not strongly typed
* May differ from external schemas

---

### 6. Heuristic Errors

* Possible false positives:

  * bold short lines
* Possible false negatives:

  * subtle headings without formatting

---

### 7. No Semantic Understanding

* Parser relies only on structure and formatting
* Does not interpret meaning of content

---

# Summary

This parser:

* Extracts structured data from `.docx` files
* Builds hierarchical node trees
* Uses both style-based and heuristic heading detection
* Supports text, lists, tables, and images
* Produces deterministic, testable JSON output

---