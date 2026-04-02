// Import required packages.
using DocumentFormat.OpenXml.Packaging;         // Needed for WordProcessingDocument.
using DocumentFormat.OpenXml.Wordprocessing;    // Needed for all Word schema objects (Body, Paragraph, etc.)
namespace TemplateParser.Core;

public sealed class DocxParser
{
    public ParserResult ParseDocxTemplate(string filePath, Guid templateId)
    {
        // TODO (Week 1-4): Implement core DOCX parsing here.
        // Recommended responsibilities for this method:

        // 1) [Week 1] Learn DOCX structure and print paragraphs from the document.
        // 2) [Week 2] Build section hierarchy using Word heading styles.

       
        // Objective: Create a ReadAllParagraphs() method
        // Objective: Enumerate all paragraph styles and text
        // Objective: Create a Node for each heading paragraph with correct hierarchy based on heading levels
        /* Steps:

            // Week 1:
            1. Open the word document in read mode.
            2. Parse the document.xml into XML object using the DocumentFormat.OpenXml library.
            3. Loop through every paragraph.
            4. Extract and display the paragraph style.
            5. Extract and display the actual text.

            // Week 2:
            6. Create a new node representing a heading
            7. Fix hierarchy using stack
        */
            // Result object that will collect all parsed nodes (headings)
            var result = new ParserResult();
            // Tracks ordering of nodes as they appear in the document
            int orderIndex = 0;
            // Stack used to maintain hierarchy of headings
            // Each entry stores the heading level and its corresponding node
            var stack = new Stack<(int Level, Node Node)>();

            // 1. Open the word document in read mode.
            // 2. Parse the document.xml into XML object using the DocumentFormat.OpenXml library.
            using (WordprocessingDocument wordProcessingDocument = WordprocessingDocument.Open(filePath, false))
            {
                // The original line we wrote in class:
                // Body body = wordProcessingDocument.MainDocumentPart.Document.Body;
                
                // A more robust version that fails gracefully if the document is not structured properly:
                Body? body = wordProcessingDocument?.MainDocumentPart?.Document?.Body;
                ArgumentNullException.ThrowIfNull(body, "Document is empty.");
                
                // 3. Loop through every paragraph.
                foreach (Paragraph p in body.Descendants<Paragraph>())
                {
                    // 4. Extract and display the paragraph style.
                    // The original line we wrote in class:
                    // string style = p?.ParagraphProperties?.ParagraphStyleId?.Val;
                    // A more robust version:
                    string? style = p?.ParagraphProperties?.ParagraphStyleId?.Val ?? "No Style";

                    // Only process headings
                    if (style == null || !style.StartsWith("Heading"))
                        continue;
                     // Convert "Heading1" -> 1, "Heading2" -> 2, etc.
                    int level = ExtractHeadingLevel(style);

                    // 5. Extract and display the actual text.
                    string title = p?.InnerText ?? string.Empty;

                    // 6. Create a new node representing a heading
                    var node = new Node
                    {
                        Id = Guid.NewGuid(),
                        TemplateId = templateId,
                        Type = style,
                        Title = title,
                        OrderIndex = orderIndex++,
                        MetadataJson = "{}"
                    };

                    // 7. Fix hierarchy using stack
                    while (stack.Count > 0 && stack.Peek().Level >= level)
                    {
                        stack.Pop();
                    }
                     // Assign parent if exists (top of stack is the parent)
                    node.ParentId = stack.Count > 0 ? stack.Peek().
                    Node.Id : null;
                    // Push current node onto stack for future children
                    stack.Push((level, node));
                    // Add node to result list
                    result.Nodes.Add(node);
                    
                }
                
            }
        return result;
        
        // 3) [Week 3] Detect tables, lists, and images as structured content nodes.
        // 4) [Week 4] Add formatting heuristics for files missing heading styles.
        // 5) [Week 2-4] Create Node instances with:
        //    - Id: new Guid for each node
        //    - TemplateId: the templateId argument
        //    - ParentId: null for root nodes, set for child nodes
        //    - Type/Title/OrderIndex/MetadataJson based on parsed content
        // 6) [Week 4] Return ParserResult with Nodes in deterministic order.
        //
        // Helper guidance [Week 3-6]:
        // - YES, create helper classes if this method gets long or hard to read.
        // - Keep helpers inside TemplateParser.Core (for example, Parsing/ or Utilities/ folders).
        // - Keep this method as the high-level orchestration entry point.
        // - In Week 6, refactor large blocks from this method into focused helper classes.
        //
        // Do not place parsing logic in the CLI project; keep it in Core.
    }
    private static int ExtractHeadingLevel(string style)
    {
        // Example: "Heading1" -> 1
        if (style.StartsWith("Heading") &&
            int.TryParse(style.Substring(7), out int level))
        {
            return level;
        }

        return int.MaxValue; // fallback (shouldn't happen for valid headings)
    }
}