// Import required packages.
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;         // Needed for WordProcessingDocument.
using DocumentFormat.OpenXml.Wordprocessing;    // Needed for all Word schema objects (Body, Paragraph, etc.)
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
namespace TemplateParser.Core;

public sealed class DocxParser
{
    public ParserResult ParseDocxTemplate(string filePath, Guid templateId)
    {
        // TODO (Week 1-4): Implement core DOCX parsing here.
        // Recommended responsibilities for this method:

        // 1) [Week 1] Learn DOCX structure and print paragraphs from the document.
        // 2) [Week 2] Build section hierarchy using Word heading styles.
        // 3) [Week 3] Detect tables, lists, and images as structured content nodes.

       
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

            // Week 3:
            8. Detect images by looking for Drawing elements within the paragraph
            9. Detect lists by checking for NumberingProperties in paragraph properties
            10. Classify remaining text content as either "Sentence" or "Paragraph" based on heuristics
            11. Detect tables by checking for Table elements at the body level
            
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
                
                // Convert to list so we can manually control index
                // This is required for grouping consecutive list items
                var elements = body.Elements().ToList();

                for (int i = 0; i < elements.Count; i++)
                {
                    OpenXmlElement element = elements[i];
                    if (element is Paragraph p)
                    {

                        // 4. Extract and display the paragraph style.
                        // The original line we wrote in class:
                        // string style = p?.ParagraphProperties?.ParagraphStyleId?.Val;
                        // A more robust version:
                        string? style = p?.ParagraphProperties?.ParagraphStyleId?.Val ?? "No Style";

                        // 5. Extract and display the actual text.
                        string text = p?.InnerText ?? string.Empty;

                        // Only process headings
                        if (null != style && style.StartsWith("Heading"))
                        {
                            // Convert "Heading1" -> 1, "Heading2" -> 2, etc.
                            int level = ExtractHeadingLevel(style);

                            // 6. Create a new node representing a heading
                            var node = new Node
                            {
                                Id = Guid.NewGuid(),
                                TemplateId = templateId,
                                Type = style,
                                Title = text,
                                OrderIndex = orderIndex++,
                                MetadataJson = "{}"
                            };

                            // 7. Fix hierarchy using stack
                            while (stack.Count > 0 && stack.Peek().Level >= level)
                            {
                                stack.Pop();
                            }
                            // Assign parent if exists (top of stack is the parent)
                            node.ParentId = stack.Count > 0 ? stack.Peek().Node.Id : null;
                            // Push current node onto stack for future children
                            stack.Push((level, node));
                            // Add node to result list
                            result.Nodes.Add(node);
                            continue;
                        }
                    
                        //8. Detect images by looking for Drawing elements within the paragraph
                        Drawing? drawing =
                            p?.Descendants<Drawing>().FirstOrDefault();

                        if (drawing != null)
                        {
                            Extent? extent =
                                drawing.Descendants<Extent>().FirstOrDefault();

                            long widthEmu = extent?.Cx ?? 0;
                            long heightEmu = extent?.Cy ?? 0;

                            var metadata = JsonSerializer.Serialize(new
                            {
                                widthEmu,
                                heightEmu
                            });

                            var imageNode = new Node
                            {
                                Id = Guid.NewGuid(),
                                TemplateId = templateId,
                                ParentId = GetCurrentParentId(stack),
                                Type = "Image",
                                Title = "Image",
                                OrderIndex = orderIndex++,
                                MetadataJson = metadata
                            };

                            result.Nodes.Add(imageNode);
                            continue;
                        }
                   
                        // 9. Detect lists by checking for NumberingProperties in paragraph properties
                        if (p?.ParagraphProperties?.NumberingProperties != null)
                        {
                            var listItems = new List<string>();

                            // Continue consuming consecutive list paragraphs
                            while (i < elements.Count &&
                                elements[i] is Paragraph listParagraph &&
                                listParagraph.ParagraphProperties?.NumberingProperties != null)
                            {
                                listItems.Add(listParagraph.InnerText ?? string.Empty);
                                i++;
                            }

                            // Step back one so outer for-loop increments correctly
                            i--;

                            var metadata = JsonSerializer.Serialize(new
                            {
                                items = listItems
                            });

                            var listNode = new Node
                            {
                                Id = Guid.NewGuid(),
                                TemplateId = templateId,
                                ParentId = GetCurrentParentId(stack),
                                Type = "List",
                                Title = "List",
                                OrderIndex = orderIndex++,
                                MetadataJson = metadata
                            };

                            result.Nodes.Add(listNode);
                            continue;
                        }

                        // 10. Classify remaining text content as either "Sentence" or "Paragraph" based on heuristics
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            string type = IsSentence(text)
                                ? "Sentence"
                                : "Paragraph";

                            var metadata = JsonSerializer.Serialize(new
                            {
                                classification = type
                            });

                            var textNode = new Node
                            {
                                Id = Guid.NewGuid(),
                                TemplateId = templateId,
                                ParentId = GetCurrentParentId(stack),
                                Type = type,
                                Title = text,
                                OrderIndex = orderIndex++,
                                MetadataJson = metadata
                            };

                            result.Nodes.Add(textNode);
                        }
                    }

                    // 11. Detect tables by checking for Table elements at the body level  
                    else if (element is Table table)
                    {
                        var tableData = new List<List<string>>();

                        foreach (TableRow row in table.Elements<TableRow>())
                        {
                            var rowData = new List<string>();

                            foreach (TableCell cell in row.Elements<TableCell>())
                            {
                                rowData.Add(cell.InnerText);
                            }

                            tableData.Add(rowData);
                        }

                        int rows = tableData.Count;
                        int columns = rows > 0 ? tableData[0].Count : 0;

                        var metadata = JsonSerializer.Serialize(new
                        {
                            rows,
                            columns,
                            tableData
                        });

                        var tableNode = new Node
                        {
                            Id = Guid.NewGuid(),
                            TemplateId = templateId,
                            ParentId = GetCurrentParentId(stack),
                            Type = "Table",
                            Title = "Table",
                            OrderIndex = orderIndex++,
                            MetadataJson = metadata
                        };

                        result.Nodes.Add(tableNode);
                    }
            }
            return result;
        }
    }
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

     private static Guid? GetCurrentParentId(
        Stack<(int Level, Node Node)> headingStack)
    {
        // Non-heading content belongs under the most recent heading
        return headingStack.Count > 0
            ? headingStack.Peek().Node.Id
            : null;
    }

    private static bool IsListParagraph(Paragraph paragraph)
    {
        // NumberingProperties indicates bulleted/numbered list 
        return paragraph.ParagraphProperties?.NumberingProperties != null;
    }

    private static bool IsSentence(string text)
    {
        // Short text with <= 1 sentence-ending punctuation mark
        // is treated as a sentence instead of full paragraph.
        int punctuationCount =
            text.Count(c => c == '.' || c == '!' || c == '?');

        return punctuationCount <= 1 && text.Length < 120;
    }
}