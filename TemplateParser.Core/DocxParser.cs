// Import required packages.
using System.Text.Json;
using System.Text.RegularExpressions;
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
        // 4) [Week 4] Add formatting heuristics for files missing heading styles.
        // 5) [Week 2-4] Create Node instances with:
        //    - Id: new Guid for each node
        //    - TemplateId: the templateId argument
        //    - ParentId: null for root nodes, set for child nodes
        //    - Type/Title/OrderIndex/MetadataJson based on parsed content
        // 6) [Week 4] Return ParserResult with Nodes in deterministic order.

       
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
            10. Classify remaining text content
            11. Detect tables by checking for Table elements at the body level

            // Week 4:
            12. Add heuristics to infer headings in documents that don't use Word's built-in heading styles

            
        */
            // Result object that will collect all parsed nodes (headings)
            var result = new ParserResult();
            // Counter to generate deterministic GUIDs for nodes
             int idCounter = 1;
            // Tracks ordering of nodes as they appear in the document
            var childOrderMap = new Dictionary<Guid?, int>();
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
                Body? body = wordProcessingDocument.MainDocumentPart?.Document?.Body;
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

                        // Style-based headings
                        if (!string.IsNullOrEmpty(style) && style.StartsWith("Heading"))
                        {
                            int level = ExtractHeadingLevel(style);

                            while (stack.Count > 0 && stack.Peek().Level >= level)
                            {
                                stack.Pop();
                            }

                            Guid? parentId = GetCurrentParentId(stack);

                            if (!childOrderMap.ContainsKey(parentId))
                                childOrderMap[parentId] = 0;

                            int localOrder = childOrderMap[parentId]++;

                            var node = new Node
                            {
                                ParentId = parentId,
                                Id = GenerateDeterministicGuid(idCounter++),
                                TemplateId = templateId,
                                Type = MapHeadingType(level),
                                Title = text,
                                OrderIndex = localOrder,
                                MetadataJson = null
                            };

                            stack.Push((level, node));
                            result.Nodes.Add(node);
                            continue;
                        }

                        // Heuristic-based headings
                        int? inferredLevel = HeuristicHeadingDetector.InferHeadingLevel(p);

                        if (inferredLevel != null)
                        {
                            int level = inferredLevel.Value;


                            while (stack.Count > 0 && stack.Peek().Level >= level)
                            {
                                stack.Pop();
                            }

                            Guid? parentId = GetCurrentParentId(stack);

                            if (!childOrderMap.ContainsKey(parentId))
                                childOrderMap[parentId] = 0;

                            int localOrder = childOrderMap[parentId]++;

                            var node = new Node
                            {
                                ParentId = parentId,
                                Id = GenerateDeterministicGuid(idCounter++),
                                TemplateId = templateId,
                                Type = MapHeadingType(level),
                                Title = text,
                                OrderIndex = localOrder,
                                MetadataJson = null
                            };

                            stack.Push((level, node));
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

                            Guid? parentId = GetCurrentParentId(stack);

                            if (!childOrderMap.ContainsKey(parentId))
                                childOrderMap[parentId] = 0;

                            int localOrder = childOrderMap[parentId]++;

                            var imageNode = new Node
                            {
                                ParentId = parentId,
                                Id = GenerateDeterministicGuid(idCounter++),
                                TemplateId = templateId,
                                Type = "Image",
                                Title = "Image",
                                OrderIndex = localOrder,
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

                            string listType = p?.ParagraphProperties?.NumberingProperties != null
                                ? "numbered"
                                : "bullet";
                            var metadata = JsonSerializer.Serialize(new
                            {
                                listType,
                                items = listItems
                            });

                            Guid? parentId = GetCurrentParentId(stack);

                            if (!childOrderMap.ContainsKey(parentId))
                                childOrderMap[parentId] = 0;

                            int localOrder = childOrderMap[parentId]++;

                            var listNode = new Node
                            {
                                ParentId = parentId,
                                Id = GenerateDeterministicGuid(idCounter++),
                                TemplateId = templateId,
                                Type = "List",
                                Title = "List",
                                OrderIndex = localOrder,
                                MetadataJson = metadata
                            };

                            result.Nodes.Add(listNode);
                            continue;
                        }

                        // 10. Classify remaining text
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            string type = "Text";

                            var metadata = JsonSerializer.Serialize(new
                            {
                                defaultText = text
                            });

                            Guid? parentId = GetCurrentParentId(stack);

                            if (!childOrderMap.ContainsKey(parentId))
                                childOrderMap[parentId] = 0;

                            int localOrder = childOrderMap[parentId]++;

                            var textNode = new Node
                            {
                                ParentId = parentId,
                                Id = GenerateDeterministicGuid(idCounter++),
                                TemplateId = templateId,
                                Type = type,
                                Title = text,
                                OrderIndex = localOrder,
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

                        Guid? parentId = GetCurrentParentId(stack);

                        if (!childOrderMap.ContainsKey(parentId))
                            childOrderMap[parentId] = 0;

                        int localOrder = childOrderMap[parentId]++;

                        var tableNode = new Node
                        {   ParentId = parentId,
                            Id = GenerateDeterministicGuid(idCounter++),
                            TemplateId = templateId,
                            Type = "Table",
                            Title = "Table",
                            OrderIndex = localOrder,
                            MetadataJson = metadata
                        };

                        result.Nodes.Add(tableNode);
                    }
            }
            return result;
        }
    }
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

    private static string MapHeadingType(int level)
    {
        return level switch
        {
            1 => "section",
            2 => "subsection",
            _ => "subsection"
        };
    }
     private static Guid? GetCurrentParentId(
        Stack<(int Level, Node Node)> headingStack)
    {
        // Non-heading content belongs under the most recent heading
        return headingStack.Count > 0
            ? headingStack.Peek().Node.Id
            : null;
    }
    private static Guid GenerateDeterministicGuid(int i)
    {
        return Guid.Parse($"00000000-0000-4000-8000-{i.ToString().PadLeft(12, '0')}");
    }
}
//12. Add heuristics to infer headings in documents that don't use Word's built-in heading styles
public static class HeuristicHeadingDetector
{
    public static int? InferHeadingLevel(Paragraph p)
    {
        string text = p.InnerText?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(text))
            return null;
        
        // Filters (prevent false positives)
        if (text.Length > 120) return null;
        if (text.EndsWith(".")) return null;
        if (text.Count(c => c == '.') > 2) return null;

        // Numbering prefix
        var match = Regex.Match(text, @"^\(?\d+([.\)]\d+){0,2}\)?");
        if (match.Success)
        {
            int dots = match.Value.Count(c => c == '.');
            return dots + 1; // 1 -> level 1, 1.1 -> level 2, etc.
        }
        int score = 0;

        // Font size
        score += GetFontSizeScore(p);

        // Bold
        if (p.Descendants<Bold>().Any())
            score++;

        // Spacing
        if (HasSpacing(p))
            score++;

        // Final decision based on cumulative score
        if (score >= 4) return 1; // Section
        if (score >= 3) return 2; // Subsection
        if (score >= 2) return 3; // Sub-subsection

        return null; // Not a heading        
    }

        private static bool HasSpacing(Paragraph p)
        {
            var spacing = p.ParagraphProperties?.SpacingBetweenLines;
            return spacing?.Before != null || spacing?.After != null;
        }

        private static int GetFontSizeScore(Paragraph p)
        {
            var run = p.Descendants<Run>().FirstOrDefault();
            var size = run?.RunProperties?.FontSize?.Val;

            if (size == null) return 0;

            if (int.TryParse(size.Value, out int sz))
            {
                if (sz >= 32) return 2;
                if (sz >= 28) return 1;
            }

            return 0;
        }
        
    }