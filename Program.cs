// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ParagraphExtractor;

Console.WriteLine("Hello, World!");
string[] files =  {
    //"Bezbednost i zdravlje na radu u BIB",
    "Digitalno bankarstvo za FL 6"
};
var line = string.Empty;
foreach (var file in files)
{
    using (var doc = WordprocessingDocument.Open("documents\\" + file + ".docx", false))
    {
        var document = new ExtractedDocument
        {
            Paragraphs = new List<ExtractedParagraph>()
        };
        var heading1 = string.Empty;
        var heading2 = string.Empty;
        var heading3 = string.Empty;
        var heading4 = string.Empty;
        var heading5 = string.Empty;
        var heading6 = string.Empty;
        var level0counter = 0;
        var level1counter = 0;
        var level2counter = 0;
        var cnt = 1;
        string lastText = string.Empty;
        foreach (var item in doc.MainDocumentPart.Document.Body.ChildElements)
        {
            if (item is not Paragraph && item is not Table)
                continue;
            if (string.IsNullOrEmpty(item.InnerText))
                continue;
            if (item is Table)
            {
                var table = (Table)item;
                var tableText = string.Empty;
                var counter = 0;
                foreach (var row in table.ChildElements.OfType<TableRow>())
                {
                    counter++;
                    tableText += "[" + counter + "] - ";
                    var counter2 = 0;
                    foreach (var cell in row.ChildElements.OfType<TableCell>())
                    {
                        counter2++;
                        tableText += "(" + counter2 + ") " + cell.InnerText;
                    }
                    tableText += " \\n ";
                }
                document.Paragraphs.Add(new ExtractedParagraph(heading1, heading2, FillAdditionalHeading(heading3, heading4, heading5, heading6), tableText));
            }
            if (item is Paragraph)
            {
                var paragraph = (Paragraph)item;
                if (paragraph.InnerText.Contains("PAGEREF _Toc") || paragraph.InnerText.Contains("TOC \\o"))
                    continue;
                var images = GetImagesFromParagraph(paragraph, doc);
                if (images != null && images.Any())
                {
                    foreach (var image in images)
                    {
                        using (var fileStream = File.Create(Guid.NewGuid() + ExtensionFromContentType(image.ContentType)))
                        {
                            image.GetStream().CopyTo(fileStream);
                        }
                    }
                }
                var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if ((!string.IsNullOrEmpty(style) && (style.Contains("Heading") || style.Contains("List"))) || paragraph.ParagraphProperties?.NumberingProperties != null)
                {
                    if (!string.IsNullOrEmpty(style) && style.Contains("Heading1"))
                    {
                        heading1 = paragraph.InnerText;
                        heading2 = string.Empty;
                        heading3 = string.Empty;
                        heading4 = string.Empty;
                        heading5 = string.Empty;
                        heading6 = string.Empty;
                        level0counter = 0;
                    }
                    else if (!string.IsNullOrEmpty(style) && style.Contains("Heading2"))
                    {
                        heading2 = paragraph.InnerText;
                        heading3 = string.Empty;
                        heading4 = string.Empty;
                        heading5 = string.Empty;
                        heading6 = string.Empty;
                        level0counter = 0;
                    }
                    else if (!string.IsNullOrEmpty(style) && style.Contains("Heading3"))
                    {
                        heading3 = paragraph.InnerText;
                        heading4 = string.Empty;
                        heading5 = string.Empty;
                        heading6 = string.Empty;
                        level0counter = 0;
                    }
                    else if (!string.IsNullOrEmpty(style) && style.Contains("Heading"))
                    {
                        if (style.Contains("Heading4"))
                        {
                            heading4 = paragraph.InnerText;
                            heading5 = string.Empty;
                            heading6 = string.Empty;
                        }
                        if (style.Contains("Heading5"))
                        {
                            heading5 = paragraph.InnerText;
                            heading6 = string.Empty;
                        }
                        if (style.Contains("Heading6"))
                            heading6 = paragraph.InnerText;
                        level0counter = 0;
                    }
                    else if (paragraph.ParagraphProperties?.NumberingProperties != null || (!string.IsNullOrEmpty(style) && style.Contains("List")))
                    {
                        var last = document.Paragraphs.LastOrDefault();
                        if (last == null || last.Heading1 != heading1 || last.Heading2 != heading2 || last.Heading3 != heading3
                            || (last != null && last.ParagraphType == ParagraphType.List && ShouldCreateNew(last.Paragraph, paragraph.InnerText))
                            || (last != null && last.ParagraphType == ParagraphType.Text && last.Paragraph != lastText))
                        {
                            document.Paragraphs.Add(new ExtractedParagraph(heading1, heading2, FillAdditionalHeading(heading3, heading4, heading5, heading6), lastText, ParagraphType.List));
                            last = document.Paragraphs.LastOrDefault();
                        }
                        last.ParagraphType = ParagraphType.List;
                        if ((paragraph.ParagraphProperties?.NumberingProperties != null && paragraph.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val.Value == 0) 
                            || (!string.IsNullOrEmpty(style) && style.Contains("ListBullet2")))
                        {
                            level0counter++;
                            level1counter = 0;
                            level2counter = 0;
                            last.Paragraph += " \\n [" + level0counter + "] " + paragraph.InnerText;
                        }
                        else if ((paragraph.ParagraphProperties?.NumberingProperties != null && paragraph.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val.Value == 1)
                            || (!string.IsNullOrEmpty(style) && style.Contains("ListBullet3")))
                        {
                            level1counter++;
                            level2counter = 0;
                            last.Paragraph += " \\n [" + level0counter + "." + level1counter + "] " + paragraph.InnerText;
                        }
                        else if ((paragraph.ParagraphProperties?.NumberingProperties != null && paragraph.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val.Value == 2)
                            || (!string.IsNullOrEmpty(style) && style.Contains("ListBullet4")))
                        {
                            level2counter++;
                            last.Paragraph += " \\n [" + level0counter + "." + level1counter + "." + level2counter + "] " + paragraph.InnerText;
                        }
                        else if (!string.IsNullOrEmpty(style) && style.Contains("List"))
                        {
                            level0counter++;
                            level1counter = 0;
                            level2counter = 0;
                            last.Paragraph += " \\n [" + level0counter + "] " + paragraph.InnerText;
                        }
                    }
                }
                else
                {
                    var last = document.Paragraphs.LastOrDefault();
                    lastText = paragraph.InnerText;
                    if (last == null || last.Heading1 != heading1 || last.Heading2 != heading2 || last.Heading3 != heading3 || last.ParagraphType != ParagraphType.Text)
                        document.Paragraphs.Add(new ExtractedParagraph(heading1, heading2, FillAdditionalHeading(heading3, heading4, heading5, heading6), paragraph.InnerText));
                    else
                    {
                        if (ShouldCreateNew(last.Paragraph, paragraph.InnerText))
                            document.Paragraphs.Add(new ExtractedParagraph(heading1, heading2, FillAdditionalHeading(heading3, heading4, heading5, heading6), paragraph.InnerText));
                        else
                            last.Paragraph += " \\n " + paragraph.InnerText;
                    }
                    level0counter = 0;
                }
            }
        }
        foreach(var paragraph in document.Paragraphs)
        {
            line += cnt.ToString() + "\t" + file + "\t" + paragraph.ToString();
            line += Environment.NewLine;
            cnt++;
        }
    }
}
File.WriteAllTextAsync("all.csv", line);
Console.ReadLine();

bool ShouldCreateNew(string currentParagraph, string newParagraph)
{
    var testParagraph = currentParagraph + " \\n " + newParagraph;
    return (CountTokens(testParagraph) > 200 && CountTokens(newParagraph) > 100) || CountTokens(testParagraph) > 250;
}

string FillAdditionalHeading(string heading3, string heading4, string heading5, string heading6)
{
    if (!string.IsNullOrEmpty(heading4))
        heading3 += ";" + heading4;
    if (!string.IsNullOrEmpty(heading5))
        heading3 += ";" + heading5;
    if (!string.IsNullOrEmpty(heading6))
        heading3 += ";" + heading6;
    return heading3;
}

IEnumerable<ImagePart> GetImagesFromParagraph(Paragraph paragraph, WordprocessingDocument doc)
{
    var images = from graphic in paragraph
                    .Descendants<DocumentFormat.OpenXml.Drawing.Graphic>()
                 let graphicData = graphic.Descendants<DocumentFormat.OpenXml.Drawing.GraphicData>().FirstOrDefault()
                 let pic = graphicData.ElementAt(0)
                 let nvPicPrt = pic.ElementAt(0).FirstOrDefault()
                 let blip = pic.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault()
                 join part in doc.MainDocumentPart.Parts on blip.Embed.Value equals part
                     .RelationshipId
                 let image = part.OpenXmlPart as ImagePart
                 select image;
    return images;
}

string ExtensionFromContentType(string contentType)
{
    switch(contentType)
    {
        case "image/png":
            return ".png";
        case "image/jpeg":
        case "image/jpg":
            return ".jpg";
    }
    return string.Empty;
}

decimal CountTokens(string text)
{
    int wordCount = 0, index = 0;
    while (index < text.Length && char.IsWhiteSpace(text[index]))
        index++;

    while (index < text.Length)
    {
        while (index < text.Length && !char.IsWhiteSpace(text[index]))
            index++;
        wordCount++;
        while (index < text.Length && char.IsWhiteSpace(text[index]))
            index++;
    }

    return (decimal)(wordCount * 1.25);
}