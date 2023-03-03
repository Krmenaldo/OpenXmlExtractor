// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ParagraphExtractor;

Console.WriteLine("Hello, World!");
string[] files =  {
    "Attendance Policy_2023",
    "Compensatory Off Policy_Jan2023",
    "Emergency Financial Assistance Policy",
    "Employee Referral Policy_2023",
    "POSH Policy 2023",
    "SL- Leave Policy_2023",
    "SL Relocation Policy_2023",
    "Technical Certification Policy_2023",
    "Whistle Blower Policy_2023"
};
var cnt = 0;
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
        foreach (var item in doc.MainDocumentPart.Document.Body.ChildElements)
        {
            if (item is not Paragraph && item is not Table)
                continue;
            if (string.IsNullOrEmpty(item.InnerText))
                continue;
            if (item is Table)
            {
                var table = (Table)item;
                var lastRow = table.ChildElements.OfType<TableRow>().LastOrDefault();
                var numOfColumns = 0;
                if (lastRow != null)
                    numOfColumns = lastRow.ChildElements.OfType<TableCell>().Count();

                if (numOfColumns > 0)
                {
                    var tableHeader = table.ChildElements.OfType<TableRow>().FirstOrDefault(t => t.ChildElements.OfType<TableCell>().Count() == numOfColumns);
                    if (tableHeader != null)
                    {
                        var tableText = string.Empty;
                        foreach (var row in table.ChildElements.OfType<TableRow>().Where(t => t.ChildElements.OfType<TableCell>().Count() == numOfColumns).Skip(1))
                        {
                            var counter = 0;
                            var rowName = string.Empty;
                            foreach (var cell in row.ChildElements.OfType<TableCell>())
                            {
                                if (counter == 0)
                                {
                                    rowName = cell.InnerText;
                                }
                                else
                                    tableText += "[" + rowName + ":" + tableHeader.ChildElements.OfType<TableCell>().ElementAt(counter)?.InnerText + "] " + cell.InnerText + " ";
                                counter++;
                            }
                        }
                        document.Paragraphs.Add(new ExtractedParagraph(heading1, heading2, FillAdditionalHeading(heading3, heading4, heading5, heading6), tableText));
                    }
                }

            }
            if (item is Paragraph)
            {
                var paragraph = (Paragraph)item;
                //var images = GetImagesFromParagraph(paragraph, doc);
                //if(images != null && images.Any())
                //{
                //    foreach (var image in images)
                //    {
                //        using (var fileStream = File.Create(Guid.NewGuid() + ExtensionFromContentType(image.ContentType)))
                //        {
                //            image.GetStream().CopyTo(fileStream);
                //        }
                //    }
                //}
                var style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if ((!string.IsNullOrEmpty(style) && style.Contains("Heading")) || paragraph.ParagraphProperties?.NumberingProperties != null)
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
                    else if (paragraph.ParagraphProperties?.NumberingProperties != null)
                    {
                        var last = document.Paragraphs.LastOrDefault();
                        var addNewLine = true;
                        if (last == null || last.Heading1 != heading1 || last.Heading2 != heading2 || last.Heading3 != heading3)
                        {
                            document.Paragraphs.Add(new ExtractedParagraph(heading1, heading2, FillAdditionalHeading(heading3, heading4, heading5, heading6), string.Empty));
                            last = document.Paragraphs.LastOrDefault();
                            addNewLine = false;
                        }
                        if (paragraph.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val.Value == 0)
                        {
                            level0counter++;
                            level1counter = 0;
                            level2counter = 0;
                            last.Paragraph += (addNewLine ? " \\n " : string.Empty) + "[" + level0counter + "] " + paragraph.InnerText;
                        }
                        if (paragraph.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val.Value == 1)
                        {
                            level1counter++;
                            level2counter = 0;
                            last.Paragraph += (addNewLine ? " \\n " : string.Empty) + "[" + level0counter + "." + level1counter + "] " + paragraph.InnerText;
                        }
                        if (paragraph.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val.Value == 2)
                        {
                            level2counter++;
                            last.Paragraph += (addNewLine ? " \\n " : string.Empty) + "[" + level0counter + "." + level1counter + "." + level2counter + "] " + paragraph.InnerText;
                        }
                    }
                }
                else
                {
                    document.Paragraphs.Add(new ExtractedParagraph(heading1, heading2, FillAdditionalHeading(heading3, heading4, heading5, heading6), paragraph.InnerText));
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