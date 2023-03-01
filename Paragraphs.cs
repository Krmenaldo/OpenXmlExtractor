namespace ParagraphExtractor
{
    public class ExtractedParagraph
    {
        public string Heading1 { get; set; }
        public string Heading2 { get; set; }
        public string Heading3 { get; set; }
        public string Paragraph { get; set; }

        public ExtractedParagraph(string heading1, string heading2, string heading3, string paragraph)
        {
            Heading1 = heading1;
            Heading2 = heading2;
            Heading3 = heading3;
            Paragraph = paragraph;
        }
        public override string ToString()
        {
            return Heading1 + "|" + Heading2 + "|" + Heading3 + "|" + Paragraph;
        }
    }

    public class ExtractedDocument
    {
        public List<ExtractedParagraph> Paragraphs { get; set; }
    }
}
