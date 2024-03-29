﻿namespace ParagraphExtractor
{
    public class ExtractedParagraph
    {
        public string Heading1 { get; set; }
        public string Heading2 { get; set; }
        public string Heading3 { get; set; }
        public string Paragraph { get; set; }
        public ParagraphType ParagraphType { get; set; }

        public ExtractedParagraph(string heading1, string heading2, string heading3, string paragraph)
        {
            Heading1 = heading1;
            Heading2 = heading2;
            Heading3 = heading3;
            Paragraph = paragraph;
            ParagraphType = ParagraphType.Text;
        }
        public ExtractedParagraph(string heading1, string heading2, string heading3, string paragraph, ParagraphType paragraphType)
        {
            Heading1 = heading1;
            Heading2 = heading2;
            Heading3 = heading3;
            Paragraph = paragraph;
            ParagraphType = paragraphType;
        }
        public override string ToString()
        {
            return Heading1 + "\t" + Heading2 + "\t" + Heading3 + "\t" + Paragraph;
        }
    }

    public class ExtractedDocument
    {
        public List<ExtractedParagraph> Paragraphs { get; set; }
    }

    public enum ParagraphType
    {
        Text,
        List,
        Table
    }
}
