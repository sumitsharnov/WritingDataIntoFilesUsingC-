using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace Assignment_Information_Encoding_4
{
    class WordDocument
    {
        public void  WritingInWordDocument()
        {
            using (DocumentFormat.OpenXml.Packaging.WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(@"C:\\Users\\Public\\Music\\info.doc", WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Hello, my name is Sumit."));
            }
        }
    }
}
