// <copyright file="WriteToPrintableWordDoc.cs" company="Dustin Ellis">
// Copyright (c) Dustin Ellis. All rights reserved.
// </copyright>

namespace Generator
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml;
    using Flashcard_Generator;

    public class WriteToPrintableWordDoc
    {
        public static void Write(IList<Flashcard> flashcards, FileInfo outputFile)
        {
            // Create a document by supplying the filepath.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputFile.FullName, WordprocessingDocumentType.Document))
            {
                // Add a main document part.
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));

                // Setup the page margins
                PageMargin pageMargin1 = new PageMargin()
                {
                    Top = 720,
                    Right = (UInt32Value)720U,
                    Bottom = 720,
                    Left = (UInt32Value)720U,
                    Header = (UInt32Value)720U,
                    Footer = (UInt32Value)720U,
                    Gutter = (UInt32Value)0U,
                };
                mainPart.Document.Append(pageMargin1);
            }
        }
    }
}