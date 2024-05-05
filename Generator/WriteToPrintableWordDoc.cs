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

                Table table = new();

                TableProperties props = new(
                    new TableBorders(
                    new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    }));

                table.AppendChild<TableProperties>(props);

                for (var i = 0; i < flashcards.Count; i++)
                {
                    var tr = new TableRow();
                    var tcQuestion = new TableCell();
                    tcQuestion.Append(new Paragraph(new Run(new Text(flashcards[i].Question))));
                    tcQuestion.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50", }));

                    tr.Append(tcQuestion);

                    var tcAnswer = new TableCell();
                    tcAnswer.Append(new Paragraph(new Run(new Text(flashcards[i].Answer))));
                    tcAnswer.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "50", }));
                    tr.Append(tcAnswer);

                    table.Append(tr);
                }

                mainPart.Document.Body.Append(table);

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

                mainPart.Document.Save();
            }
        }
    }
}