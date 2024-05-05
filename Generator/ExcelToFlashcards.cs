// <copyright file="ExcelToFlashcards.cs" company="Dustin Ellis">
// Copyright (c) Dustin Ellis. All rights reserved.
// </copyright>

namespace Flashcard_Generator
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    /// <summary>
    /// Excel to Flashcards class.
    /// </summary>
    public class ExcelToFlashcards
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelToFlashcards"/> class.
        /// </summary>
        public ExcelToFlashcards()
        {
            this.Flashcards = new List<Flashcard>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelToFlashcards"/> class.
        /// </summary>
        /// <param name="questionColumnName">The question column name.</param>
        /// <param name="answerColumnName">The answer column name.</param>
        public ExcelToFlashcards(string questionColumnName, string answerColumnName)
            : this()
        {
            this.QuestionColumnName = questionColumnName;
            this.AnswerColumnName = answerColumnName;
        }

        /// <summary>
        /// Read the input file from Excel.
        /// </summary>
        /// <param name="inputFile">The input Excel file to read.</param>
        /// <exception cref="ArgumentNullException">Throws an exception if <paramref name="inputFile"/> is null.</exception>
        public void ReadDataFromExcel(FileInfo inputFile)
        {
            if (inputFile is null)
            {
                throw new ArgumentNullException(nameof(inputFile));
            }

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(inputFile.FullName, false))
            {

                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? spreadsheetDocument.AddWorkbookPart();
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;

                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        text = c?.CellValue?.Text;
                        Console.Write(text + " ");
                    }
                }
            }
        }

        /// <summary>
        /// Gets the name of the Question Column.
        /// </summary>
        public string QuestionColumnName { get; }

        /// <summary>
        /// Gets the name of the Answer column.
        /// </summary>
        public string AnswerColumnName { get; }

        /// <summary>
        /// Gets a collection of <see cref="Flashcard"/> objects.
        /// </summary>
        internal List<Flashcard> Flashcards { get; }
    }
}
