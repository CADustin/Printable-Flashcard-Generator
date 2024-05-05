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
    public class ReadFromExcel
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ReadFromExcel"/> class.
        /// </summary>
        /// <param name="questionColumnName">The question column name.</param>
        /// <param name="answerColumnName">The answer column name.</param>
        public ReadFromExcel(string questionColumnName, string answerColumnName)
            : this()
        {
            this.QuestionColumnName = questionColumnName;
            this.AnswerColumnName = answerColumnName;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ReadFromExcel"/> class.
        /// </summary>
        private ReadFromExcel()
        {
            this.Flashcards = new List<Flashcard>();

            this.QuestionColumnName = string.Empty;
            this.AnswerColumnName = string.Empty;
        }

        /// <summary>
        /// Gets the name of the Answer column.
        /// </summary>
        public string AnswerColumnName { get; }

        /// <summary>
        /// Gets a collection of <see cref="Flashcard"/> objects.
        /// </summary>
        public List<Flashcard> Flashcards { get; private set; }

        /// <summary>
        /// Gets the name of the Question Column.
        /// </summary>
        public string QuestionColumnName { get; }

        /// <summary>
        /// Read the input file from Excel.
        /// </summary>
        /// <param name="inputFile">The input Excel file to read.</param>
        /// <exception cref="ArgumentNullException">Throws an exception if <paramref name="inputFile"/> is null.</exception>
        public void ReadDataFromExcel(FileInfo inputFile)
        {
            ArgumentNullException.ThrowIfNull(inputFile);

            if (!inputFile.Exists)
            {
                throw new FileNotFoundException(nameof(inputFile));
            }

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(inputFile.FullName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // 1. Establish expectations about the required headers
                Dictionary<string, int> expectedHeaders = new Dictionary<string, int>()
                {
                    { this.QuestionColumnName, -1 },
                    { this.AnswerColumnName, -1 },
                };

                // 2. Find the headers
                var headerRow = sheetData.Elements<Row>().First();
                if (headerRow.HasAttributes)
                {
                    var cellDataInThisRow = GetDataFromRow(workbookPart, headerRow);

                    for (int i = 0; i < cellDataInThisRow.Count; i++)
                    {
                        string data = cellDataInThisRow[i];
                        if (expectedHeaders.ContainsKey(data))
                        {
                            expectedHeaders[data] = i;
                        }
                    }
                }

                // 3. Generate the Flashcards
                if (expectedHeaders.Values.All(x => x >= 0))
                {
                    var allRowsExceptHeader = sheetData.Elements<Row>().Skip(1).Where(x => x.HasAttributes);
                    foreach (var row in allRowsExceptHeader)
                    {
                        // It might be more efficient to only return the columns of data we need -- but then again, that might add complexity thats not needed since we can control the input format fairly well.
                        var dataFromRow = GetDataFromRow(workbookPart, row);

                        var answer = dataFromRow[expectedHeaders[this.AnswerColumnName]];
                        Flashcard card = new Flashcard()
                        {
                            Answer = dataFromRow[expectedHeaders[this.AnswerColumnName]],
                            Question = dataFromRow[expectedHeaders[this.QuestionColumnName]],
                        };

                        this.Flashcards.Add(card);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the data from a row.
        /// </summary>
        /// <param name="workbookPart">The Workbook.</param>
        /// <param name="row">The Row of interest.</param>
        /// <returns>Collection of strings containing the data from the row.</returns>
        private static List<string> GetDataFromRow(WorkbookPart workbookPart, Row row)
        {
            List<string> rowStrings = new List<string>();
            if (row != null && workbookPart?.SharedStringTablePart != null)
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                    {
                        if (int.TryParse(cell.InnerText, out int id))
                        {
                            SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);

                            // At first I thought I should check to see if the string was empty, null, or whitespace ... but its possible the cell is an empty string, and we would want to bring that back.
                            string cellValue = ssi.Text != null ? ssi.Text.Text : ssi.InnerText;

                            // We dont need whitespace, so we might as well trim any as early as possible.
                            rowStrings.Add(cellValue.Trim());
                        }
                    }
                }
            }

            return rowStrings;
        }
    }
}