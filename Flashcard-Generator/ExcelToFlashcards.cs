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

    /// <summary>
    /// Excel to Flashcards class.
    /// </summary>
    public class ExcelToFlashcards
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelToFlashcards"/> class.
        /// </summary>
        /// <param name="inputFile">The Input File.</param>
        /// <param name="questionColumnName">The question column name.</param>
        /// <param name="answerColumnName">The answer column name.</param>
        public ExcelToFlashcards(FileInfo inputFile, string questionColumnName, string answerColumnName)
        {
            this.InputFile = inputFile;
            this.QuestionColumnName = questionColumnName;
            this.AnswerColumnName = answerColumnName;
        }

        /// <summary>
        /// Gets the input file that is to be read in.
        /// </summary>
        public FileInfo InputFile { get; }

        /// <summary>
        /// Gets the name of the Question Column.
        /// </summary>
        public string QuestionColumnName { get; }

        /// <summary>
        /// Gets the name of the Answer column.
        /// </summary>
        public string AnswerColumnName { get; }
    }
}
