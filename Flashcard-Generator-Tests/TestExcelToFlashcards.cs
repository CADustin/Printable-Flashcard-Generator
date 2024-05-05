// <copyright file="TestExcelToFlashcards.cs" company="Dustin Ellis">
// Copyright (c) Dustin Ellis. All rights reserved.
// </copyright>

namespace FlashcardGeneratorTests
{
    using Flashcard_Generator;

    [TestClass]
    public class TestExcelToFlashcards
    {
        [TestMethod]
        [DataRow("Question", "Answer")]
        [DeploymentItem(@"Data\ExcelDataset1.xlsx")]
        public void TestMethod1(string questionColumnName, string answerColumnName)
        {
            var inputFile = new FileInfo("ExcelDataset1.xlsx");
            var generator = new ExcelToFlashcards(inputFile, questionColumnName, answerColumnName);
        }
    }
}