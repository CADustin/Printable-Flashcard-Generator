// <copyright file="TestReadFromExcel.cs" company="Dustin Ellis">
// Copyright (c) Dustin Ellis. All rights reserved.
// </copyright>

namespace FlashcardGeneratorTests
{
    using System.Linq;
    using Flashcard_Generator;

    /// <summary>
    /// Test class for the <see cref="ReadFromExcel"/> class.
    /// </summary>
    [TestClass]
    public class TestReadFromExcel
    {
        /// <summary>
        /// Test reading from an Excel file.
        /// </summary>
        [TestMethod]
        [DeploymentItem(@"Data\ExcelDataset1.xlsx")]
        public void TestReadFromExcel_ReadDataFromExcel()
        {
            var inputFile = new FileInfo("ExcelDataset1.xlsx");
            var generator = new ReadFromExcel("Question", "Answer");
            generator.ReadDataFromExcel(inputFile);

            Assert.AreEqual(3, generator.Flashcards.Count, $"{generator.Flashcards}.Count");

            Flashcard second = generator.Flashcards[1];
            Assert.AreEqual("question2", second.Question, nameof(second.Question));
            Assert.AreEqual("answer2", second.Answer, nameof(second.Answer));
        }

        /// <summary>
        /// Test reading from an excel file when our headers wont be found.
        /// </summary>
        [TestMethod]
        [DeploymentItem(@"Data\ExcelDataset1.xlsx")]
        public void TestReadFromExcel_ReadDataFromExcel_InvalidHeaders()
        {
            var inputFile = new FileInfo("ExcelDataset1.xlsx");
            var generator = new ReadFromExcel("qestion headers that wont be found because its long and we cant spell", "column for answers");
            generator.ReadDataFromExcel(inputFile);

            Assert.AreEqual(0, generator.Flashcards.Count, $"{generator.Flashcards}.Count");
        }
    }
}