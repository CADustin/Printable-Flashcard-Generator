using Flashcard_Generator;
using Generator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Generator.Tests
{
    [TestClass()]
    public class WriteToPrintableWordDocTests
    {
        [TestMethod()]
        [DeploymentItem(@"Data\ExcelDataset1.xlsx")]
        public void TestWriteToPrintableWordDoc_Write()
        {
            var inputFile = new FileInfo("ExcelDataset1.xlsx");
            var generator = new ReadFromExcel("Question", "Answer");
            generator.ReadDataFromExcel(inputFile);

            FileInfo tempWordFile = new FileInfo(Path.GetTempFileName());
            WriteToPrintableWordDoc.Write(generator.Flashcards, tempWordFile);

            // Our unit tests shouldn't leave behind temp files on the system.
            tempWordFile.Delete();
        }
    }
}