using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using SelTestSolo.SeleniumTodos;
using System.Diagnostics;

namespace SelTestSolo
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void ReadingDataFromExcelTestMethod()
        {
            try
            {
                ExcelReader.PopulateInCollection(@"C:\Users\jacke\source\repos\dotnet-sqldb-tutorial-master\SeleniumTests\ExcelSpreadsheets\DataDrivenTesting.xlsx");
                Debug.WriteLine(ExcelReader.ReadData(1,"Name"));
                Debug.WriteLine(ExcelReader.ReadData(2,"Name"));
                Debug.WriteLine(ExcelReader.ReadData(3,"Name"));
                Debug.WriteLine(ExcelReader.ReadData(4,"Name"));
                Debug.WriteLine(ExcelReader.ReadData(5,"Name"));
                Debug.WriteLine(ExcelReader.ReadData(6,"Name"));
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
