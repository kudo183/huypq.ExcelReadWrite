using System;

namespace ExcelReadWriteTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var result = ExcelReadWrite.ExcelReader.Read(@"C:\test.xlsx");
            Console.Write("done");
        }
    }
}
