using System;
using System.IO;

namespace SpreadsheetCalculator
{
    class Program
    {
        static void Main(string[] args)
        {
            var tempPath = Path.GetTempPath();
            var fileName = @"test.xlsm";
            var shortFileName = Path.GetFileNameWithoutExtension(fileName);
            var workbookFolder = Path.Combine(tempPath, shortFileName);
            if (Directory.Exists(workbookFolder))
                Directory.Delete(workbookFolder, true);

            System.IO.Compression.ZipFile.ExtractToDirectory(fileName, workbookFolder);


            var workbookReader = new WorkbookReader();
            var workbook = workbookReader.Read(workbookFolder);

            workbook.GetDefinedNameValues(workbook.DefinedNames[0].Name);
            
            Console.WriteLine("Hello World!");
        }
    }
}