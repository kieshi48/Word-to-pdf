using System.IO;
//using System.Reflection.Metadata;
using Microsoft.Office.Interop.Word;
//using static System.Net.Mime.MediaTypeNames;

string folderPath = @"C:\Users\kiesh\Downloads\картинки апк\Калинина — копия";
string[] files = Directory.GetFiles(folderPath, "*.doc*");

Application word = new Application();

foreach (string file in files)
{
    Document doc = word.Documents.Open(file);

    string pdfName = Path.ChangeExtension(file, ".pdf");
    Console.WriteLine(pdfName);
    doc.ExportAsFixedFormat(pdfName, WdExportFormat.wdExportFormatPDF);

    doc.Close(false);
}

word.Quit();

System.Console.WriteLine("All Word documents in the folder have been converted to PDF.");
System.Console.ReadLine();