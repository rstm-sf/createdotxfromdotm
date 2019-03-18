using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CreateDocxFromDotx
{
    class Program
    {
        static void Main(string[] args)
        {
            var destinationFile = Path.Combine(
                Environment.CurrentDirectory, "Sample\\Doc.docx");
            var sourceFile = Path.Combine(
                Environment.CurrentDirectory, "Sample\\TemplateDoc.dotx");
            try
            {
                File.Copy(sourceFile, destinationFile, true);
                using (var document = WordprocessingDocument.Open(destinationFile, true))
                {
                    document.ChangeDocumentType(WordprocessingDocumentType.Document);

                    var tableMem = Utility.CreateTestDataTable("TestTable");
                    var docTable = document.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault();

                    if (docTable != null)
                    {
                        var rows = docTable.Descendants<TableRow>().ToList();
                        var myRow = (TableRow)rows.Last().Clone();
                        rows.Last().Remove();

                        for (var i = 0; i < tableMem.Rows.Count; ++i)
                        {
                            var rowDoc = (TableRow)myRow.Clone();
                            var cellsDoc = rowDoc.Descendants<TableCell>().ToList();
                            for (var j = 0; j < tableMem.Columns.Count; ++j)
                                cellsDoc[j].Descendants<Text>().FirstOrDefault().Text = tableMem.Rows[i][j].ToString();

                            docTable.Descendants<TableRow>().Last().InsertAfterSelf(rowDoc);
                        }
                    }

                    document.Save();
                    Console.WriteLine("Document generated at " + destinationFile);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("\nPress Enter to continue…");
                Console.ReadLine();
            }
        }
    }
}
