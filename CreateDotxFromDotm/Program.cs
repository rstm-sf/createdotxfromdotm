using System;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CreateDotxFromDotm
{
    class Program
    {
        static void Main(string[] args)
        {
            string destinationFile = Path.Combine(
                Environment.CurrentDirectory, "Sample\\Doc.docx");
            string sourceFile = Path.Combine(
                Environment.CurrentDirectory, "Sample\\TemplateDoc.dotx");
            try
            {
                File.Copy(sourceFile, destinationFile, true);
                using (var document = WordprocessingDocument.Open(destinationFile, true))
                {
                    document.ChangeDocumentType(WordprocessingDocumentType.Document);

                    var tableMem = CreateTestDataTable("TestTable");
                    var tableDoc = document.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault();

                    var rows = tableDoc.Descendants<TableRow>().ToList();
                    var myRow = (TableRow)rows.Last().Clone();
                    rows.Last().Remove();

                    for (int i = 0; i < tableMem.Rows.Count; ++i)
                    {
                        var rowDoc = (TableRow)myRow.Clone();
                        var cellsDoc = rowDoc.Descendants<TableCell>().ToList();
                        for (int j = 0; j < tableMem.Columns.Count; ++j)
                            cellsDoc[j].Descendants<Text>().FirstOrDefault().Text = tableMem.Rows[i][j].ToString();

                        tableDoc.Descendants<TableRow>().Last().InsertAfterSelf(rowDoc);
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

        private static DataTable CreateTestDataTable(string tableName)
        {
            var table = new DataTable(tableName);
            var column = new DataColumn("id", typeof(System.Int32))
            {
                AutoIncrement = true
            };
            table.Columns.Add(column);

            column = new DataColumn("item", typeof(System.String));
            table.Columns.Add(column);

            DataRow row;
            int rowSize = 10;
            for (int i = 0; i < rowSize; ++i)
            {
                row = table.NewRow();
                row["item"] = "item " + i;
                table.Rows.Add(row);
            }

            table.AcceptChanges();
            return table;
        }
    }
}
