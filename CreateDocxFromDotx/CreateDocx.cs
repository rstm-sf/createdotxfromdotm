using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CreateDocxFromDotx
{
    public class CreateDocx
    {
        public CreateDocx()
        {
            try
            {
                File.Copy(_sourceFile, _destinationFile, true);
                using (var document = WordprocessingDocument.Open(_destinationFile, true))
                {
                    document.ChangeDocumentType(WordprocessingDocumentType.Document);

                    var tableMem = Utility.CreateTestDataTable("TestTable");
                    var docTable = document.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault();

                    if (docTable != null)
                    {
                        var rows = docTable.Descendants<TableRow>().ToList();
                        var myRow = (TableRow)rows.Last().Clone();
                        rows.Last().Remove();

                        for (var rowIdx = 0; rowIdx < tableMem.Rows.Count; ++rowIdx)
                        {
                            var rowDoc = (TableRow)myRow.Clone();
                            var cellsDoc = rowDoc.Descendants<TableCell>().ToList();
                            for (var cellIdx = 0; cellIdx < tableMem.Columns.Count; ++cellIdx)
                                ReplaceText(cellsDoc[cellIdx], tableMem.Rows[rowIdx][cellIdx]);

                            docTable.Descendants<TableRow>().Last().InsertAfterSelf(rowDoc);
                        }
                    }

                    document.Save();
                    Console.WriteLine("Document generated at " + _destinationFile);
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        public void Open()
        {
            Process.Start(_destinationFile);
        }

        private static void ReplaceText(OpenXmlElement tCell, object dCell)
        {
            try
            {
                var first = tCell.Descendants<Text>().FirstOrDefault();
                first.Text = dCell.ToString();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        private static readonly string SampleFolder = Path.Combine(Environment.CurrentDirectory, "Sample");
        private readonly string _destinationFile = Path.Combine(SampleFolder, "Doc.docx");
        private readonly string _sourceFile = Path.Combine(SampleFolder, "TemplateDoc.dotx");
    }
}
