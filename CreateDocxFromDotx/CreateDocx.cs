using System;
using System.Data;
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

                    var dataTable = Utility.CreateTestDataTable("TestTable");
                    var docTable = document.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault();

                    if (docTable != null)
                    {
                        InsertSimpleTable(dataTable, docTable);
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

        private static void InsertSimpleTable(DataTable dataTable, OpenXmlElement docTable)
        {
            var docRows = docTable.Descendants<TableRow>().ToList();
            var patternRow = (TableRow)docRows.Last().Clone();
            docRows.Last().Remove();

            for (var rIdx = 0; rIdx < dataTable.Rows.Count; ++rIdx)
            {
                var docRow = (TableRow)patternRow.Clone();

                var docCells = docRow.Descendants<TableCell>().ToList();
                for (var cIdx = 0; cIdx < dataTable.Columns.Count; ++cIdx)
                    ReplaceText(docCells[cIdx], dataTable.Rows[rIdx][cIdx]);

                docTable.Descendants<TableRow>().ToList().Last().InsertAfterSelf(docRow);
            }
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
