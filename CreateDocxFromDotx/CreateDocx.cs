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

                    var docTables = document.MainDocumentPart.Document.Body.Elements<Table>().ToList();

                    var dataTable = Utility.CreateTestDataTable("TestTable");
                    InsertSimpleTable(dataTable, docTables[0]);

                    dataTable = Utility.CreateTestDataTableWith5Column("TestTableWith5");
                    InsertSimpleTableWithAddColumn(dataTable, docTables[1]);

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
            try
            {
                var docRows = docTable.Descendants<TableRow>().ToList();
                var patternRow = (TableRow) docRows.Last().Clone();
                docRows.Last().Remove();

                for (var rIdx = 0; rIdx < dataTable.Rows.Count; ++rIdx)
                {
                    var docRow = (TableRow) patternRow.Clone();

                    var docCells = docRow.Descendants<TableCell>().ToList();
                    for (var cIdx = 0; cIdx < dataTable.Columns.Count; ++cIdx)
                        ReplaceText(docCells[cIdx], dataTable.Rows[rIdx][cIdx]);

                    docTable.Descendants<TableRow>().ToList().Last().InsertAfterSelf(docRow);
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        private static void InsertSimpleTableWithAddColumn(DataTable dataTable, OpenXmlElement docTable)
        {
            try
            {
                var docRows = docTable.Descendants<TableRow>().ToList();
                var patternRow = (TableRow)docRows.Last().Clone();
                docRows.Last().Remove();

                var docRow = docTable.Descendants<TableRow>().ToList().Last();
                var docCells = docRow.Descendants<TableCell>().ToList();
                var midCell = (TableCell)docCells[1].Clone();
                ReplaceText(docCells[1], dataTable.Columns[1]);
                for (var cIdx = 2; cIdx < dataTable.Columns.Count - 1; ++cIdx)
                {
                    var docCell = (TableCell)midCell.Clone();
                    ReplaceText(docCell, dataTable.Columns[cIdx]);
                    docRow.Descendants<TableCell>().ToList()[cIdx - 1].InsertAfterSelf(docCell);

                    docCell = (TableCell) patternRow.Descendants<TableCell>().ToList()[1].Clone();
                    patternRow.Descendants<TableCell>().ToList()[cIdx - 1].InsertAfterSelf(docCell);
                }
                ReplaceText(docCells.Last(), dataTable.Columns[dataTable.Columns.Count - 1]);

                for (var rIdx = 0; rIdx < dataTable.Rows.Count; ++rIdx)
                {
                    docRow = (TableRow)patternRow.Clone();

                    docCells = docRow.Descendants<TableCell>().ToList();
                    for (var cIdx = 0; cIdx < dataTable.Columns.Count; ++cIdx)
                        ReplaceText(docCells[cIdx], dataTable.Rows[rIdx][cIdx]);

                    docTable.Descendants<TableRow>().ToList().Last().InsertAfterSelf(docRow);
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
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
