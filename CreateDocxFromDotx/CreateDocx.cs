using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

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

                    InsertAPicture(document, Path.Combine(SampleFolder, "picture1.jpg"));
                    InsertAPicture(document, Path.Combine(SampleFolder, "picture1.jpg"));
                    InsertAPicture(document, Path.Combine(SampleFolder, "picture1.jpg"));

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

        public static void InsertAPicture(WordprocessingDocument document, string fileNamePic)
        {
            var mainPart = document.MainDocumentPart;
            var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (var stream = new FileStream(fileNamePic, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            AddImageToBody(document, mainPart.GetIdOfPart(imagePart));
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            const long picSizeCx = 990000L + 990000L;
            const long picSizeCy = picSizeCx;

            var pictureProperties = new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties()
                {
                    Id = (UInt32Value) 0U,
                    Name = "New Bitmap Image.jpg"
                },
                new PIC.NonVisualPictureDrawingProperties());

            var picBlipFill = new PIC.BlipFill(
                new A.Blip(
                    new A.BlipExtensionList(
                        new A.BlipExtension()
                        {
                            Uri ="{28A0092B-C50C-407E-A947-70E740481C1C}"
                        }))
                {
                    Embed = relationshipId,
                    CompressionState = A.BlipCompressionValues.Print
                },
                new A.Stretch(new A.FillRectangle()));

            var picShapeProperties = new PIC.ShapeProperties(
                new A.Transform2D(
                    new A.Offset()
                    {
                        X = 0L,
                        Y = 0L
                    },
                    new A.Extents()
                    {
                        Cx = picSizeCx,
                        Cy = picSizeCy
                    }),
                new A.PresetGeometry(new A.AdjustValueList())
                {
                    Preset = A.ShapeTypeValues.Rectangle
                });

            var aGraphicData = new A.GraphicData(
                new PIC.Picture(
                    pictureProperties, picBlipFill, picShapeProperties))
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
            };

            var dwInline = new DW.Inline(
                new DW.Extent()
                {
                    Cx = picSizeCx,
                    Cy = picSizeCy
                },
                new DW.EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties()
                {
                    Id = (UInt32Value) 1U,
                    Name = "Picture 1"
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks()
                    {
                        NoChangeAspect = true
                    }),
                new A.Graphic(aGraphicData));

            // Define the reference of the image.
            var element = new Drawing(dwInline);

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(
                new Paragraph(new Run(element)));
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
