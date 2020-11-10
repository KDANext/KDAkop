using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace KDAkop.Lab2
{
    public partial class ExelTable : Component
    {
        public ExelTable()
        {
            InitializeComponent();
        }

        public ExelTable(IContainer container)
        {
            container.Add(this);
            InitializeComponent();
        }

        public void Create<T>(String Path,bool rotation, List<String> head, IEnumerable<IEnumerable<T>> data)
        {
            Document document = new Document();
            DefineStyles(document);
            var table = document.AddSection().AddTable();
            List<string> columns = new List<string>();
            if (rotation)
            {
                HorizontalTable(head, table, data);              
            }
            else
            {
                VerticalTable(head, table, data);     
            }
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            pdfRenderer.PdfDocument.Save(Path);
        }

        private void VerticalTable<T>(List<string> head, Table table, IEnumerable<IEnumerable<T>> data)
        {
            table.AddColumn("3cm");
            foreach (var i in data)
            {
                table.AddColumn("3cm");
            }
            var iter = head.GetEnumerator();
            foreach (var varObject in data)
            {
                var Text = new List<String>();
                iter.MoveNext();
                Text.Add(iter.Current);
                foreach (var temp in varObject)
                {
                    Text.Add(temp.ToString());
                }
                CreateRow(new PdfRowParameters
                {
                    Table = table,
                    Texts = Text,
                    Style = "Normal",
                    ParagraphAlignment = ParagraphAlignment.Left
                });

            }
        }

        private void HorizontalTable<T>(List<String> head,Table table, IEnumerable<IEnumerable<T>> data)
        {
            foreach (var i in head)
            {
                table.AddColumn("3cm");
            }
            CreateRow(new PdfRowParameters
            {
                Table = table,
                Texts = head,
                Style = "NormalTitle",
                ParagraphAlignment = ParagraphAlignment.Center
            });
            foreach (var varObject in data)
            {
                var Text = new List<String>();
                foreach (var temp in varObject)
                {
                    Text.Add(temp.ToString());
                }
                CreateRow(new PdfRowParameters
                {
                    Table = table,
                    Texts = Text,
                    Style = "Normal",
                    ParagraphAlignment = ParagraphAlignment.Left
                });
            }
        }

        private void DefineStyles(Document document)
        {
            Style style = document.Styles["Normal"];
            style.Font.Name = "Times New Roman";
            style.Font.Size = 14;
            style = document.Styles.AddStyle("NormalTitle", "Normal");
            style.Font.Bold = true;
        }

        private static void CreateRow(PdfRowParameters rowParameters)
        {
            Row row = rowParameters.Table.AddRow();
            for (int i = 0; i < rowParameters.Texts.Count; ++i)
            {
                FillCell(new PdfCellParameters
                {
                    Cell = row.Cells[i],
                    Text = rowParameters.Texts[i],
                    Style = rowParameters.Style,
                    BorderWidth = 0.5,
                    ParagraphAlignment = rowParameters.ParagraphAlignment
                });
            }
        }

        private static void FillCell(PdfCellParameters cellParameters)
        {
            cellParameters.Cell.AddParagraph(cellParameters.Text);
            if (!string.IsNullOrEmpty(cellParameters.Style))
            {
                cellParameters.Cell.Style = cellParameters.Style;
            }
            cellParameters.Cell.Borders.Left.Width = cellParameters.BorderWidth;
            cellParameters.Cell.Borders.Right.Width = cellParameters.BorderWidth;
            cellParameters.Cell.Borders.Top.Width = cellParameters.BorderWidth;
            cellParameters.Cell.Borders.Bottom.Width = cellParameters.BorderWidth;
            cellParameters.Cell.Format.Alignment = cellParameters.ParagraphAlignment;
            cellParameters.Cell.VerticalAlignment = VerticalAlignment.Center;
        }
    }
}
