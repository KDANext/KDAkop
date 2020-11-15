using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Text;
using Microsoft.Office.Interop.Word;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace components36
{
    public partial class ComponentWordSummary : Component
    {

        public int columnCount;

        public int rowCount;

        private List<List<String>> data;
        public ComponentWordSummary()
        {
            InitializeComponent();
        }

        public ComponentWordSummary(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }
        
        
        public void SetData(List<List<String>> contentsColumns)
        {
            this.data = contentsColumns;
        }
        
        public void CreateTable(String[] columnNames, String[] rowNames, String fileName)
        {
            var winword = new Microsoft.Office.Interop.Word.Application();
            try
            {
               // rowCount = rowNames!=null ? rowNames.Length : 0;
               // columnCount = columnNames!=null? columnNames.Length : 0;
                object missing = System.Reflection.Missing.Value;
                //создаем документ
                Microsoft.Office.Interop.Word.Document document =
                winword.Documents.Add(ref missing, ref missing, ref missing, ref
               missing);
                //получаем ссылку на параграф
                var paragraph = document.Paragraphs.Add(missing);
                var range = paragraph.Range;
                //задаем настройки шрифта
                var font = range.Font;
                font.Size = 16;
                font.Name = "Times New Roman";
                font.Bold = 1;
                //задаем настройки абзаца
                var paragraphFormat = range.ParagraphFormat;
                paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                paragraphFormat.SpaceAfter = 10;
                paragraphFormat.SpaceBefore = 0;
                //добавляем абзац в документ
                range.InsertParagraphAfter();
                //создаем таблицу
                var paragraphTable = document.Paragraphs.Add(Type.Missing);
                var rangeTable = paragraphTable.Range;
                var table = document.Tables.Add(rangeTable, rowCount + (columnNames != null ? 1 : 0), columnCount + (rowNames != null ? 1 : 0), ref missing, ref missing);
                font = table.Range.Font;
                font.Size = 14;
                font.Name = "Times New Roman";
                var paragraphTableFormat = table.Range.ParagraphFormat;
                paragraphTableFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                paragraphTableFormat.SpaceAfter = 0;
                paragraphTableFormat.SpaceBefore = 0;
                int columnOffset = 0;
                int rowOffset = 0;
                if (columnNames != null)
                {
                    for (int i = 0; i < columnCount; i++)
                    {
                        table.Cell(1, i + 1 + (rowNames != null ? 1 : 0)).Range.Text = columnNames[i];
                    }
                    rowOffset = 1;
                }
                if (rowNames != null)
                {
                    for (int i = 0; i < rowCount; i++)
                    {
                        table.Cell(i + 1 + (columnNames != null ? 1 : 0), 1).Range.Text = rowNames[i];
                    }
                    columnOffset = 1;
                }
                for (int i = 0; i < rowCount; ++i)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        table.Cell(i + rowOffset + 1, j + columnOffset + 1).Range.Text = data[j][i];
                    }
                }
                //задаем границы таблицы
                table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleInset;
                table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                object fileFormat = WdSaveFormat.wdFormatXMLDocument;
                document.SaveAs(fileName, ref fileFormat, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing);
                document.Close(ref missing, ref missing, ref missing);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                winword.Quit();
                System.Diagnostics.Process.Start(fileName);
            }

        }
        /* public void saveData<T>(String path, bool IsHorisontalHead, T[] data)
         {
             //Create Table
             Document doc = new Document();
             Section s = doc.AddSection();
             Table table = s.AddTable(true);

             //Create Header and Data
             /*String[] Header = { "Item", "Description", "Qty", "Unit Price", "Price" };
             String[][] data = {
                                   new String[]{ "Spire.Doc for .NET",".NET Word Component","1","$799.00","$799.00"},
                                   new String[]{"Spire.XLS for .NET",".NET Excel Component","2","$799.00","$1,598.00"},
                                   new String[]{"Spire.Office for .NET",".NET Office Component","1","$1,899.00","$1,899.00"},
                                   new String[]{"Spire.PDF for .NET",".NET PDFComponent","2","$599.00","$1,198.00"},
                               };
             //Add Cells
             table.ResetCells(data.Length + 1, data.Length);

             //Header Row
             TableRow FRow = table.Rows[0];
             FRow.IsHeader = true;
             //Row Height
             FRow.Height = 23;
             var fields = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance | BindingFlags.DeclaredOnly);
             var fieldNumHead = 1;
             //Header Format           
             foreach (var field in fields)
             {
                 //Cell Alignment
                 var column = this.GetWordColumnName(IsHorisontalHead ? fieldNumHead : 1);
                 var rowIndex = IsHorisontalHead ? (int)1 : (int)fieldNumHead;
                 Paragraph p = FRow.Cells[rowIndex].AddParagraph();
                 FRow.Cells[rowIndex].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                 p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                 //Data Format
                 TextRange TR = p.AppendText(field.Name);
                 TR.CharacterFormat.FontName = "Calibri";
                 TR.CharacterFormat.FontSize = 14;
                 TR.CharacterFormat.TextColor = Color.Teal;
                 TR.CharacterFormat.Bold = true;
                 fieldNumHead++;
             }
             //Data Row
             var dataRowNum = 2;
             foreach(var dataRow in data)
             {
                 var fieldNum = 1;
                 TableRow DataRow = table.Rows[fieldNum];

                 //Row Height
                 DataRow.Height = 20;

                 //C Represents Column.
                 foreach(var field in fields)
                 {
                     var column = string.Empty;
                     var rowIndex = 0U;
                     column = this.GetWordColumnName(IsHorisontalHead ? fieldNum : dataRowNum);
                     rowIndex = (uint)(IsHorisontalHead ? dataRowNum : fieldNum);
                     //Fill Data in Rows
                     Paragraph p2 = DataRow.Cells[Convert.ToInt32(rowIndex)].AddParagraph();
                     TextRange TR2 = p2.AppendText(field.GetValue(dataRow).ToString());
                     //Format Cells
                     p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                     TR2.CharacterFormat.FontName = "Calibri";
                     TR2.CharacterFormat.FontSize = 12;
                     TR2.CharacterFormat.TextColor = Color.Brown;
                     fieldNum++;
                 }
                 dataRowNum++;
             }

             //Save and Launch
             doc.SaveToFile(path, FileFormat.Docx2013);
             System.Diagnostics.Process.Start(path);
         }
         private string GetWordColumnName(int columnNumber)
         {
             int dividend = columnNumber;
             string columnName = string.Empty;
             int modulo;

             while (dividend > 0)
             {
                 modulo = (dividend - 1) % 26;
                 columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                 dividend = (int)((dividend - modulo) / 26);
             }

             return columnName;
         }*/

    }
}
