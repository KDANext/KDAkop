using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
namespace components36
{
    public partial class ComponentExcelDiagram : Component
    {
        public ComponentExcelDiagram()
        {
            InitializeComponent();
        }

        public ComponentExcelDiagram(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }
        public void BuiltChart(List<Setting> list, string fileName)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            var excel = new Application();

            try
            {
                excel.SheetsInNewWorkbook = 1;
                excel.Workbooks.Add(Type.Missing);
                excel.Workbooks[1].SaveAs(fileName, Type.Missing,
                                            Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, XlSaveAsAccessMode.xlNoChange,
                                            Type.Missing, Type.Missing, Type.Missing,
                                            Type.Missing, Type.Missing);

                Sheets excelsheets = excel.Workbooks[1].Worksheets;
                var excelworksheet = (Worksheet)excelsheets.get_Item(1);
                excelworksheet.Cells.Clear();
                excelworksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                excelworksheet.PageSetup.CenterHorizontally = true;
                excelworksheet.PageSetup.CenterVertically = true;


                for (int i = 0; i < list.Count; i++)
                {
                    excelworksheet.Cells[1, i + 1] = list[i].legend;
                    excelworksheet.Cells[2, i + 1] = list[i].value;
                }

                ChartObjects chartObjs = (ChartObjects)excelworksheet.ChartObjects();
                ChartObject chartObj = chartObjs.Add(5, 50, 300, 300);
                Chart xlChart = chartObj.Chart;


                Range rng2 = excelworksheet.Range["A1", (Convert.ToChar(65 + list.Count - 1)).ToString() + "2"];

                xlChart.ChartType = XlChartType.xlPie;
                xlChart.SetSourceData(rng2);
                Series series = (Series)xlChart.SeriesCollection(1);
                xlChart.Legend.Delete();
                series.HasDataLabels = true;
                excel.Workbooks[1].Save();
                excel.Workbooks.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
            finally
            {
                excel.Quit();
            }
        }
    }
}
