using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace KDAkop
{
    public partial class WordDiagram : Component
    {
        public WordDiagram()
        {
            InitializeComponent();
        }

        public WordDiagram(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }

        public void CreateDiagramInWord(List<String> Legend, List<int> Values, string path)
        {
            DocX document = DocX.Create(path);
            Series series = CreateSeries(Legend, Values);
            document.InsertChart(CreateBarChart(series));
            document.Save();
        }
        
        private static Chart CreateBarChart(Series series)
        {
            PieChart pieChart = new PieChart();
            pieChart.AddLegend(ChartLegendPosition.Top, false);
            pieChart.AddSeries(series);
            return pieChart;
        }

        public static Series CreateSeries(List<String> Legend, List<int> Values)
        {
            Series series = new Series("");
            series.Bind(Legend, Values);
            return series;
        }
    }
}
