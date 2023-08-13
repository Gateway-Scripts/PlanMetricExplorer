using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

[assembly: AssemblyVersion("1.0.0.1")]

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context, System.Windows.Window window, ScriptEnvironment environment)
        {
            // TODO : Add here the code that is called when the script is launched from Eclipse.
            PlanSetup plan = context.PlanSetup;
            if(plan == null)
            {
                MessageBox.Show("No valid plan selected");
                return;
            }
            if (!plan.IsDoseValid)
            {
                MessageBox.Show("The plan selected has no valid dose.");
                return;
            }
            Structure target = plan.StructureSet.Structures.FirstOrDefault(st=>st.Id.Equals(plan.TargetVolumeID));
            if(target == null)
            {
                MessageBox.Show("Plan contains no target volume");
                return;
            }

            //set up plan metrics
            Dictionary<string, string> planMetrics = new Dictionary<string, string>();
            //add metrics to report.
            planMetrics.Add("Target", target.Id);
            planMetrics.Add("Target Volume", target.Volume.ToString("F1")+"cc");

            window.Width = 450;
            window.Height = 600;
            window.Content = AddMetricsToView(planMetrics);
            
        }
        private FlowDocumentScrollViewer AddMetricsToView(Dictionary<string,string> parameters)
        {
            FlowDocumentScrollViewer flowScroller = new FlowDocumentScrollViewer();
            FlowDocument flowDocument = new FlowDocument();
            flowDocument.Blocks.Add(new Paragraph(new Run("Dosimetric Plan Metrics")) { TextAlignment = TextAlignment.Center });
            //add values to table.
            Table table = new Table();
            table.RowGroups.Add(new TableRowGroup());
            table.RowGroups.First().Rows.Add(new TableRow());
            table.RowGroups.Last().Rows.Last().Cells.Add(new TableCell(new Paragraph(new Run("Property") { FontWeight = FontWeights.Bold })));
            table.RowGroups.Last().Rows.Last().Cells.Add(new TableCell(new Paragraph(new Run("Value") { FontWeight = FontWeights.Bold })));
            foreach(var metric in parameters)
            {
                table.RowGroups.First().Rows.Add(new TableRow());
                table.RowGroups.Last().Rows.Last().Cells.Add(new TableCell(new Paragraph(new Run(metric.Key))) { BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(Color.FromRgb(0x1E,0x52,0x88)) });
                table.RowGroups.Last().Rows.Last().Cells.Add(new TableCell(new Paragraph(new Run(metric.Value))) { BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(Color.FromRgb(0x1E,0x52,0x88)) });
            }
            flowDocument.Blocks.Add(table);
            flowScroller.Document = flowDocument;
            return flowScroller;
        }
    }
}
