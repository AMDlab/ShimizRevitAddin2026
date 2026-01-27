using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Windows.Interop;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using ShimizRevitAddin2026.Services;
using ShimizRevitAddin2026.UI.Models;
using ShimizRevitAddin2026.UI.Windows;

namespace ShimizRevitAddin2026.Commands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    internal class OpenRebarTagCheckWindowCommand : IExternalCommand
    {
        private const string BendingDetailTagName = "曲げ加工詳細";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiapp = commandData.Application;
                var uidoc = uiapp.ActiveUIDocument;
                var doc = uidoc.Document;
                var activeView = doc.ActiveView;

                var rebars = CollectRebars(doc, activeView);
                var items = BuildRebarItems(rebars);

                var highlighter = BuildHighlighter();
                var window = new RebarTagCheckWindow(uidoc, activeView, highlighter, items);
                SetOwner(uiapp, window);

                window.ShowDialog();
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RebarTag", ex.ToString());
                message = ex.Message;
                return Result.Failed;
            }
        }

        private IReadOnlyList<Rebar> CollectRebars(Document doc, View activeView)
        {
            var collector = new RebarCollectorInView();
            return collector.Collect(doc, activeView);
        }

        private IReadOnlyList<RebarListItem> BuildRebarItems(IReadOnlyList<Rebar> rebars)
        {
            if (rebars == null) return new List<RebarListItem>();

            return rebars
                .Select(r => new RebarListItem(r.Id, BuildDisplayText(r)))
                .ToList();
        }

        private string BuildDisplayText(Rebar rebar)
        {
            if (rebar == null) return string.Empty;

            var mark = GetMark(rebar);
            if (string.IsNullOrWhiteSpace(mark))
            {
                return $"ID: {rebar.Id.Value}";
            }

            return $"ID: {rebar.Id.Value} / Mark: {mark}";
        }

        private string GetMark(Rebar rebar)
        {
            try
            {
                var p = rebar.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                return p == null ? string.Empty : p.AsString();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return string.Empty;
            }
        }

        private RebarTagHighlighter BuildHighlighter()
        {
            var matcher = new RebarTagNameMatcher(BendingDetailTagName);
            var collector = new RebarDependentTagCollector(matcher);
            return new RebarTagHighlighter(collector);
        }

        private void SetOwner(UIApplication uiapp, System.Windows.Window window)
        {
            if (uiapp == null || window == null)
            {
                return;
            }

            try
            {
                var helper = new WindowInteropHelper(window);
                helper.Owner = uiapp.MainWindowHandle;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RebarTag", ex.ToString());
            }
        }
    }
}

