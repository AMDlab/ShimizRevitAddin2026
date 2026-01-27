using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Windows.Interop;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using ShimizRevitAddin2026.ExternalEvents;
using ShimizRevitAddin2026.Services;
using ShimizRevitAddin2026.UI.Models;
using ShimizRevitAddin2026.UI.Windows;

namespace ShimizRevitAddin2026.Commands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    internal class OpenRebarTagCheckWindowCommand : IExternalCommand
    {
        private const string BendingDetailTagName = "曲げ加工詳細";
        private static readonly RebarTagCheckWindowStore _windowStore = new RebarTagCheckWindowStore();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (_windowStore.TryActivateExisting())
                {
                    return Result.Succeeded;
                }

                var uiapp = commandData.Application;
                var uidoc = uiapp.ActiveUIDocument;
                var doc = uidoc.Document;
                var activeView = doc.ActiveView;

                var rebars = CollectRebars(doc, activeView);
                var dependentTagCollector = BuildDependentTagCollector();
                var mismatchRebarIds = CollectMismatchRebarIds(doc, activeView, rebars, dependentTagCollector);
                var items = BuildRebarItems(rebars, mismatchRebarIds);
                var rebarCount = GetRebarCount(rebars);

                var highlighter = BuildHighlighter(dependentTagCollector);
                var externalEventService = new RebarTagHighlightExternalEventService(highlighter);
                var window = new RebarTagCheckWindow(uidoc, activeView, externalEventService, items, rebarCount);
                SetOwner(uiapp, window);

                _windowStore.SetCurrent(window);
                window.Show();
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

        private IReadOnlyList<RebarListItem> BuildRebarItems(IReadOnlyList<Rebar> rebars, IReadOnlyCollection<ElementId> mismatchRebarIds)
        {
            if (rebars == null) return new List<RebarListItem>();

            return rebars
                .Select(r => new RebarListItem(r.Id, BuildDisplayText(r), IsMismatchRebar(mismatchRebarIds, r.Id)))
                .ToList();
        }

        private bool IsMismatchRebar(IReadOnlyCollection<ElementId> mismatchRebarIds, ElementId rebarId)
        {
            if (mismatchRebarIds == null || mismatchRebarIds.Count == 0)
            {
                return false;
            }

            if (rebarId == null || rebarId == ElementId.InvalidElementId)
            {
                return false;
            }

            return mismatchRebarIds.Contains(rebarId);
        }

        private int GetRebarCount(IReadOnlyList<Rebar> rebars)
        {
            return rebars == null ? 0 : rebars.Count;
        }

        private string BuildDisplayText(Rebar rebar)
        {
            if (rebar == null) return string.Empty;

            var typeName = GetTypeName(rebar);
            if (string.IsNullOrWhiteSpace(typeName))
            {
                return $"ID: {rebar.Id.Value}";
            }

            return $"{typeName} / ID: {rebar.Id.Value}";
        }

        private string GetTypeName(Rebar rebar)
        {
            try
            {
                if (rebar == null)
                {
                    return string.Empty;
                }

                var doc = rebar.Document;
                if (doc == null)
                {
                    return string.Empty;
                }

                var typeId = rebar.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId)
                {
                    return string.Empty;
                }

                var type = doc.GetElement(typeId) as ElementType;
                return type == null ? string.Empty : (type.Name ?? string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return string.Empty;
            }
        }

        private RebarDependentTagCollector BuildDependentTagCollector()
        {
            var matcher = new RebarTagNameMatcher(BendingDetailTagName);
            return new RebarDependentTagCollector(matcher);
        }

        private IReadOnlyCollection<ElementId> CollectMismatchRebarIds(
            Document doc,
            View activeView,
            IReadOnlyList<Rebar> rebars,
            RebarDependentTagCollector dependentTagCollector)
        {
            try
            {
                if (doc == null || activeView == null || rebars == null || dependentTagCollector == null)
                {
                    return new List<ElementId>();
                }

                var checker = new RebarTagLeaderBendingDetailConsistencyService(dependentTagCollector);
                return checker.CollectMismatchRebarIds(doc, rebars, activeView);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ElementId>();
            }
        }

        private RebarTagHighlighter BuildHighlighter(RebarDependentTagCollector collector)
        {
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

