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
                var consistencyService = BuildConsistencyService(dependentTagCollector);
                var mismatchRebarIds = CollectMismatchRebarIds(doc, activeView, rebars, consistencyService);
                var hostRebarIdsWithBendingDetail = CollectHostRebarIdsWithBendingDetail(doc, activeView, consistencyService);
                var items = BuildRebarItems(doc, activeView, rebars, mismatchRebarIds, dependentTagCollector, hostRebarIdsWithBendingDetail);
                var rebarCount = GetRebarCount(items);

                var highlighter = BuildHighlighter(dependentTagCollector);
                var externalEventService = new RebarTagHighlightExternalEventService(highlighter, consistencyService);
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

        private IReadOnlyList<RebarListItem> BuildRebarItems(
            Document doc,
            View activeView,
            IReadOnlyList<Rebar> rebars,
            IReadOnlyCollection<ElementId> mismatchRebarIds,
            RebarDependentTagCollector dependentTagCollector,
            IReadOnlyCollection<ElementId> hostRebarIdsWithBendingDetail)
        {
            if (rebars == null) return new List<RebarListItem>();

            return rebars
                .Where(r => !ShouldHideRebar(doc, activeView, r, dependentTagCollector, hostRebarIdsWithBendingDetail))
                .Select(r => new RebarListItem(r.Id, BuildDisplayText(r), IsMismatchRebar(mismatchRebarIds, r.Id)))
                .ToList();
        }

        private bool ShouldHideRebar(
            Document doc,
            View activeView,
            Rebar rebar,
            RebarDependentTagCollector dependentTagCollector,
            IReadOnlyCollection<ElementId> hostRebarIdsWithBendingDetail)
        {
            // 構造タグがあり、かつ曲げ詳細鉄筋（曲げ詳細要素）が無い場合は表示しない
            if (!HasStructureRebarTag(doc, activeView, rebar, dependentTagCollector))
            {
                return false;
            }

            if (rebar == null || rebar.Id == null || rebar.Id == ElementId.InvalidElementId)
            {
                return false;
            }

            if (hostRebarIdsWithBendingDetail == null || hostRebarIdsWithBendingDetail.Count == 0)
            {
                return true;
            }

            return !hostRebarIdsWithBendingDetail.Contains(rebar.Id);
        }

        private bool HasStructureRebarTag(
            Document doc,
            View activeView,
            Rebar rebar,
            RebarDependentTagCollector dependentTagCollector)
        {
            try
            {
                if (doc == null || activeView == null || rebar == null || dependentTagCollector == null)
                {
                    return false;
                }

                var model = dependentTagCollector.Collect(doc, rebar, activeView);
                return model != null && model.StructureTagIds != null && model.StructureTagIds.Count > 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return false;
            }
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

        private int GetRebarCount(IReadOnlyList<RebarListItem> items)
        {
            return items == null ? 0 : items.Count;
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
            RebarTagLeaderBendingDetailConsistencyService consistencyService)
        {
            try
            {
                if (doc == null || activeView == null || rebars == null || consistencyService == null)
                {
                    return new List<ElementId>();
                }

                return consistencyService.CollectMismatchRebarIds(doc, rebars, activeView);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ElementId>();
            }
        }

        private IReadOnlyCollection<ElementId> CollectHostRebarIdsWithBendingDetail(
            Document doc,
            View activeView,
            RebarTagLeaderBendingDetailConsistencyService consistencyService)
        {
            try
            {
                if (doc == null || activeView == null || consistencyService == null)
                {
                    return new List<ElementId>();
                }

                return consistencyService.CollectHostRebarIdsWithBendingDetail(doc, activeView);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ElementId>();
            }
        }

        private RebarTagLeaderBendingDetailConsistencyService BuildConsistencyService(RebarDependentTagCollector dependentTagCollector)
        {
            return new RebarTagLeaderBendingDetailConsistencyService(dependentTagCollector);
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

