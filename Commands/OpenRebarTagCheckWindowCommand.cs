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

                var dependentTagCollector = BuildDependentTagCollector();
                var consistencyService = BuildConsistencyService(dependentTagCollector);

                var (sheet, targetViews, includeViewNamePrefix) = ResolveTargetViews(doc, activeView);
                var items = CollectRebarItemsForViews(doc, targetViews, dependentTagCollector, consistencyService, includeViewNamePrefix);
                var rebarCount = GetRebarCount(items);

                var highlighter = BuildHighlighter(dependentTagCollector);
                var externalEventService = new RebarTagHighlightExternalEventService(highlighter, consistencyService);
                var window = BuildWindow(uidoc, activeView, sheet, targetViews, externalEventService, items, rebarCount);
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

        private IReadOnlyList<RebarListItem> CollectRebarItemsForViews(
            Document doc,
            IReadOnlyList<View> views,
            RebarDependentTagCollector dependentTagCollector,
            RebarTagLeaderBendingDetailConsistencyService consistencyService,
            bool includeViewNamePrefix)
        {
            if (views == null || views.Count == 0)
            {
                return new List<RebarListItem>();
            }

            var result = new List<RebarListItem>();
            foreach (var view in views)
            {
                if (view == null)
                {
                    continue;
                }

                result.AddRange(CollectRebarItemsForSingleView(doc, view, dependentTagCollector, consistencyService, includeViewNamePrefix));
            }

            return result;
        }

        private IReadOnlyList<RebarListItem> CollectRebarItemsForSingleView(
            Document doc,
            View view,
            RebarDependentTagCollector dependentTagCollector,
            RebarTagLeaderBendingDetailConsistencyService consistencyService,
            bool includeViewNamePrefix)
        {
            var rebars = CollectRebars(doc, view);
            var (mismatchRebarIds, leaderLineNotFoundRebarIds) = CollectIssueRebarIds(doc, view, rebars, consistencyService);
            var hostRebarIdsWithBendingDetail = CollectHostRebarIdsWithBendingDetail(doc, view, consistencyService);
            return BuildRebarItems(doc, view, rebars, mismatchRebarIds, leaderLineNotFoundRebarIds, dependentTagCollector, hostRebarIdsWithBendingDetail, includeViewNamePrefix);
        }

        private (ViewSheet sheet, IReadOnlyList<View> views, bool includeViewNamePrefix) ResolveTargetViews(Document doc, View activeView)
        {
            var sheet = activeView as ViewSheet;
            if (sheet == null)
            {
                return (null, new List<View> { activeView }, false);
            }

            var collector = new SheetPlacedViewCollector();
            var views = collector.Collect(doc, sheet);
            return (sheet, views, true);
        }

        private RebarTagCheckWindow BuildWindow(
            UIDocument uidoc,
            View activeView,
            ViewSheet sheet,
            IReadOnlyList<View> targetViews,
            RebarTagHighlightExternalEventService externalEventService,
            IReadOnlyList<RebarListItem> items,
            int rebarCount)
        {
            if (sheet == null)
            {
                return new RebarTagCheckWindow(uidoc, activeView, externalEventService, items, rebarCount);
            }

            return new RebarTagCheckWindow(uidoc, sheet, targetViews ?? new List<View>(), externalEventService, items, rebarCount);
        }

        private IReadOnlyList<RebarListItem> BuildRebarItems(
            Document doc,
            View activeView,
            IReadOnlyList<Rebar> rebars,
            IReadOnlyCollection<ElementId> mismatchRebarIds,
            IReadOnlyCollection<ElementId> leaderLineNotFoundRebarIds,
            RebarDependentTagCollector dependentTagCollector,
            IReadOnlyCollection<ElementId> hostRebarIdsWithBendingDetail,
            bool includeViewNamePrefix)
        {
            if (rebars == null) return new List<RebarListItem>();

            return rebars
                .Where(r => !ShouldHideRebar(doc, activeView, r, dependentTagCollector, hostRebarIdsWithBendingDetail))
                .Select(r => BuildRebarListItem(doc, activeView, r, mismatchRebarIds, leaderLineNotFoundRebarIds, dependentTagCollector, hostRebarIdsWithBendingDetail, includeViewNamePrefix))
                .ToList();
        }

        private RebarListItem BuildRebarListItem(
            Document doc,
            View activeView,
            Rebar rebar,
            IReadOnlyCollection<ElementId> mismatchRebarIds,
            IReadOnlyCollection<ElementId> leaderLineNotFoundRebarIds,
            RebarDependentTagCollector dependentTagCollector,
            IReadOnlyCollection<ElementId> hostRebarIdsWithBendingDetail,
            bool includeViewNamePrefix)
        {
            if (rebar == null)
            {
                return new RebarListItem(ElementId.InvalidElementId, ElementId.InvalidElementId, string.Empty, false, false, false);
            }

            var isMismatch = IsMismatchRebar(mismatchRebarIds, rebar.Id);
            var isLeaderLineNotFound = IsMismatchRebar(leaderLineNotFoundRebarIds, rebar.Id);
            var isNoTagAndNoBendingDetail = IsNoTagAndNoBendingDetail(doc, activeView, rebar, dependentTagCollector, hostRebarIdsWithBendingDetail);
            return new RebarListItem(
                rebar.Id,
                activeView?.Id ?? ElementId.InvalidElementId,
                BuildDisplayText(activeView, rebar, includeViewNamePrefix),
                isMismatch,
                isLeaderLineNotFound,
                isNoTagAndNoBendingDetail);
        }

        private bool IsNoTagAndNoBendingDetail(
            Document doc,
            View activeView,
            Rebar rebar,
            RebarDependentTagCollector dependentTagCollector,
            IReadOnlyCollection<ElementId> hostRebarIdsWithBendingDetail)
        {
            // 鉄筋のみ（構造タグなし・曲げ詳細なし）を青で識別するための判定
            if (rebar == null || rebar.Id == null || rebar.Id == ElementId.InvalidElementId)
            {
                return false;
            }

            var hasStructureTag = HasStructureRebarTag(doc, activeView, rebar, dependentTagCollector);
            if (hasStructureTag)
            {
                return false;
            }

            var hasBendingDetail = HasBendingDetail(hostRebarIdsWithBendingDetail, rebar.Id);
            return !hasBendingDetail;
        }

        private bool HasBendingDetail(IReadOnlyCollection<ElementId> hostRebarIdsWithBendingDetail, ElementId rebarId)
        {
            if (rebarId == null || rebarId == ElementId.InvalidElementId)
            {
                return false;
            }

            if (hostRebarIdsWithBendingDetail == null || hostRebarIdsWithBendingDetail.Count == 0)
            {
                return false;
            }

            return hostRebarIdsWithBendingDetail.Contains(rebarId);
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

        private string BuildDisplayText(View view, Rebar rebar, bool includeViewNamePrefix)
        {
            if (rebar == null) return string.Empty;

            var typeName = GetTypeName(rebar);
            var baseText = string.IsNullOrWhiteSpace(typeName)
                ? $"ID: {rebar.Id.Value}"
                : $"{typeName} / ID: {rebar.Id.Value}";

            if (!includeViewNamePrefix)
            {
                return baseText;
            }

            var viewName = view == null ? string.Empty : (view.Name ?? string.Empty);
            if (string.IsNullOrWhiteSpace(viewName))
            {
                return baseText;
            }

            return $"{viewName} / {baseText}";
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

        private (IReadOnlyCollection<ElementId> mismatchRebarIds, IReadOnlyCollection<ElementId> leaderLineNotFoundRebarIds) CollectIssueRebarIds(
            Document doc,
            View activeView,
            IReadOnlyList<Rebar> rebars,
            RebarTagLeaderBendingDetailConsistencyService consistencyService)
        {
            try
            {
                if (doc == null || activeView == null || rebars == null || consistencyService == null)
                {
                    return (new List<ElementId>(), new List<ElementId>());
                }

                return consistencyService.CollectIssueRebarIds(doc, rebars, activeView);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (new List<ElementId>(), new List<ElementId>());
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

