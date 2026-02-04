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
        private const string TargetSheetNameToken = "配筋図";
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

                var highlighter = BuildHighlighter(dependentTagCollector);
                var externalEventService = new RebarTagHighlightExternalEventService(highlighter, consistencyService);

                var scope = SelectTargetScope();
                if (scope == TargetScope.Cancelled)
                {
                    return Result.Cancelled;
                }

                if (scope == TargetScope.ActiveContext)
                {
                    return ExecuteForActiveContext(uidoc, uiapp, doc, activeView, dependentTagCollector, consistencyService, externalEventService);
                }

                return ExecuteForAllTargetSheets(uidoc, uiapp, doc, activeView, dependentTagCollector, consistencyService, externalEventService);
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

        private Result ExecuteForActiveContext(
            UIDocument uidoc,
            UIApplication uiapp,
            Document doc,
            View activeView,
            RebarDependentTagCollector dependentTagCollector,
            RebarTagLeaderBendingDetailConsistencyService consistencyService,
            RebarTagHighlightExternalEventService externalEventService)
        {
            if (uidoc == null || uiapp == null || doc == null || activeView == null)
            {
                return Result.Cancelled;
            }

            var sheet = activeView as ViewSheet;
            if (sheet != null)
            {
                var placedViews = CollectPlacedViews(doc, sheet);
                var targets = BuildTargetsForSheets(doc, new List<ViewSheet> { sheet });
                var items = CollectRebarItemsForTargets(doc, targets, dependentTagCollector, consistencyService);
                var rebarCount = GetRebarCount(items);

                var window = new RebarTagCheckWindow(uidoc, sheet, placedViews, externalEventService, items, rebarCount);
                SetOwner(uiapp, window);
                _windowStore.SetCurrent(window);
                window.Show();
                return Result.Succeeded;
            }

            var viewItems = CollectRebarItemsForSingleView(doc, activeView, dependentTagCollector, consistencyService, string.Empty);
            var viewRebarCount = GetRebarCount(viewItems);
            var singleWindow = new RebarTagCheckWindow(uidoc, activeView, externalEventService, viewItems, viewRebarCount);
            SetOwner(uiapp, singleWindow);
            _windowStore.SetCurrent(singleWindow);
            singleWindow.Show();
            return Result.Succeeded;
        }

        private Result ExecuteForAllTargetSheets(
            UIDocument uidoc,
            UIApplication uiapp,
            Document doc,
            View activeView,
            RebarDependentTagCollector dependentTagCollector,
            RebarTagLeaderBendingDetailConsistencyService consistencyService,
            RebarTagHighlightExternalEventService externalEventService)
        {
            var sheets = CollectTargetSheets(doc);
            if (sheets.Count == 0)
            {
                TaskDialog.Show("RebarTag", $"[{TargetSheetNameToken}] を含むシートが見つかりません。");
                return Result.Cancelled;
            }

            var targets = BuildTargetsForSheets(doc, sheets);
            var items = CollectRebarItemsForTargets(doc, targets, dependentTagCollector, consistencyService);
            var rebarCount = GetRebarCount(items);

            var viewCount = CountViews(targets);
            var window = BuildWindow(uidoc, activeView, sheets, viewCount, externalEventService, items, rebarCount);
            SetOwner(uiapp, window);
            _windowStore.SetCurrent(window);
            window.Show();
            return Result.Succeeded;
        }

        private IReadOnlyList<View> CollectPlacedViews(Document doc, ViewSheet sheet)
        {
            try
            {
                var collector = new SheetPlacedViewCollector();
                return collector.Collect(doc, sheet);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<View>();
            }
        }

        private TargetScope SelectTargetScope()
        {
            // 実行時に対象範囲を選択する
            var td = new TaskDialog("RebarTag")
            {
                MainInstruction = "対象範囲を選択してください。",
                MainContent = "現在のビュー、またはシート名に[配筋図]を含む全シートを対象に検証します。",
                CommonButtons = TaskDialogCommonButtons.Cancel
            };

            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "現在のビュー（Sheetの場合はそのSheetに配置されたビュー）");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, $"全シート（シート名に{TargetSheetNameToken}を含む）");

            var r = td.Show();
            if (r == TaskDialogResult.CommandLink1)
            {
                return TargetScope.ActiveContext;
            }

            if (r == TaskDialogResult.CommandLink2)
            {
                return TargetScope.AllTargetSheets;
            }

            return TargetScope.Cancelled;
        }

        private enum TargetScope
        {
            Cancelled = 0,
            ActiveContext = 1,
            AllTargetSheets = 2,
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

                var prefix = includeViewNamePrefix ? (view.Name ?? string.Empty) : string.Empty;
                result.AddRange(CollectRebarItemsForSingleView(doc, view, dependentTagCollector, consistencyService, prefix));
            }

            return result;
        }

        private IReadOnlyList<RebarListItem> CollectRebarItemsForTargets(
            Document doc,
            IReadOnlyList<SheetViewTarget> targets,
            RebarDependentTagCollector dependentTagCollector,
            RebarTagLeaderBendingDetailConsistencyService consistencyService)
        {
            if (targets == null || targets.Count == 0)
            {
                return new List<RebarListItem>();
            }

            var result = new List<RebarListItem>();
            foreach (var t in targets)
            {
                if (t == null || t.View == null)
                {
                    continue;
                }

                var prefix = BuildSheetViewPrefix(t.Sheet, t.View);
                result.AddRange(CollectRebarItemsForSingleView(doc, t.View, dependentTagCollector, consistencyService, prefix));
            }

            return result;
        }

        private IReadOnlyList<RebarListItem> CollectRebarItemsForSingleView(
            Document doc,
            View view,
            RebarDependentTagCollector dependentTagCollector,
            RebarTagLeaderBendingDetailConsistencyService consistencyService,
            string displayPrefix)
        {
            var rebars = CollectRebars(doc, view);
            var (mismatchRebarIds, leaderLineNotFoundRebarIds) = CollectIssueRebarIds(doc, view, rebars, consistencyService);
            var hostRebarIdsWithBendingDetail = CollectHostRebarIdsWithBendingDetail(doc, view, consistencyService);
            return BuildRebarItems(doc, view, rebars, mismatchRebarIds, leaderLineNotFoundRebarIds, dependentTagCollector, hostRebarIdsWithBendingDetail, displayPrefix);
        }

        private RebarTagCheckWindow BuildWindow(
            UIDocument uidoc,
            View activeView,
            IReadOnlyList<ViewSheet> sheets,
            int viewCount,
            RebarTagHighlightExternalEventService externalEventService,
            IReadOnlyList<RebarListItem> items,
            int rebarCount)
        {
            if (sheets == null || sheets.Count == 0)
            {
                return new RebarTagCheckWindow(uidoc, activeView, externalEventService, items, rebarCount);
            }

            return new RebarTagCheckWindow(uidoc, sheets, viewCount, externalEventService, items, rebarCount);
        }

        private IReadOnlyList<RebarListItem> BuildRebarItems(
            Document doc,
            View activeView,
            IReadOnlyList<Rebar> rebars,
            IReadOnlyCollection<ElementId> mismatchRebarIds,
            IReadOnlyCollection<ElementId> leaderLineNotFoundRebarIds,
            RebarDependentTagCollector dependentTagCollector,
            IReadOnlyCollection<ElementId> hostRebarIdsWithBendingDetail,
            string displayPrefix)
        {
            if (rebars == null) return new List<RebarListItem>();

            return rebars
                .Where(r => !ShouldHideRebar(doc, activeView, r, dependentTagCollector, hostRebarIdsWithBendingDetail))
                .Select(r => BuildRebarListItem(doc, activeView, r, mismatchRebarIds, leaderLineNotFoundRebarIds, dependentTagCollector, hostRebarIdsWithBendingDetail, displayPrefix))
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
            string displayPrefix)
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
                BuildDisplayText(displayPrefix, rebar),
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

        private string BuildDisplayText(string displayPrefix, Rebar rebar)
        {
            if (rebar == null) return string.Empty;

            var typeName = GetTypeName(rebar);
            var baseText = string.IsNullOrWhiteSpace(typeName)
                ? $"ID: {rebar.Id.Value}"
                : $"{typeName} / ID: {rebar.Id.Value}";

            if (string.IsNullOrWhiteSpace(displayPrefix))
            {
                return baseText;
            }

            return $"{displayPrefix} / {baseText}";
        }

        private IReadOnlyList<ViewSheet> CollectTargetSheets(Document doc)
        {
            // シート名にトークンが含まれるシートを対象にする
            if (doc == null)
            {
                return new List<ViewSheet>();
            }

            try
            {
                return new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Where(s => s != null && !s.IsTemplate && IsTargetSheetName(s.Name))
                    .OrderBy(s => s.SheetNumber ?? string.Empty)
                    .ThenBy(s => s.Name ?? string.Empty)
                    .ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ViewSheet>();
            }
        }

        private bool IsTargetSheetName(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return false;
            }

            return sheetName.Contains(TargetSheetNameToken);
        }

        private IReadOnlyList<SheetViewTarget> BuildTargetsForSheets(Document doc, IReadOnlyList<ViewSheet> sheets)
        {
            if (doc == null || sheets == null || sheets.Count == 0)
            {
                return new List<SheetViewTarget>();
            }

            var collector = new SheetPlacedViewCollector();
            var result = new List<SheetViewTarget>();
            foreach (var sheet in sheets)
            {
                if (sheet == null)
                {
                    continue;
                }

                var views = collector.Collect(doc, sheet);
                foreach (var v in views)
                {
                    if (v == null)
                    {
                        continue;
                    }

                    result.Add(new SheetViewTarget(sheet, v));
                }
            }

            return result;
        }

        private int CountViews(IReadOnlyList<SheetViewTarget> targets)
        {
            if (targets == null || targets.Count == 0)
            {
                return 0;
            }

            return targets
                .Where(t => t != null && t.View != null && t.View.Id != null && t.View.Id != ElementId.InvalidElementId)
                .Select(t => t.View.Id)
                .Distinct()
                .Count();
        }

        private string BuildSheetViewPrefix(ViewSheet sheet, View view)
        {
            var sheetText = BuildSheetDisplayText(sheet);
            var viewName = view == null ? string.Empty : (view.Name ?? string.Empty);

            if (string.IsNullOrWhiteSpace(sheetText))
            {
                return viewName ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(viewName))
            {
                return sheetText;
            }

            return $"{sheetText} / {viewName}";
        }

        private string BuildSheetDisplayText(ViewSheet sheet)
        {
            if (sheet == null)
            {
                return string.Empty;
            }

            var no = sheet.SheetNumber ?? string.Empty;
            var name = sheet.Name ?? string.Empty;
            if (string.IsNullOrWhiteSpace(no))
            {
                return name;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                return no;
            }

            return $"{no} {name}";
        }

        private class SheetViewTarget
        {
            public ViewSheet Sheet { get; }
            public View View { get; }

            public SheetViewTarget(ViewSheet sheet, View view)
            {
                Sheet = sheet;
                View = view;
            }
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

