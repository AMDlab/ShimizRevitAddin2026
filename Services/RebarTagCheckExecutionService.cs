using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using ShimizRevitAddin2026.UI.Models;

namespace ShimizRevitAddin2026.Services
{
    internal class RebarTagCheckExecutionService
    {
        private readonly RebarDependentTagCollector _dependentTagCollector;
        private readonly RebarTagLeaderBendingDetailConsistencyService _consistencyService;
        private readonly SheetPlacedViewCollector _sheetPlacedViewCollector;
        private readonly RebarCollectorInView _rebarCollectorInView;

        public RebarTagCheckExecutionService(
            RebarDependentTagCollector dependentTagCollector,
            RebarTagLeaderBendingDetailConsistencyService consistencyService)
        {
            _dependentTagCollector = dependentTagCollector;
            _consistencyService = consistencyService;
            _sheetPlacedViewCollector = new SheetPlacedViewCollector();
            _rebarCollectorInView = new RebarCollectorInView();
        }

        public RebarTagCheckExecutionResult ExecuteForActiveContext(Document doc, View activeView)
        {
            try
            {
                if (doc == null || activeView == null) return BuildFailed("doc/view が null です。");

                if (activeView is ViewSheet sheet)
                {
                    var placedViews = CollectPlacedViewsSafe(doc, sheet);
                    var targets = BuildTargetsForSheets(doc, new List<ViewSheet> { sheet });
                    var items = CollectRebarItemsForTargets(doc, targets);
                    var rebarCount = GetRebarCount(items);

                    return new RebarTagCheckExecutionResult(
                        true, string.Empty,
                        RebarTagCheckExecutionTargetKind.SheetWithPlacedViews,
                        null, sheet, placedViews, new List<ViewSheet>(), 0, items, rebarCount);
                }

                var viewItems = CollectRebarItemsForSingleView(doc, activeView, string.Empty);
                var viewRebarCount = GetRebarCount(viewItems);
                return new RebarTagCheckExecutionResult(
                    true, string.Empty,
                    RebarTagCheckExecutionTargetKind.SingleView,
                    activeView, null, new List<View>(), new List<ViewSheet>(), 0, viewItems, viewRebarCount);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return BuildFailed(ex.ToString());
            }
        }

        public RebarTagCheckExecutionResult ExecuteForKeywordSheets(Document doc, string keyword)
        {
            try
            {
                if (doc == null) return BuildFailed("doc が null です。");
                var k = (keyword ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(k)) return BuildFailed("Keyword が空です。");

                var sheets = CollectTargetSheetsByToken(doc, k);
                if (sheets.Count == 0) return BuildFailed($"[{k}] を含むシートが見つかりません。");

                var targets = BuildTargetsForSheets(doc, sheets);
                var items = CollectRebarItemsForTargets(doc, targets);
                var rebarCount = GetRebarCount(items);
                var viewCount = CountViews(targets);

                return new RebarTagCheckExecutionResult(
                    true, string.Empty,
                    RebarTagCheckExecutionTargetKind.MultipleSheets,
                    null, null, new List<View>(), sheets, viewCount, items, rebarCount);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return BuildFailed(ex.ToString());
            }
        }

        private RebarTagCheckExecutionResult BuildFailed(string message)
        {
            return new RebarTagCheckExecutionResult(
                false, message ?? string.Empty,
                RebarTagCheckExecutionTargetKind.None,
                null, null, new List<View>(), new List<ViewSheet>(), 0,
                new List<RebarListItem>(), 0);
        }

        private IReadOnlyList<View> CollectPlacedViewsSafe(Document doc, ViewSheet sheet)
        {
            try { return _sheetPlacedViewCollector.Collect(doc, sheet); }
            catch (Exception ex) { Debug.WriteLine(ex); return new List<View>(); }
        }

        private IReadOnlyList<Rebar> CollectRebars(Document doc, View view)
        {
            return _rebarCollectorInView.Collect(doc, view);
        }

        private IReadOnlyList<RebarListItem> CollectRebarItemsForTargets(
            Document doc, IReadOnlyList<SheetViewTarget> targets)
        {
            if (targets == null || targets.Count == 0) return new List<RebarListItem>();

            var result = new List<RebarListItem>();
            foreach (var t in targets)
            {
                if (t == null || t.View == null) continue;
                var prefix = BuildSheetViewPrefix(t.Sheet, t.View);
                result.AddRange(CollectRebarItemsForSingleView(doc, t.View, prefix));
            }
            return result;
        }

        private IReadOnlyList<RebarListItem> CollectRebarItemsForSingleView(
            Document doc, View view, string displayPrefix)
        {
            var rebars = CollectRebars(doc, view);
            var hostRebarIdsWithBendingDetail = CollectHostRebarIdsWithBendingDetail(doc, view);
            var (group2Ids, group3MatchIds, group3MismatchIds, group4MatchIds, group4MismatchIds) =
                CollectGroupedRebarIds(doc, view, rebars);

            return BuildRebarItems(
                doc, view, rebars,
                hostRebarIdsWithBendingDetail,
                group2Ids, group3MatchIds, group3MismatchIds,
                group4MatchIds, group4MismatchIds,
                displayPrefix);
        }

        private IReadOnlyList<RebarListItem> BuildRebarItems(
            Document doc,
            View activeView,
            IReadOnlyList<Rebar> rebars,
            IReadOnlyCollection<ElementId> hostRebarIdsWithBendingDetail,
            IReadOnlyCollection<ElementId> group2Ids,
            IReadOnlyCollection<ElementId> group3MatchIds,
            IReadOnlyCollection<ElementId> group3MismatchIds,
            IReadOnlyCollection<ElementId> group4MatchIds,
            IReadOnlyCollection<ElementId> group4MismatchIds,
            string displayPrefix)
        {
            if (rebars == null) return new List<RebarListItem>();

            return rebars
                .Select(r => BuildRebarListItem(
                    activeView, r,
                    hostRebarIdsWithBendingDetail,
                    group2Ids, group3MatchIds, group3MismatchIds,
                    group4MatchIds, group4MismatchIds,
                    displayPrefix))
                .ToList();
        }

        private RebarListItem BuildRebarListItem(
            View activeView,
            Rebar rebar,
            IReadOnlyCollection<ElementId> hostRebarIdsWithBendingDetail,
            IReadOnlyCollection<ElementId> group2Ids,
            IReadOnlyCollection<ElementId> group3MatchIds,
            IReadOnlyCollection<ElementId> group3MismatchIds,
            IReadOnlyCollection<ElementId> group4MatchIds,
            IReadOnlyCollection<ElementId> group4MismatchIds,
            string displayPrefix)
        {
            if (rebar == null)
                return new RebarListItem(ElementId.InvalidElementId, ElementId.InvalidElementId, string.Empty);

            var isNoTagOrBendingDetail = IsNoTagOrBendingDetail(activeView, rebar, hostRebarIdsWithBendingDetail);
            var isFreeEndTagNotFound = IsInCollection(group2Ids, rebar.Id);
            var isLeaderPointingRebarMismatch = IsInCollection(group3MismatchIds, rebar.Id);
            var isLeaderPointingRebar = isLeaderPointingRebarMismatch || IsInCollection(group3MatchIds, rebar.Id);
            var isLeaderPointingBendingDetailMismatch = IsInCollection(group4MismatchIds, rebar.Id);
            var isLeaderPointingBendingDetail = isLeaderPointingBendingDetailMismatch || IsInCollection(group4MatchIds, rebar.Id);

            return new RebarListItem(
                rebar.Id,
                activeView?.Id ?? ElementId.InvalidElementId,
                BuildDisplayText(displayPrefix, rebar),
                isNoTagOrBendingDetail,
                isFreeEndTagNotFound,
                isLeaderPointingRebar,
                isLeaderPointingRebarMismatch,
                isLeaderPointingBendingDetail,
                isLeaderPointingBendingDetailMismatch);
        }

        /// <summary>① タグまたは曲げの詳細がホストされてない（鉄筋モデルに曲げ詳細・タグが存在しない）</summary>
        private bool IsNoTagOrBendingDetail(
            View activeView,
            Rebar rebar,
            IReadOnlyCollection<ElementId> hostRebarIdsWithBendingDetail)
        {
            if (rebar == null || rebar.Id == null || rebar.Id == ElementId.InvalidElementId)
                return false;

            var hasStructureTag = HasStructureRebarTag(activeView, rebar);
            if (hasStructureTag) return false;

            return !HasBendingDetail(hostRebarIdsWithBendingDetail, rebar.Id);
        }

        private bool HasBendingDetail(IReadOnlyCollection<ElementId> hostRebarIdsWithBendingDetail, ElementId rebarId)
        {
            if (rebarId == null || rebarId == ElementId.InvalidElementId) return false;
            if (hostRebarIdsWithBendingDetail == null || hostRebarIdsWithBendingDetail.Count == 0) return false;
            return hostRebarIdsWithBendingDetail.Contains(rebarId);
        }

        private bool HasStructureRebarTag(View activeView, Rebar rebar)
        {
            try
            {
                if (activeView == null || rebar == null || _dependentTagCollector == null) return false;
                var doc = rebar.Document;
                if (doc == null) return false;
                var model = _dependentTagCollector.Collect(doc, rebar, activeView);
                return model != null && model.StructureTagIds != null && model.StructureTagIds.Count > 0;
            }
            catch (Exception ex) { Debug.WriteLine(ex); return false; }
        }

        private bool IsInCollection(IReadOnlyCollection<ElementId> ids, ElementId id)
        {
            if (ids == null || ids.Count == 0) return false;
            if (id == null || id == ElementId.InvalidElementId) return false;
            return ids.Contains(id);
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
            return string.IsNullOrWhiteSpace(displayPrefix) ? baseText : $"{displayPrefix} / {baseText}";
        }

        private string GetTypeName(Rebar rebar)
        {
            try
            {
                if (rebar == null) return string.Empty;
                var doc = rebar.Document;
                if (doc == null) return string.Empty;
                var typeId = rebar.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId) return string.Empty;
                var type = doc.GetElement(typeId) as ElementType;
                return type == null ? string.Empty : (type.Name ?? string.Empty);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return string.Empty; }
        }

        private (IReadOnlyCollection<ElementId> group2Ids,
                 IReadOnlyCollection<ElementId> group3MatchIds,
                 IReadOnlyCollection<ElementId> group3MismatchIds,
                 IReadOnlyCollection<ElementId> group4MatchIds,
                 IReadOnlyCollection<ElementId> group4MismatchIds)
            CollectGroupedRebarIds(Document doc, View activeView, IReadOnlyList<Rebar> rebars)
        {
            var empty = (
                (IReadOnlyCollection<ElementId>)new List<ElementId>(),
                (IReadOnlyCollection<ElementId>)new List<ElementId>(),
                (IReadOnlyCollection<ElementId>)new List<ElementId>(),
                (IReadOnlyCollection<ElementId>)new List<ElementId>(),
                (IReadOnlyCollection<ElementId>)new List<ElementId>());

            try
            {
                if (doc == null || activeView == null || rebars == null || _consistencyService == null)
                    return empty;
                return _consistencyService.CollectGroupedRebarIds(doc, rebars, activeView);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return empty; }
        }

        private IReadOnlyCollection<ElementId> CollectHostRebarIdsWithBendingDetail(Document doc, View activeView)
        {
            try
            {
                if (doc == null || activeView == null || _consistencyService == null)
                    return new List<ElementId>();
                return _consistencyService.CollectHostRebarIdsWithBendingDetail(doc, activeView);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return new List<ElementId>(); }
        }

        private IReadOnlyList<ViewSheet> CollectTargetSheetsByToken(Document doc, string token)
        {
            if (doc == null) return new List<ViewSheet>();
            try
            {
                return new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Where(s => s != null && !s.IsTemplate && IsTargetSheet(s, token))
                    .OrderBy(s => s.SheetNumber ?? string.Empty)
                    .ThenBy(s => s.Name ?? string.Empty)
                    .ToList();
            }
            catch (Exception ex) { Debug.WriteLine(ex); return new List<ViewSheet>(); }
        }

        private bool IsTargetSheet(ViewSheet sheet, string token)
        {
            if (sheet == null || string.IsNullOrWhiteSpace(token)) return false;
            var sheetNo = sheet.SheetNumber ?? string.Empty;
            var sheetName = sheet.Name ?? string.Empty;
            return ContainsToken(sheetNo, token) || ContainsToken(sheetName, token);
        }

        private bool ContainsToken(string text, string token)
        {
            if (string.IsNullOrWhiteSpace(token) || string.IsNullOrWhiteSpace(text)) return false;
            return text.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private IReadOnlyList<SheetViewTarget> BuildTargetsForSheets(
            Document doc, IReadOnlyList<ViewSheet> sheets)
        {
            if (doc == null || sheets == null || sheets.Count == 0) return new List<SheetViewTarget>();
            var result = new List<SheetViewTarget>();
            foreach (var sheet in sheets)
            {
                if (sheet == null) continue;
                var views = _sheetPlacedViewCollector.Collect(doc, sheet);
                foreach (var v in views)
                {
                    if (v != null) result.Add(new SheetViewTarget(sheet, v));
                }
            }
            return result;
        }

        private int CountViews(IReadOnlyList<SheetViewTarget> targets)
        {
            if (targets == null || targets.Count == 0) return 0;
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
            if (string.IsNullOrWhiteSpace(sheetText)) return viewName ?? string.Empty;
            if (string.IsNullOrWhiteSpace(viewName)) return sheetText;
            return $"{sheetText} / {viewName}";
        }

        private string BuildSheetDisplayText(ViewSheet sheet)
        {
            if (sheet == null) return string.Empty;
            var no = sheet.SheetNumber ?? string.Empty;
            var name = sheet.Name ?? string.Empty;
            if (string.IsNullOrWhiteSpace(no)) return name;
            if (string.IsNullOrWhiteSpace(name)) return no;
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
    }
}
