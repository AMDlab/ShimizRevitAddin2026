using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using ShimizRevitAddin2026.Model;

namespace ShimizRevitAddin2026.Services
{
    internal class RebarTagLeaderBendingDetailConsistencyService
    {
        private readonly RebarDependentTagCollector _dependentTagCollector;
        private const double Eps = 1e-9;

        public RebarTagLeaderBendingDetailConsistencyService(RebarDependentTagCollector dependentTagCollector)
        {
            _dependentTagCollector = dependentTagCollector;
        }

        public IReadOnlyList<RebarTagLeaderBendingDetailCheckItem> Check(Document doc, Rebar rebar, View view)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (rebar == null) throw new ArgumentNullException(nameof(rebar));
            if (view == null) throw new ArgumentNullException(nameof(view));

            var bendingDetails = CollectBendingDetailElementsInView(doc, view);
            var allRebars = CollectAllRebarsInView(doc, view);
            var item = BuildResultForRebar(doc, rebar, view, bendingDetails, allRebars);
            return item == null
                ? new List<RebarTagLeaderBendingDetailCheckItem>()
                : new List<RebarTagLeaderBendingDetailCheckItem> { item };
        }

        /// <summary>
        /// ビュー内の全鉄筋をグループ②③④に分類して返す。
        /// ① はExecutionService側で IsNoTagOrBendingDetail として判定する。
        /// </summary>
        public (
            IReadOnlyCollection<ElementId> group2Ids,
            IReadOnlyCollection<ElementId> group3MatchIds,
            IReadOnlyCollection<ElementId> group3MismatchIds,
            IReadOnlyCollection<ElementId> group4MatchIds,
            IReadOnlyCollection<ElementId> group4MismatchIds)
            CollectGroupedRebarIds(Document doc, IReadOnlyList<Rebar> rebars, View view)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));

            var empty = (
                (IReadOnlyCollection<ElementId>)new List<ElementId>(),
                (IReadOnlyCollection<ElementId>)new List<ElementId>(),
                (IReadOnlyCollection<ElementId>)new List<ElementId>(),
                (IReadOnlyCollection<ElementId>)new List<ElementId>(),
                (IReadOnlyCollection<ElementId>)new List<ElementId>());

            if (rebars == null || rebars.Count == 0) return empty;

            var bendingDetails = CollectBendingDetailElementsInView(doc, view);
            return CollectGroupedRebarIdsInternal(doc, rebars, view, bendingDetails);
        }

        public IReadOnlyCollection<ElementId> CollectHostRebarIdsWithBendingDetail(Document doc, View view)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));

            var bendingDetails = CollectBendingDetailElementsInView(doc, view);
            return CollectHostRebarIdsWithBendingDetailInternal(doc, bendingDetails);
        }

        // -------------------------------------------------------------------------
        // internal classification
        // -------------------------------------------------------------------------

        private (
            IReadOnlyCollection<ElementId> group2Ids,
            IReadOnlyCollection<ElementId> group3MatchIds,
            IReadOnlyCollection<ElementId> group3MismatchIds,
            IReadOnlyCollection<ElementId> group4MatchIds,
            IReadOnlyCollection<ElementId> group4MismatchIds)
            CollectGroupedRebarIdsInternal(
                Document doc,
                IReadOnlyList<Rebar> rebars,
                View view,
                IReadOnlyList<Element> bendingDetails)
        {
            var group2 = new HashSet<ElementId>();
            var group3Match = new HashSet<ElementId>();
            var group3Mismatch = new HashSet<ElementId>();
            var group4Match = new HashSet<ElementId>();
            var group4Mismatch = new HashSet<ElementId>();

            foreach (var rebar in rebars)
            {
                if (rebar == null) continue;

                var item = BuildResultForRebar(doc, rebar, view, bendingDetails, rebars);
                if (item == null) continue;

                if (item.IsLeaderPointingRebarDirectly)
                {
                    if (item.IsMatch) group3Match.Add(rebar.Id);
                    else group3Mismatch.Add(rebar.Id);
                    continue;
                }

                var hasBendingDetail = item.PointedBendingDetailId != null
                    && item.PointedBendingDetailId != ElementId.InvalidElementId;

                if (hasBendingDetail)
                {
                    if (item.IsMatch) group4Match.Add(rebar.Id);
                    else group4Mismatch.Add(rebar.Id);
                    continue;
                }

                // 引き出し線の交点が取得できない → グループ②
                group2.Add(rebar.Id);
            }

            return (
                group2.ToList(),
                group3Match.ToList(),
                group3Mismatch.ToList(),
                group4Match.ToList(),
                group4Mismatch.ToList());
        }

        // -------------------------------------------------------------------------
        // bending detail host collection
        // -------------------------------------------------------------------------

        private IReadOnlyCollection<ElementId> CollectHostRebarIdsWithBendingDetailInternal(
            Document doc,
            IReadOnlyList<Element> bendingDetails)
        {
            try
            {
                var result = new HashSet<ElementId>();
                if (doc == null || bendingDetails == null || bendingDetails.Count == 0)
                    return result.ToList();

                foreach (var bd in bendingDetails)
                {
                    if (bd == null) continue;
                    var (ok, hostRebarId) = TryResolveHostRebarIdFromBendingDetail(doc, bd);
                    if (!ok || hostRebarId == null || hostRebarId == ElementId.InvalidElementId) continue;
                    result.Add(hostRebarId);
                }

                return result.ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ElementId>();
            }
        }

        // -------------------------------------------------------------------------
        // core check per rebar
        // -------------------------------------------------------------------------

        private RebarTagLeaderBendingDetailCheckItem BuildResultForRebar(
            Document doc,
            Rebar rebar,
            View view,
            IReadOnlyList<Element> bendingDetails,
            IReadOnlyList<Rebar> allRebarsInView)
        {
            var stage = "Init";
            try
            {
                stage = "CollectStructureRebarTagIds";
                var tagIds = CollectStructureRebarTagIds(doc, rebar, view);
                if (tagIds == null || tagIds.Count == 0)
                {
                    return CreateFail(ElementId.InvalidElementId, rebar?.Id ?? ElementId.InvalidElementId, null,
                        BuildStageMessage(stage, "構造鉄筋タグがありません。"));
                }

                stage = "ResolveIndependentTags";
                var tags = ResolveIndependentTags(doc, tagIds);
                if (tags == null || tags.Count == 0)
                {
                    return CreateFail(tagIds.FirstOrDefault() ?? ElementId.InvalidElementId,
                        rebar?.Id ?? ElementId.InvalidElementId, null,
                        BuildStageMessage(stage, "構造鉄筋タグ要素を取得できません。"));
                }

                stage = "TrySelectFreeEndTagWithLeaderPoints";
                var (freeTag, segStart0, segEnd0, selectReason) = TrySelectFreeEndTagWithLeaderPoints(tags, view);
                if (freeTag == null)
                {
                    return CreateFail(tagIds.FirstOrDefault() ?? ElementId.InvalidElementId,
                        rebar?.Id ?? ElementId.InvalidElementId, null,
                        BuildStageMessage(stage, selectReason));
                }

                stage = "TryResolveTaggedRebarId";
                var (hasTaggedRebar, taggedRebarId) = TryResolveTaggedRebarId(doc, freeTag);
                if (!hasTaggedRebar)
                {
                    return CreateFail(freeTag.Id, ElementId.InvalidElementId, null,
                        BuildStageMessage(stage, "自由端タグに紐づく鉄筋を取得できません。"));
                }

                var segStart = segStart0;
                var segEnd = segEnd0;

                stage = "TryBuildRayFromLastSegment";
                var (hasRay, rayOrigin, rayDir, rayReason) = TryBuildRayFromLastSegment(segStart, segEnd);
                if (!hasRay)
                {
                    return CreateFail(freeTag.Id, taggedRebarId, segEnd,
                        BuildStageMessage(stage, rayReason));
                }

                // -----------------------------------------------------------------
                // 距離比較：鉄筋モデル直接 vs 曲げ詳細
                // -----------------------------------------------------------------
                stage = "TryFindFirstIntersectionWithBendingDetail";
                var (hasBendingHit, hitDetail, hitBendingPoint, tBending, bendingReason) =
                    TryFindFirstIntersectionWithBendingDetail(bendingDetails, view, rayOrigin, rayDir);

                stage = "TryFindFirstIntersectionWithRebar";
                var (hasRebarHit, hitRebar, hitRebarPoint, tRebar, rebarReason) =
                    TryFindFirstIntersectionWithRebar(allRebarsInView, view, rayOrigin, rayDir);

                // グループ③：鉄筋の方が近い（または曲げ詳細がない）
                if (hasRebarHit && (!hasBendingHit || tRebar <= tBending))
                {
                    var isMatch = hitRebar.Id == taggedRebarId;
                    return new RebarTagLeaderBendingDetailCheckItem(
                        freeTag.Id,
                        taggedRebarId,
                        hitRebarPoint,
                        ElementId.InvalidElementId,
                        ElementId.InvalidElementId,
                        isMatch,
                        isMatch
                            ? "鉄筋モデル直接一致。"
                            : $"鉄筋モデル直接不一致（指先={hitRebar.Id.Value} / タグ対象={taggedRebarId.Value}）。",
                        isLeaderPointingRebarDirectly: true,
                        pointedDirectRebarId: hitRebar.Id);
                }

                // グループ④：曲げ詳細の方が近い（または鉄筋がない）
                if (hasBendingHit && hitDetail != null)
                {
                    stage = "TryResolveHostRebarIdFromBendingDetail";
                    var (hasHost, hostRebarId) = TryResolveHostRebarIdFromBendingDetail(doc, hitDetail);
                    if (!hasHost)
                    {
                        return new RebarTagLeaderBendingDetailCheckItem(
                            freeTag.Id,
                            taggedRebarId,
                            hitBendingPoint,
                            hitDetail.Id,
                            ElementId.InvalidElementId,
                            false,
                            BuildStageMessage(stage, "曲げ詳細から鉄筋を特定できません。"));
                    }

                    if (hostRebarId == taggedRebarId)
                    {
                        return new RebarTagLeaderBendingDetailCheckItem(
                            freeTag.Id,
                            taggedRebarId,
                            hitBendingPoint,
                            hitDetail.Id,
                            hostRebarId,
                            true,
                            "一致。");
                    }

                    return new RebarTagLeaderBendingDetailCheckItem(
                        freeTag.Id,
                        taggedRebarId,
                        hitBendingPoint,
                        hitDetail.Id,
                        hostRebarId,
                        false,
                        $"不一致（曲げ詳細Host={hostRebarId.Value} / タグ対象鉄筋={taggedRebarId.Value}）。");
                }

                // グループ②：どちらも見つからない
                var noHitReason = BuildNoHitReason(hasBendingHit, bendingReason, hasRebarHit, rebarReason);
                return CreateFail(freeTag.Id, taggedRebarId, segEnd,
                    BuildStageMessage("TryFindIntersection", noHitReason));
            }
            catch (Autodesk.Revit.Exceptions.InternalException iex)
            {
                Debug.WriteLine(iex);
                var rebarIdText = rebar?.Id == null ? string.Empty : rebar.Id.Value.ToString();
                return CreateFail(
                    rebar?.Id ?? ElementId.InvalidElementId,
                    rebar?.Id ?? ElementId.InvalidElementId,
                    null,
                    $"{stage}: InternalException: {iex.Message} (RebarId={rebarIdText})");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return CreateFail(
                    rebar?.Id ?? ElementId.InvalidElementId,
                    rebar?.Id ?? ElementId.InvalidElementId,
                    null,
                    $"{stage}: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private string BuildNoHitReason(bool hasBendingHit, string bendingReason, bool hasRebarHit, string rebarReason)
        {
            var parts = new List<string>();
            if (!hasBendingHit && !string.IsNullOrWhiteSpace(bendingReason))
                parts.Add($"曲げ詳細: {bendingReason}");
            if (!hasRebarHit && !string.IsNullOrWhiteSpace(rebarReason))
                parts.Add($"鉄筋: {rebarReason}");
            if (parts.Count == 0)
                return "引き出し線の先に曲げ詳細も鉄筋モデルも見つかりません。";
            return string.Join(" / ", parts);
        }

        // -------------------------------------------------------------------------
        // rebar intersection (new)
        // -------------------------------------------------------------------------

        private (bool ok, Rebar rebar, XYZ hitPoint, double t, string reason)
            TryFindFirstIntersectionWithRebar(
                IReadOnlyList<Rebar> rebars,
                View view,
                XYZ rayOrigin,
                XYZ rayDir)
        {
            try
            {
                if (rebars == null || rebars.Count == 0)
                    return (false, null, null, 0, "ビュー内に鉄筋が見つかりません。");

                var (hasBasis, right, up, basisReason) = TryGetViewBasis(view);
                if (!hasBasis)
                    return (false, null, null, 0, string.IsNullOrWhiteSpace(basisReason) ? "ビュー座標系を取得できません。" : basisReason);

                if (rayOrigin == null || rayDir == null)
                    return (false, null, null, 0, "レイが無効です。");

                var (okO, uo, vo) = TryProjectToUv(rayOrigin, right, up);
                var (okD, du, dv) = TryProjectDirToUv(rayDir, right, up);
                if (!okO || !okD)
                    return (false, null, null, 0, "レイの投影に失敗しました。");

                Rebar best = null;
                XYZ bestPoint = null;
                var bestT = double.MaxValue;
                var errors = new List<string>();

                foreach (var rebar in rebars)
                {
                    if (rebar == null) continue;
                    try
                    {
                        var (ok, t, reason) = TryIntersectRayWithBoundingBox2D(rebar, view, uo, vo, du, dv, right, up);
                        if (!ok)
                        {
                            if (!string.IsNullOrWhiteSpace(reason))
                                errors.Add($"{rebar.Id.Value}: {reason}");
                            continue;
                        }

                        if (t < bestT)
                        {
                            bestT = t;
                            best = rebar;
                            bestPoint = rayOrigin + (rayDir * t);
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.InternalException iex)
                    {
                        errors.Add($"{rebar.Id.Value}: InternalException: {iex.Message}");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{rebar.Id.Value}: {ex.GetType().Name}: {ex.Message}");
                    }
                }

                if (best == null)
                {
                    var msg = errors.Count == 0
                        ? "引き出し線の延長と鉄筋の交点を取得できません。"
                        : "鉄筋との交点取得で例外が発生しました。\n" + string.Join("\n", errors.Distinct().Take(8).ToList());
                    return (false, null, null, 0, msg);
                }

                return (true, best, bestPoint, bestT, string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null, null, 0, $"TryFindFirstIntersectionWithRebar: {ex.GetType().Name}: {ex.Message}");
            }
        }

        // -------------------------------------------------------------------------
        // bending detail intersection (t を追加)
        // -------------------------------------------------------------------------

        private (bool ok, Element bendingDetail, XYZ hitPoint, double t, string reason)
            TryFindFirstIntersectionWithBendingDetail(
                IReadOnlyList<Element> bendingDetails,
                View view,
                XYZ rayOrigin,
                XYZ rayDir)
        {
            try
            {
                if (bendingDetails == null || bendingDetails.Count == 0)
                    return (false, null, null, 0, "曲げ詳細が見つかりません。");

                var (hasBasis, right, up, basisReason) = TryGetViewBasis(view);
                if (!hasBasis)
                    return (false, null, null, 0, string.IsNullOrWhiteSpace(basisReason) ? "ビュー座標系を取得できません。" : basisReason);

                if (rayOrigin == null || rayDir == null)
                    return (false, null, null, 0, "レイが無効です。");

                var (okO, uo, vo) = TryProjectToUv(rayOrigin, right, up);
                var (okD, du, dv) = TryProjectDirToUv(rayDir, right, up);
                if (!okO || !okD)
                    return (false, null, null, 0, "レイの投影に失敗しました。");

                Element best = null;
                XYZ bestPoint = null;
                var bestT = double.MaxValue;
                var errors = new List<string>();

                foreach (var e in bendingDetails)
                {
                    try
                    {
                        var (ok, t, reason) = TryIntersectRayWithBoundingBox2D(e, view, uo, vo, du, dv, right, up);
                        if (!ok)
                        {
                            if (!string.IsNullOrWhiteSpace(reason) && e != null)
                                errors.Add($"{e.Id.Value}: {reason}");
                            continue;
                        }

                        if (t < bestT)
                        {
                            bestT = t;
                            best = e;
                            bestPoint = rayOrigin + (rayDir * t);
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.InternalException iex)
                    {
                        if (e != null) errors.Add($"{e.Id.Value}: InternalException: {iex.Message}");
                        else errors.Add($"InternalException: {iex.Message}");
                    }
                    catch (Exception ex)
                    {
                        if (e != null) errors.Add($"{e.Id.Value}: {ex.GetType().Name}: {ex.Message}");
                        else errors.Add($"{ex.GetType().Name}: {ex.Message}");
                    }
                }

                if (best == null)
                {
                    if (errors.Count == 0)
                        return (false, null, null, 0, "引き出し線の延長と曲げ詳細の交点を取得できません。");

                    var unique = errors.Distinct().Take(8).ToList();
                    return (false, null, null, 0, "曲げ詳細との交点取得で例外が発生しました。\n" + string.Join("\n", unique));
                }

                return (true, best, bestPoint, bestT, string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null, null, 0, $"TryFindFirstIntersectionWithBendingDetail: {ex.GetType().Name}: {ex.Message}");
            }
        }

        // -------------------------------------------------------------------------
        // rebar collection helpers
        // -------------------------------------------------------------------------

        private IReadOnlyList<Rebar> CollectAllRebarsInView(Document doc, View view)
        {
            try
            {
                if (doc == null || view == null) return new List<Rebar>();
                return new FilteredElementCollector(doc, view.Id)
                    .OfClass(typeof(Rebar))
                    .Cast<Rebar>()
                    .ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<Rebar>();
            }
        }

        // -------------------------------------------------------------------------
        // bending detail collection
        // -------------------------------------------------------------------------

        private IReadOnlyList<Element> CollectBendingDetailElementsInView(Document doc, View view)
        {
            try
            {
                var collector = BuildBendingDetailCollector(doc, view);
                var all = collector.ToElements();
                return all.Where(IsBendingDetailElement).ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<Element>();
            }
        }

        private FilteredElementCollector BuildBendingDetailCollector(Document doc, View view)
        {
            try
            {
                var collector = new FilteredElementCollector(doc, view.Id);
                var (hasBic, bic) = TryGetBendingDetailBuiltInCategory();
                if (hasBic) collector = collector.OfCategory(bic);
                return collector.WhereElementIsNotElementType();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType();
            }
        }

        private (bool ok, BuiltInCategory bic) TryGetBendingDetailBuiltInCategory()
        {
            try
            {
                var candidates = new[] { "OST_RebarBendingDetails", "OST_RebarBendingDetail" };
                foreach (var name in candidates)
                {
                    if (Enum.TryParse(name, true, out BuiltInCategory bic))
                        return (true, bic);
                }
                return (false, default(BuiltInCategory));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, default(BuiltInCategory));
            }
        }

        private bool IsBendingDetailElement(Element e)
        {
            try
            {
                if (e == null) return false;
                if (IsBendingDetailCategory(e.Category)) return true;
                var (hasTypeHit, _) = TryMatchBendingDetailTypeElement(e);
                if (hasTypeHit) return true;
                var runtimeTypeName = e.GetType().Name ?? string.Empty;
                return runtimeTypeName.IndexOf("RebarBendingDetail", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return false;
            }
        }

        private bool IsBendingDetailCategory(Category category)
        {
            try
            {
                if (category == null) return false;
                var (hasBic, bic) = TryGetBendingDetailBuiltInCategory();
                if (hasBic && category.Id != null && category.Id.Value == (int)bic) return true;
                var name = category.Name ?? string.Empty;
                if (name.IndexOf("Bending", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    name.IndexOf("Detail", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                return name.IndexOf("曲げ", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return false;
            }
        }

        private (bool ok, Element typeElement) TryMatchBendingDetailTypeElement(Element e)
        {
            try
            {
                if (e == null) return (false, null);
                var doc = e.Document;
                if (doc == null) return (false, null);
                var typeId = e.GetTypeId();
                if (typeId == null || typeId == ElementId.InvalidElementId) return (false, null);
                var typeElement = doc.GetElement(typeId);
                if (typeElement == null) return (false, null);
                var typeName = typeElement.GetType().Name ?? string.Empty;
                return (typeName.IndexOf("RebarBendingDetailType", StringComparison.OrdinalIgnoreCase) >= 0, typeElement);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null);
            }
        }

        // -------------------------------------------------------------------------
        // tag helpers
        // -------------------------------------------------------------------------

        private IReadOnlyList<ElementId> CollectStructureRebarTagIds(Document doc, Rebar rebar, View view)
        {
            try
            {
                if (_dependentTagCollector == null) return new List<ElementId>();
                var model = _dependentTagCollector.Collect(doc, rebar, view);
                return model?.StructureTagIds?.ToList() ?? new List<ElementId>();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ElementId>();
            }
        }

        private IReadOnlyList<IndependentTag> ResolveIndependentTags(Document doc, IReadOnlyList<ElementId> tagIds)
        {
            try
            {
                if (doc == null || tagIds == null || tagIds.Count == 0) return new List<IndependentTag>();
                var result = new List<IndependentTag>();
                foreach (var id in tagIds)
                {
                    if (id == null || id == ElementId.InvalidElementId) continue;
                    var e = doc.GetElement(id) as IndependentTag;
                    if (e != null) result.Add(e);
                }
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<IndependentTag>();
            }
        }

        private (IndependentTag tag, XYZ start, XYZ end, string reason)
            TrySelectFreeEndTagWithLeaderPoints(IReadOnlyList<IndependentTag> tags, View view)
        {
            try
            {
                if (tags == null || tags.Count == 0)
                    return (null, null, null, "曲げ加工詳細タグがありません。");

                var reasons = new List<string>();
                foreach (var tag in tags)
                {
                    if (tag == null) continue;
                    try
                    {
                        var tagIdText = SafeElementIdText(tag);
                        var (ok, s, e, r) = TryGetLeaderLinePoints(tag, view);
                        if (ok) return (tag, s, e, string.Empty);

                        var cond = TryGetLeaderEndConditionText(tag);
                        reasons.Add(string.IsNullOrWhiteSpace(cond)
                            ? $"{tagIdText}: {r}"
                            : $"{tagIdText}: LeaderEndCondition={cond} / {r}");
                    }
                    catch (Autodesk.Revit.Exceptions.InternalException iex)
                    {
                        reasons.Add($"{SafeElementIdText(tag)}: InternalException: {iex.Message}");
                    }
                    catch (Exception ex)
                    {
                        reasons.Add($"{SafeElementIdText(tag)}: {ex.GetType().Name}: {ex.Message}");
                    }
                }

                if (reasons.Count == 0)
                    return (null, null, null, "自由端タグの候補がありません。");

                return (null, null, null, "自由端タグの線分を取得できません。\n" + string.Join("\n", reasons));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (null, null, null, $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        private (bool ok, XYZ start, XYZ end, string reason)
            TryGetLeaderLinePoints(IndependentTag tag, View view)
        {
            try
            {
                if (tag == null) return (false, null, null, "タグが null です。");

                var (refs, refReason) = TryGetTaggedReferencesWithReason(tag);
                if (refs.Count == 0)
                    return (false, null, null, string.IsNullOrWhiteSpace(refReason) ? "タグ参照を取得できません。" : refReason);

                var errors = new List<string>();
                foreach (var r in refs)
                {
                    if (r == null) continue;
                    XYZ end;
                    try { end = tag.GetLeaderEnd(r); }
                    catch (Autodesk.Revit.Exceptions.InternalException iex)
                    { errors.Add($"GetLeaderEnd: InternalException: {iex.Message}"); continue; }
                    catch (Exception ex)
                    { errors.Add($"GetLeaderEnd: {ex.GetType().Name}: {ex.Message}"); continue; }

                    if (end == null) { errors.Add("GetLeaderEnd が null を返しました。"); continue; }

                    var elbow = TryGetLeaderElbow(tag, r);
                    var start = elbow ?? TryGetTagHeadOrFallback(tag, view);
                    if (start == null) { errors.Add("始点を取得できません。"); continue; }

                    return (true, start, end, string.Empty);
                }

                if (errors.Count == 0)
                    return (false, null, null, "GetLeaderEnd に失敗しました。");
                return (false, null, null, string.Join(" / ", errors.Distinct().ToList()));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null, null, $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        private XYZ TryGetTagHeadOrFallback(IndependentTag tag, View view)
        {
            try
            {
                if (tag == null) return null;
                try { return tag.TagHeadPosition; }
                catch (Autodesk.Revit.Exceptions.InternalException) { }
                catch (Exception) { }

                try
                {
                    var bb = tag.get_BoundingBox(view);
                    if (bb == null || bb.Min == null || bb.Max == null) return null;
                    return (bb.Min + bb.Max) * 0.5;
                }
                catch (Exception ex) { Debug.WriteLine(ex); return null; }
            }
            catch (Exception ex) { Debug.WriteLine(ex); return null; }
        }

        private string TryGetLeaderEndConditionText(IndependentTag tag)
        {
            try
            {
                if (tag == null) return string.Empty;
                var p = tag.GetType().GetProperty("LeaderEndCondition");
                if (p == null) return string.Empty;
                var v = p.GetValue(tag, null);
                return v == null ? string.Empty : v.ToString();
            }
            catch (Exception ex) { Debug.WriteLine(ex); return string.Empty; }
        }

        private (IReadOnlyList<Reference> refs, string reason) TryGetTaggedReferencesWithReason(IndependentTag tag)
        {
            try
            {
                if (tag == null) return (new List<Reference>(), "タグが null です。");
                var m = tag.GetType().GetMethod("GetTaggedReferences", Type.EmptyTypes);
                if (m == null) return (new List<Reference>(), "GetTaggedReferences メソッドが見つかりません。");

                object obj;
                try { obj = m.Invoke(tag, null); }
                catch (System.Reflection.TargetInvocationException tex)
                {
                    var inner = tex.InnerException;
                    if (inner is Autodesk.Revit.Exceptions.InternalException iex)
                        return (new List<Reference>(), $"GetTaggedReferences: InternalException: {iex.Message}");
                    return (new List<Reference>(), $"GetTaggedReferences: {inner?.GetType().Name ?? tex.GetType().Name}: {inner?.Message ?? tex.Message}");
                }

                if (obj is IEnumerable<Reference> refs)
                    return (refs.Where(x => x != null).ToList(), string.Empty);

                return (new List<Reference>(), "GetTaggedReferences の戻り値が想定外です。");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (new List<Reference>(), $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        private XYZ TryGetLeaderElbow(IndependentTag tag, Reference r)
        {
            try
            {
                if (tag == null || r == null) return null;
                var m = tag.GetType().GetMethod("GetLeaderElbow", new[] { typeof(Reference) });
                if (m == null) return null;
                var obj = m.Invoke(tag, new object[] { r });
                return obj as XYZ;
            }
            catch (Exception ex) { Debug.WriteLine(ex); return null; }
        }

        private (bool ok, ElementId rebarId) TryResolveTaggedRebarId(Document doc, IndependentTag tag)
        {
            try
            {
                if (doc == null || tag == null) return (false, ElementId.InvalidElementId);
                var ids = tag.GetTaggedLocalElementIds();
                if (ids == null || ids.Count == 0) return (false, ElementId.InvalidElementId);
                foreach (var id in ids)
                {
                    if (id == null || id == ElementId.InvalidElementId) continue;
                    if (doc.GetElement(id) is Rebar) return (true, id);
                }
                return (false, ElementId.InvalidElementId);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, ElementId.InvalidElementId); }
        }

        // -------------------------------------------------------------------------
        // ray helpers
        // -------------------------------------------------------------------------

        private (bool ok, XYZ origin, XYZ dir, string reason)
            TryBuildRayFromLastSegment(XYZ start, XYZ end)
        {
            try
            {
                if (start == null || end == null) return (false, null, null, "線分が無効です。");
                var v = end - start;
                var len = v.GetLength();
                if (len < Eps) return (false, null, null, "線分の長さが 0 です。");
                return (true, end, v / len, string.Empty);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, null, null, $"{ex.GetType().Name}: {ex.Message}"); }
        }

        private (bool ok, XYZ right, XYZ up, string reason) TryGetViewBasis(View view)
        {
            try
            {
                if (view == null) return (false, null, null, "View が null です。");
                var right = view.RightDirection;
                var up = view.UpDirection;
                if (right == null || up == null) return (false, null, null, "RightDirection/UpDirection を取得できません。");
                return (true, right, up, string.Empty);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, null, null, $"{ex.GetType().Name}: {ex.Message}"); }
        }

        private (bool ok, double u, double v) TryProjectToUv(XYZ p, XYZ right, XYZ up)
        {
            try
            {
                if (p == null || right == null || up == null) return (false, 0, 0);
                return (true, p.DotProduct(right), p.DotProduct(up));
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, 0, 0); }
        }

        private (bool ok, double du, double dv) TryProjectDirToUv(XYZ d, XYZ right, XYZ up)
        {
            try
            {
                if (d == null || right == null || up == null) return (false, 0, 0);
                var du = d.DotProduct(right);
                var dv = d.DotProduct(up);
                if (Math.Abs(du) < Eps && Math.Abs(dv) < Eps) return (false, 0, 0);
                return (true, du, dv);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, 0, 0); }
        }

        private (bool ok, double tEnter, string reason) TryIntersectRayWithBoundingBox2D(
            Element element, View view,
            double uo, double vo, double du, double dv,
            XYZ right, XYZ up)
        {
            try
            {
                if (element == null) return (false, 0, "Element が null です。");

                var (hasBb, bb, bbReason) = TryGetBoundingBoxPreferView(element, view);
                if (!hasBb || bb == null || bb.Min == null || bb.Max == null)
                    return (false, 0, string.IsNullOrWhiteSpace(bbReason) ? "BoundingBox を取得できません。" : bbReason);

                var corners = BuildBoundingBoxCorners(bb.Min, bb.Max);
                var uMin = double.MaxValue; var uMax = double.MinValue;
                var vMin = double.MaxValue; var vMax = double.MinValue;

                foreach (var c in corners)
                {
                    var (ok, u, v) = TryProjectToUv(c, right, up);
                    if (!ok) continue;
                    if (u < uMin) uMin = u; if (u > uMax) uMax = u;
                    if (v < vMin) vMin = v; if (v > vMax) vMax = v;
                }

                if (uMin > uMax || vMin > vMax) return (false, 0, "2D BoundingBox を構築できません。");

                var tmin = 0.0;
                var tmax = double.MaxValue;
                if (!UpdateSlab(uo, du, uMin, uMax, ref tmin, ref tmax)) return (false, 0, string.Empty);
                if (!UpdateSlab(vo, dv, vMin, vMax, ref tmin, ref tmax)) return (false, 0, string.Empty);
                if (tmax < 0) return (false, 0, string.Empty);

                return (true, tmin, string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, 0, $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        private (bool ok, BoundingBoxXYZ bb, string reason) TryGetBoundingBoxPreferView(Element element, View view)
        {
            try
            {
                if (element == null) return (false, null, "Element が null です。");
                var (hasViewBb, viewBb, viewReason) = TryGetBoundingBox(element, view);
                if (hasViewBb) return (true, viewBb, string.Empty);
                var (hasModelBb, modelBb, modelReason) = TryGetBoundingBox(element, null);
                if (hasModelBb) return (true, modelBb, string.Empty);
                var reason = string.Join(" / ", new[] { viewReason, modelReason }.Where(x => !string.IsNullOrWhiteSpace(x)));
                return (false, null, string.IsNullOrWhiteSpace(reason) ? "BoundingBox を取得できません。" : reason);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, null, $"{ex.GetType().Name}: {ex.Message}"); }
        }

        private (bool ok, BoundingBoxXYZ bb, string reason) TryGetBoundingBox(Element element, View view)
        {
            try
            {
                if (element == null) return (false, null, "Element が null です。");
                BoundingBoxXYZ bb;
                try { bb = element.get_BoundingBox(view); }
                catch (Autodesk.Revit.Exceptions.InternalException iex)
                { return (false, null, $"get_BoundingBox: InternalException: {iex.Message}"); }
                catch (Exception ex)
                { return (false, null, $"get_BoundingBox: {ex.GetType().Name}: {ex.Message}"); }
                if (bb == null || bb.Min == null || bb.Max == null) return (false, null, "BoundingBox が null です。");
                return (true, bb, string.Empty);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, null, $"{ex.GetType().Name}: {ex.Message}"); }
        }

        private IReadOnlyList<XYZ> BuildBoundingBoxCorners(XYZ min, XYZ max)
        {
            var result = new List<XYZ>();
            if (min == null || max == null) return result;
            foreach (var x in new[] { min.X, max.X })
                foreach (var y in new[] { min.Y, max.Y })
                    foreach (var z in new[] { min.Z, max.Z })
                        result.Add(new XYZ(x, y, z));
            return result;
        }

        private bool UpdateSlab(double o, double d, double min, double max, ref double tmin, ref double tmax)
        {
            if (Math.Abs(d) < Eps)
                return !(o < min || o > max);

            var t1 = (min - o) / d;
            var t2 = (max - o) / d;
            if (t1 > t2) { var tmp = t1; t1 = t2; t2 = tmp; }
            if (t1 > tmin) tmin = t1;
            if (t2 < tmax) tmax = t2;
            return tmin <= tmax;
        }

        // -------------------------------------------------------------------------
        // bending detail host resolution
        // -------------------------------------------------------------------------

        private (bool ok, ElementId rebarId) TryResolveHostRebarIdFromBendingDetail(Document doc, Element bendingDetail)
        {
            var (okByStaticApi, idByStaticApi) = TryResolveHostByRebarBendingDetailStaticApi(doc, bendingDetail);
            if (okByStaticApi) return (true, idByStaticApi);
            var (okByApi, idByApi) = TryResolveHostByReflection(doc, bendingDetail);
            if (okByApi) return (true, idByApi);
            return TryResolveHostByParameters(doc, bendingDetail);
        }

        private (bool ok, ElementId rebarId) TryResolveHostByRebarBendingDetailStaticApi(Document doc, Element bendingDetail)
        {
            try
            {
                if (doc == null || bendingDetail == null) return (false, ElementId.InvalidElementId);
                var apiType = typeof(Autodesk.Revit.DB.Structure.RebarBendingDetail);
                var methods = apiType
                    .GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static)
                    .Where(m => string.Equals(m.Name, "GetHost", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(m.Name, "GetHosts", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var m in methods)
                {
                    var (canCall, args) = TryBuildArgsForBendingDetailApiCall(m, doc, bendingDetail);
                    if (!canCall) continue;
                    object result;
                    try { result = m.Invoke(null, args); }
                    catch (System.Reflection.TargetInvocationException tex) { Debug.WriteLine(tex.InnerException ?? tex); continue; }
                    catch (Exception ex) { Debug.WriteLine(ex); continue; }
                    var (ok, rebarId) = TryExtractRebarIdFromHostResult(doc, result);
                    if (ok) return (true, rebarId);
                }
                return (false, ElementId.InvalidElementId);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, ElementId.InvalidElementId); }
        }

        private (bool ok, object[] args) TryBuildArgsForBendingDetailApiCall(
            System.Reflection.MethodInfo method, Document doc, Element bendingDetail)
        {
            try
            {
                if (method == null || doc == null || bendingDetail == null) return (false, null);
                var ps = method.GetParameters();
                if (ps == null) return (false, null);
                if (ps.Length == 2 && ps[0].ParameterType == typeof(Document) && ps[1].ParameterType == typeof(ElementId))
                    return (true, new object[] { doc, bendingDetail.Id });
                if (ps.Length == 2 && ps[0].ParameterType == typeof(Document) && ps[1].ParameterType.IsAssignableFrom(bendingDetail.GetType()))
                    return (true, new object[] { doc, bendingDetail });
                if (ps.Length == 1 && ps[0].ParameterType == typeof(ElementId))
                    return (true, new object[] { bendingDetail.Id });
                if (ps.Length == 1 && ps[0].ParameterType.IsAssignableFrom(bendingDetail.GetType()))
                    return (true, new object[] { bendingDetail });
                return (false, null);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, null); }
        }

        private (bool ok, ElementId rebarId) TryExtractRebarIdFromHostResult(Document doc, object result)
        {
            try
            {
                if (doc == null || result == null) return (false, ElementId.InvalidElementId);
                if (result is ElementId id)
                    return doc.GetElement(id) is Rebar ? (true, id) : (false, ElementId.InvalidElementId);
                if (result is Element e)
                    return e is Rebar ? (true, e.Id) : (false, ElementId.InvalidElementId);
                if (result is Reference r)
                {
                    var eid = r.ElementId;
                    return doc.GetElement(eid) is Rebar ? (true, eid) : (false, ElementId.InvalidElementId);
                }
                if (result is System.Collections.IEnumerable en)
                {
                    foreach (var item in en)
                    {
                        var (ok, rebarId) = TryExtractRebarIdFromHostResult(doc, item);
                        if (ok) return (true, rebarId);
                    }
                }
                return (false, ElementId.InvalidElementId);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, ElementId.InvalidElementId); }
        }

        private (bool ok, ElementId rebarId) TryResolveHostByReflection(Document doc, Element bendingDetail)
        {
            try
            {
                if (doc == null || bendingDetail == null) return (false, ElementId.InvalidElementId);
                var t = bendingDetail.GetType();

                var mDocId = t.GetMethod("GetHost", new[] { typeof(Document), typeof(ElementId) });
                if (mDocId != null)
                {
                    var target = mDocId.IsStatic ? null : (object)bendingDetail;
                    var obj = mDocId.Invoke(target, new object[] { doc, bendingDetail.Id });
                    if (obj is ElementId id && doc.GetElement(id) is Rebar) return (true, id);
                    if (obj is Element host && host is Rebar) return (true, host.Id);
                }

                var mId = t.GetMethod("GetHost", new[] { typeof(ElementId) });
                if (mId != null)
                {
                    var target = mId.IsStatic ? null : (object)bendingDetail;
                    var obj = mId.Invoke(target, new object[] { bendingDetail.Id });
                    if (obj is ElementId id && doc.GetElement(id) is Rebar) return (true, id);
                    if (obj is Element host && host is Rebar) return (true, host.Id);
                }

                var m2 = t.GetMethod("GetHost", Type.EmptyTypes);
                if (m2 != null)
                {
                    var obj = m2.Invoke(bendingDetail, null);
                    if (obj is Element host && host is Rebar) return (true, host.Id);
                }

                var p = t.GetProperty("HostId");
                if (p != null && p.PropertyType == typeof(ElementId))
                {
                    var obj = p.GetValue(bendingDetail, null);
                    if (obj is ElementId hostId && doc.GetElement(hostId) is Rebar) return (true, hostId);
                }

                return (false, ElementId.InvalidElementId);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, ElementId.InvalidElementId); }
        }

        private (bool ok, ElementId rebarId) TryResolveHostByParameters(Document doc, Element bendingDetail)
        {
            try
            {
                if (doc == null || bendingDetail == null) return (false, ElementId.InvalidElementId);
                foreach (Parameter p in bendingDetail.Parameters)
                {
                    if (p == null || p.StorageType != StorageType.ElementId) continue;
                    var id = p.AsElementId();
                    if (id == null || id == ElementId.InvalidElementId) continue;
                    if (doc.GetElement(id) is Rebar) return (true, id);
                }
                return (false, ElementId.InvalidElementId);
            }
            catch (Exception ex) { Debug.WriteLine(ex); return (false, ElementId.InvalidElementId); }
        }

        // -------------------------------------------------------------------------
        // misc helpers
        // -------------------------------------------------------------------------

        private string BuildStageMessage(string stage, string reason)
        {
            if (string.IsNullOrWhiteSpace(stage)) return reason ?? string.Empty;
            if (string.IsNullOrWhiteSpace(reason)) return stage;
            return $"{stage}: {reason}";
        }

        private string SafeElementIdText(Element e)
        {
            try
            {
                if (e == null) return "?";
                var id = e.Id;
                return (id == null || id == ElementId.InvalidElementId) ? "?" : id.Value.ToString();
            }
            catch (Exception ex) { Debug.WriteLine(ex); return "?"; }
        }

        private RebarTagLeaderBendingDetailCheckItem CreateFail(
            ElementId tagId, ElementId taggedRebarId, XYZ leaderEnd, string message)
        {
            return new RebarTagLeaderBendingDetailCheckItem(
                tagId, taggedRebarId, leaderEnd,
                ElementId.InvalidElementId, ElementId.InvalidElementId,
                false, message);
        }
    }
}
