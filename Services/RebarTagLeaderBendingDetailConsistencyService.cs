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
            var item = BuildResultForRebar(doc, rebar, view, bendingDetails);
            return item == null ? new List<RebarTagLeaderBendingDetailCheckItem>() : new List<RebarTagLeaderBendingDetailCheckItem> { item };
        }

        public IReadOnlyCollection<ElementId> CollectMismatchRebarIds(Document doc, IReadOnlyList<Rebar> rebars, View view)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (rebars == null || rebars.Count == 0) return new List<ElementId>();

            var bendingDetails = CollectBendingDetailElementsInView(doc, view);
            return CollectMismatchRebarIdsInternal(doc, rebars, view, bendingDetails);
        }

        private IReadOnlyCollection<ElementId> CollectMismatchRebarIdsInternal(
            Document doc,
            IReadOnlyList<Rebar> rebars,
            View view,
            IReadOnlyList<Element> bendingDetails)
        {
            var result = new HashSet<ElementId>();

            foreach (var rebar in rebars)
            {
                if (rebar == null)
                {
                    continue;
                }

                var items = CheckInternal(doc, rebar, view, bendingDetails);
                if (HasAnyMismatch(items))
                {
                    result.Add(rebar.Id);
                }
            }

            return result.ToList();
        }

        private bool HasAnyMismatch(IReadOnlyList<RebarTagLeaderBendingDetailCheckItem> items)
        {
            if (items == null || items.Count == 0)
            {
                return false;
            }

            // 不一致、または判定不能も NG とみなす
            return items.Any(x => x != null && !x.IsMatch);
        }

        private IReadOnlyList<RebarTagLeaderBendingDetailCheckItem> CheckInternal(
            Document doc,
            Rebar rebar,
            View view,
            IReadOnlyList<Element> bendingDetails)
        {
            var item = BuildResultForRebar(doc, rebar, view, bendingDetails);
            return item == null ? new List<RebarTagLeaderBendingDetailCheckItem>() : new List<RebarTagLeaderBendingDetailCheckItem> { item };
        }

        private IReadOnlyList<ElementId> CollectBendingDetailTagIds(Document doc, Rebar rebar, View view)
        {
            try
            {
                if (_dependentTagCollector == null)
                {
                    return new List<ElementId>();
                }

                var model = _dependentTagCollector.Collect(doc, rebar, view);
                return model?.BendingDetailTagIds?.ToList() ?? new List<ElementId>();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ElementId>();
            }
        }

        private IReadOnlyList<Element> CollectBendingDetailElementsInView(Document doc, View view)
        {
            try
            {
                // 曲げ詳細はビュー専用要素として配置されるため、ビュー内の要素を走査する
                var collector = new FilteredElementCollector(doc, view.Id)
                    .WhereElementIsNotElementType();

                var all = collector.ToElements();
                return all
                    .Where(IsBendingDetailElement)
                    .ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<Element>();
            }
        }

        private bool IsBendingDetailElement(Element e)
        {
            try
            {
                if (e == null)
                {
                    return false;
                }

                // Revit API の型名に依存するが、コンパイル時の参照を避けつつ特定する
                var typeName = e.GetType().Name ?? string.Empty;
                if (typeName.IndexOf("RebarBendingDetail", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return false;
            }
        }

        private RebarTagLeaderBendingDetailCheckItem BuildResultForRebar(
            Document doc,
            Rebar rebar,
            View view,
            IReadOnlyList<Element> bendingDetails)
        {
            var stage = "Init";
            try
            {
                stage = "CollectBendingDetailTagIds";
                var tagIds = CollectBendingDetailTagIds(doc, rebar, view);
                if (tagIds == null || tagIds.Count == 0)
                {
                    return CreateFail(ElementId.InvalidElementId, rebar?.Id ?? ElementId.InvalidElementId, null, BuildStageMessage(stage, "曲げ加工詳細タグがありません。"));
                }

                stage = "ResolveIndependentTags";
                var tags = ResolveIndependentTags(doc, tagIds);
                if (tags == null || tags.Count == 0)
                {
                    return CreateFail(tagIds.FirstOrDefault() ?? ElementId.InvalidElementId, rebar?.Id ?? ElementId.InvalidElementId, null, BuildStageMessage(stage, "曲げ加工詳細タグ要素を取得できません。"));
                }

                stage = "TrySelectFreeEndTagWithLastSegment";
                var (freeTag, segStart0, segEnd0, selectReason) = TrySelectFreeEndTagWithLastSegment(tags);
                if (freeTag == null)
                {
                    return CreateFail(tagIds.FirstOrDefault() ?? ElementId.InvalidElementId, rebar?.Id ?? ElementId.InvalidElementId, null, BuildStageMessage(stage, selectReason));
                }

                stage = "TryResolveTaggedRebarId";
                var (hasTaggedRebar, taggedRebarId) = TryResolveTaggedRebarId(doc, freeTag);
                if (!hasTaggedRebar)
                {
                    return CreateFail(freeTag.Id, ElementId.InvalidElementId, null, BuildStageMessage(stage, "自由端タグに紐づく鉄筋を取得できません。"));
                }

                var segStart = segStart0;
                var segEnd = segEnd0;

                stage = "TryBuildRayFromLastSegment";
                var (hasRay, rayOrigin, rayDir, rayReason) = TryBuildRayFromLastSegment(segStart, segEnd);
                if (!hasRay)
                {
                    return CreateFail(freeTag.Id, taggedRebarId, segEnd, BuildStageMessage(stage, rayReason));
                }

                stage = "TryFindFirstIntersectionWithBendingDetail";
                var (hasHit, hitDetail, hitPoint, hitReason) = TryFindFirstIntersectionWithBendingDetail(bendingDetails, view, rayOrigin, rayDir);
                if (!hasHit || hitDetail == null)
                {
                    return CreateFail(freeTag.Id, taggedRebarId, segEnd, BuildStageMessage(stage, hitReason));
                }

                stage = "TryResolveHostRebarIdFromBendingDetail";
                var (hasHost, hostRebarId) = TryResolveHostRebarIdFromBendingDetail(doc, hitDetail);
                if (!hasHost)
                {
                    return new RebarTagLeaderBendingDetailCheckItem(
                        freeTag.Id,
                        taggedRebarId,
                        hitPoint,
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
                        hitPoint,
                        hitDetail.Id,
                        hostRebarId,
                        true,
                        "一致。");
                }

                return new RebarTagLeaderBendingDetailCheckItem(
                    freeTag.Id,
                    taggedRebarId,
                    hitPoint,
                    hitDetail.Id,
                    hostRebarId,
                    false,
                    $"不一致（曲げ詳細Host={hostRebarId.Value} / タグ対象鉄筋={taggedRebarId.Value}）。");
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

        private string BuildStageMessage(string stage, string reason)
        {
            if (string.IsNullOrWhiteSpace(stage))
            {
                return reason ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(reason))
            {
                return stage;
            }

            // ステージ名を必ず残して原因を追えるようにする
            return $"{stage}: {reason}";
        }

        private RebarTagLeaderBendingDetailCheckItem CreateFail(
            ElementId tagId,
            ElementId taggedRebarId,
            XYZ leaderEnd,
            string message)
        {
            return new RebarTagLeaderBendingDetailCheckItem(
                tagId,
                taggedRebarId,
                leaderEnd,
                ElementId.InvalidElementId,
                ElementId.InvalidElementId,
                false,
                message);
        }

        private (bool ok, ElementId rebarId) TryResolveTaggedRebarId(Document doc, IndependentTag tag)
        {
            try
            {
                if (doc == null || tag == null)
                {
                    return (false, ElementId.InvalidElementId);
                }

                var ids = tag.GetTaggedLocalElementIds();
                if (ids == null || ids.Count == 0)
                {
                    return (false, ElementId.InvalidElementId);
                }

                foreach (var id in ids)
                {
                    if (id == null || id == ElementId.InvalidElementId)
                    {
                        continue;
                    }

                    if (doc.GetElement(id) is Rebar)
                    {
                        return (true, id);
                    }
                }

                return (false, ElementId.InvalidElementId);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, ElementId.InvalidElementId);
            }
        }

        private IReadOnlyList<IndependentTag> ResolveIndependentTags(Document doc, IReadOnlyList<ElementId> tagIds)
        {
            try
            {
                if (doc == null || tagIds == null || tagIds.Count == 0)
                {
                    return new List<IndependentTag>();
                }

                var result = new List<IndependentTag>();
                foreach (var id in tagIds)
                {
                    if (id == null || id == ElementId.InvalidElementId)
                    {
                        continue;
                    }

                    var e = doc.GetElement(id) as IndependentTag;
                    if (e == null)
                    {
                        continue;
                    }

                    result.Add(e);
                }

                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<IndependentTag>();
            }
        }

        private (IndependentTag tag, XYZ segStart, XYZ segEnd, string reason) TrySelectFreeEndTagWithLastSegment(IReadOnlyList<IndependentTag> tags)
        {
            try
            {
                if (tags == null || tags.Count == 0)
                {
                    return (null, null, null, "曲げ加工詳細タグがありません。");
                }

                // 取得失敗の理由を集約して、後でメッセージに出す
                var reasons = new List<string>();

                foreach (var tag in tags)
                {
                    if (tag == null)
                    {
                        continue;
                    }

                    if (!tag.HasLeader)
                    {
                        reasons.Add($"{tag.Id.Value}: HasLeader=false");
                        continue;
                    }

                    var (ok, s, e, r) = TryGetLeaderLastSegment(tag);
                    if (ok)
                    {
                        return (tag, s, e, string.Empty);
                    }

                    var cond = TryGetLeaderEndConditionText(tag);
                    if (string.IsNullOrWhiteSpace(cond))
                    {
                        reasons.Add($"{tag.Id.Value}: {r}");
                    }
                    else
                    {
                        reasons.Add($"{tag.Id.Value}: LeaderEndCondition={cond} / {r}");
                    }
                }

                if (reasons.Count == 0)
                {
                    return (null, null, null, "自由端タグの候補がありません。");
                }

                return (null, null, null, "自由端タグの線分を取得できません。\n" + string.Join("\n", reasons));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (null, null, null, $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        private string TryGetLeaderEndConditionText(IndependentTag tag)
        {
            try
            {
                if (tag == null)
                {
                    return string.Empty;
                }

                // 引出線タイプ（自由な端点など）を表示するため反射で取得する
                var p = tag.GetType().GetProperty("LeaderEndCondition");
                if (p == null)
                {
                    return string.Empty;
                }

                var v = p.GetValue(tag, null);
                return v == null ? string.Empty : v.ToString();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return string.Empty;
            }
        }

        private (bool ok, XYZ start, XYZ end, string reason) TryGetLeaderLastSegment(IndependentTag tag)
        {
            try
            {
                if (tag == null)
                {
                    return (false, null, null, "タグが null です。");
                }

                if (!tag.HasLeader)
                {
                    return (false, null, null, "HasLeader=false");
                }

                var refs = TryGetTaggedReferences(tag);
                if (refs.Count == 0)
                {
                    return (false, null, null, "タグ参照を取得できません。");
                }

                var errors = new List<string>();
                foreach (var r in refs)
                {
                    if (r == null)
                    {
                        continue;
                    }

                    try
                    {
                        // 自由端タグのみを対象にする想定
                        XYZ end = null;
                        try
                        {
                            end = tag.GetLeaderEnd(r);
                        }
                        catch (Autodesk.Revit.Exceptions.InternalException iex)
                        {
                            errors.Add($"GetLeaderEnd: InternalException: {iex.Message}");
                            continue;
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"GetLeaderEnd: {ex.GetType().Name}: {ex.Message}");
                            continue;
                        }

                        if (end == null)
                        {
                            errors.Add("GetLeaderEnd が null を返しました。");
                            continue;
                        }

                        var elbow = TryGetLeaderElbow(tag, r);
                        XYZ head = null;
                        try
                        {
                            head = tag.TagHeadPosition;
                        }
                        catch (Autodesk.Revit.Exceptions.InternalException iex)
                        {
                            errors.Add($"TagHeadPosition: InternalException: {iex.Message}");
                            continue;
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"TagHeadPosition: {ex.GetType().Name}: {ex.Message}");
                            continue;
                        }

                        // 最後の線分: elbow->end、elbow無しなら head->end
                        var start = elbow ?? head;
                        if (start == null)
                        {
                            return (false, null, null, "始点を取得できません。");
                        }

                        return (true, start, end, string.Empty);
                    }
                    catch (Autodesk.Revit.Exceptions.InternalException iex)
                    {
                        errors.Add($"InternalException: {iex.Message}");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{ex.GetType().Name}: {ex.Message}");
                    }
                }

                if (errors.Count == 0)
                {
                    return (false, null, null, "GetLeaderEnd に失敗しました。");
                }

                var unique = errors.Distinct().ToList();
                return (false, null, null, string.Join(" / ", unique));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null, null, $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        private IReadOnlyList<Reference> TryGetTaggedReferences(IndependentTag tag)
        {
            try
            {
                if (tag == null)
                {
                    return new List<Reference>();
                }

                // API バージョン差を吸収するため、まずは GetTaggedReferences() を反射で取得する
                var m = tag.GetType().GetMethod("GetTaggedReferences", Type.EmptyTypes);
                if (m == null)
                {
                    return new List<Reference>();
                }

                var obj = m.Invoke(tag, null);
                if (obj is IEnumerable<Reference> refs)
                {
                    return refs.Where(x => x != null).ToList();
                }

                return new List<Reference>();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<Reference>();
            }
        }

        private XYZ TryGetLeaderElbow(IndependentTag tag, Reference r)
        {
            try
            {
                if (tag == null || r == null)
                {
                    return null;
                }

                // GetLeaderElbow(Reference) はバージョン差があるため反射で呼ぶ
                var m = tag.GetType().GetMethod("GetLeaderElbow", new[] { typeof(Reference) });
                if (m == null)
                {
                    return null;
                }

                var obj = m.Invoke(tag, new object[] { r });
                return obj as XYZ;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return null;
            }
        }

        private (bool ok, XYZ origin, XYZ dir, string reason) TryBuildRayFromLastSegment(XYZ start, XYZ end)
        {
            try
            {
                if (start == null || end == null)
                {
                    return (false, null, null, "線分が無効です。");
                }

                var v = end - start;
                var len = v.GetLength();
                if (len < Eps)
                {
                    return (false, null, null, "線分の長さが 0 です。");
                }

                // 最後の線分を延長: end を起点に同方向へ
                return (true, end, v / len, string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null, null, $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        private (bool ok, Element bendingDetail, XYZ hitPoint, string reason) TryFindFirstIntersectionWithBendingDetail(
            IReadOnlyList<Element> bendingDetails,
            View view,
            XYZ rayOrigin,
            XYZ rayDir)
        {
            try
            {
                if (bendingDetails == null || bendingDetails.Count == 0)
                {
                    return (false, null, null, "曲げ詳細が見つかりません。");
                }

                var (hasBasis, right, up, basisReason) = TryGetViewBasis(view);
                if (!hasBasis)
                {
                    return (false, null, null, string.IsNullOrWhiteSpace(basisReason) ? "ビュー座標系を取得できません。" : basisReason);
                }

                if (rayOrigin == null || rayDir == null)
                {
                    return (false, null, null, "レイが無効です。");
                }

                var (okO, uo, vo) = TryProjectToUv(rayOrigin, right, up);
                var (okD, du, dv) = TryProjectDirToUv(rayDir, right, up);
                if (!okO || !okD)
                {
                    return (false, null, null, "レイの投影に失敗しました。");
                }

                Element best = null;
                XYZ bestPoint = null;
                var bestT = double.MaxValue;
                var errors = new List<string>();

                foreach (var e in bendingDetails)
                {
                    try
                    {
                        var (ok, t, reason) = TryIntersectRayWithBoundingBox2D(e, uo, vo, du, dv, right, up);
                        if (!ok)
                        {
                            if (!string.IsNullOrWhiteSpace(reason) && e != null)
                            {
                                errors.Add($"{e.Id.Value}: {reason}");
                            }
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
                        if (e != null)
                        {
                            errors.Add($"{e.Id.Value}: InternalException: {iex.Message}");
                        }
                        else
                        {
                            errors.Add($"InternalException: {iex.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (e != null)
                        {
                            errors.Add($"{e.Id.Value}: {ex.GetType().Name}: {ex.Message}");
                        }
                        else
                        {
                            errors.Add($"{ex.GetType().Name}: {ex.Message}");
                        }
                    }
                }

                if (best == null)
                {
                    if (errors.Count == 0)
                    {
                        return (false, null, null, "引き出し線の延長と曲げ詳細の交点を取得できません。");
                    }

                    var unique = errors.Distinct().Take(8).ToList();
                    return (false, null, null, "曲げ詳細との交点取得で例外が発生しました。\n" + string.Join("\n", unique));
                }

                return (true, best, bestPoint, string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null, null, $"TryFindFirstIntersectionWithBendingDetail: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private (bool ok, XYZ right, XYZ up, string reason) TryGetViewBasis(View view)
        {
            try
            {
                if (view == null)
                {
                    return (false, null, null, "View が null です。");
                }

                // View の向き（2D座標系）
                var right = view.RightDirection;
                var up = view.UpDirection;
                if (right == null || up == null)
                {
                    return (false, null, null, "RightDirection/UpDirection を取得できません。");
                }

                return (true, right, up, string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null, null, $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        private (bool ok, double u, double v) TryProjectToUv(XYZ p, XYZ right, XYZ up)
        {
            try
            {
                if (p == null || right == null || up == null)
                {
                    return (false, 0, 0);
                }

                var u = p.DotProduct(right);
                var v = p.DotProduct(up);
                return (true, u, v);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, 0, 0);
            }
        }

        private (bool ok, double du, double dv) TryProjectDirToUv(XYZ d, XYZ right, XYZ up)
        {
            try
            {
                if (d == null || right == null || up == null)
                {
                    return (false, 0, 0);
                }

                var du = d.DotProduct(right);
                var dv = d.DotProduct(up);
                if (Math.Abs(du) < Eps && Math.Abs(dv) < Eps)
                {
                    return (false, 0, 0);
                }

                return (true, du, dv);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, 0, 0);
            }
        }

        private (bool ok, double tEnter, string reason) TryIntersectRayWithBoundingBox2D(
            Element element,
            double uo,
            double vo,
            double du,
            double dv,
            XYZ right,
            XYZ up)
        {
            try
            {
                if (element == null)
                {
                    return (false, 0, "Element が null です。");
                }

                var bb = element.get_BoundingBox(null);
                if (bb == null || bb.Min == null || bb.Max == null)
                {
                    return (false, 0, "BoundingBox を取得できません。");
                }

                // 8頂点をビュー座標に投影して2DのAABBを作る
                var corners = BuildBoundingBoxCorners(bb.Min, bb.Max);
                var uMin = double.MaxValue;
                var uMax = double.MinValue;
                var vMin = double.MaxValue;
                var vMax = double.MinValue;

                foreach (var c in corners)
                {
                    var (ok, u, v) = TryProjectToUv(c, right, up);
                    if (!ok)
                    {
                        continue;
                    }

                    if (u < uMin) uMin = u;
                    if (u > uMax) uMax = u;
                    if (v < vMin) vMin = v;
                    if (v > vMax) vMax = v;
                }

                if (uMin > uMax || vMin > vMax)
                {
                    return (false, 0, "2D BoundingBox を構築できません。");
                }

                var tmin = 0.0;
                var tmax = double.MaxValue;

                if (!UpdateSlab(uo, du, uMin, uMax, ref tmin, ref tmax)) return (false, 0, string.Empty);
                if (!UpdateSlab(vo, dv, vMin, vMax, ref tmin, ref tmax)) return (false, 0, string.Empty);

                if (tmax < 0)
                {
                    return (false, 0, string.Empty);
                }

                return (true, tmin, string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, 0, $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        private IReadOnlyList<XYZ> BuildBoundingBoxCorners(XYZ min, XYZ max)
        {
            var result = new List<XYZ>();
            if (min == null || max == null)
            {
                return result;
            }

            var xs = new[] { min.X, max.X };
            var ys = new[] { min.Y, max.Y };
            var zs = new[] { min.Z, max.Z };

            foreach (var x in xs)
            {
                foreach (var y in ys)
                {
                    foreach (var z in zs)
                    {
                        result.Add(new XYZ(x, y, z));
                    }
                }
            }

            return result;
        }

        private bool UpdateSlab(double o, double d, double min, double max, ref double tmin, ref double tmax)
        {
            if (Math.Abs(d) < Eps)
            {
                if (o < min || o > max)
                {
                    return false;
                }

                return true;
            }

            var t1 = (min - o) / d;
            var t2 = (max - o) / d;
            if (t1 > t2)
            {
                var tmp = t1;
                t1 = t2;
                t2 = tmp;
            }

            if (t1 > tmin) tmin = t1;
            if (t2 < tmax) tmax = t2;
            return tmin <= tmax;
        }

        private (bool ok, ElementId rebarId) TryResolveHostRebarIdFromBendingDetail(Document doc, Element bendingDetail)
        {
            // まずは型特有の API があれば反射で拾い、無ければパラメータから Rebar を探す
            var (okByApi, idByApi) = TryResolveHostByReflection(doc, bendingDetail);
            if (okByApi)
            {
                return (true, idByApi);
            }

            return TryResolveHostByParameters(doc, bendingDetail);
        }

        private (bool ok, ElementId rebarId) TryResolveHostByReflection(Document doc, Element bendingDetail)
        {
            try
            {
                if (doc == null || bendingDetail == null)
                {
                    return (false, ElementId.InvalidElementId);
                }

                var t = bendingDetail.GetType();

                // まずは GetHost(Document, ElementId) 形式（静的/インスタンス両対応）を試す
                var mDocId = t.GetMethod("GetHost", new[] { typeof(Document), typeof(ElementId) });
                if (mDocId != null)
                {
                    var target = mDocId.IsStatic ? null : (object)bendingDetail;
                    var obj = mDocId.Invoke(target, new object[] { doc, bendingDetail.Id });
                    if (obj is ElementId id && doc.GetElement(id) is Rebar)
                    {
                        return (true, id);
                    }

                    if (obj is Element host && host is Rebar)
                    {
                        return (true, host.Id);
                    }
                }

                // 次に GetHost(ElementId) 形式（静的/インスタンス両対応）を試す
                var mId = t.GetMethod("GetHost", new[] { typeof(ElementId) });
                if (mId != null)
                {
                    var target = mId.IsStatic ? null : (object)bendingDetail;
                    var obj = mId.Invoke(target, new object[] { bendingDetail.Id });
                    if (obj is ElementId id && doc.GetElement(id) is Rebar)
                    {
                        return (true, id);
                    }

                    if (obj is Element host && host is Rebar)
                    {
                        return (true, host.Id);
                    }
                }

                // GetHost() がインスタンスメソッドで Element を返す場合
                var m2 = t.GetMethod("GetHost", Type.EmptyTypes);
                if (m2 != null)
                {
                    var obj = m2.Invoke(bendingDetail, null);
                    if (obj is Element host && host is Rebar)
                    {
                        return (true, host.Id);
                    }
                }

                // HostId プロパティがある場合
                var p = t.GetProperty("HostId");
                if (p != null && p.PropertyType == typeof(ElementId))
                {
                    var obj = p.GetValue(bendingDetail, null);
                    if (obj is ElementId hostId && doc.GetElement(hostId) is Rebar)
                    {
                        return (true, hostId);
                    }
                }

                return (false, ElementId.InvalidElementId);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, ElementId.InvalidElementId);
            }
        }

        private (bool ok, ElementId rebarId) TryResolveHostByParameters(Document doc, Element bendingDetail)
        {
            try
            {
                if (doc == null || bendingDetail == null)
                {
                    return (false, ElementId.InvalidElementId);
                }

                foreach (Parameter p in bendingDetail.Parameters)
                {
                    if (p == null)
                    {
                        continue;
                    }

                    if (p.StorageType != StorageType.ElementId)
                    {
                        continue;
                    }

                    var id = p.AsElementId();
                    if (id == null || id == ElementId.InvalidElementId)
                    {
                        continue;
                    }

                    var e = doc.GetElement(id);
                    if (e is Rebar)
                    {
                        return (true, id);
                    }
                }

                return (false, ElementId.InvalidElementId);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, ElementId.InvalidElementId);
            }
        }

        // 旧ロジック（最近傍検索・距離表示）は削除済み
    }
}

