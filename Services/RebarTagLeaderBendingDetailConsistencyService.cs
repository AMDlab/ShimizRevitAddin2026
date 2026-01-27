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
            try
            {
                var tagIds = CollectBendingDetailTagIds(doc, rebar, view);
                var tags = ResolveIndependentTags(doc, tagIds);
                var (freeTag, segStart0, segEnd0) = TrySelectFreeEndTagWithLastSegment(tags);
                if (freeTag == null)
                {
                    // 自由端タグが無ければ判定しない（NGにしない）
                    return null;
                }

                var (hasTaggedRebar, taggedRebarId) = TryResolveTaggedRebarId(doc, freeTag);
                if (!hasTaggedRebar)
                {
                    return CreateFail(freeTag.Id, ElementId.InvalidElementId, null, "自由端タグに紐づく鉄筋を取得できません。");
                }

                var segStart = segStart0;
                var segEnd = segEnd0;

                var (hasRay, rayOrigin, rayDir, rayReason) = TryBuildRayFromLastSegment(segStart, segEnd);
                if (!hasRay)
                {
                    return CreateFail(freeTag.Id, taggedRebarId, segEnd, rayReason);
                }

                var (hasHit, hitDetail, hitPoint, hitReason) = TryFindFirstIntersectionWithBendingDetail(bendingDetails, view, rayOrigin, rayDir);
                if (!hasHit || hitDetail == null)
                {
                    return CreateFail(freeTag.Id, taggedRebarId, segEnd, hitReason);
                }

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
                        "曲げ詳細から鉄筋を特定できません。");
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
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return CreateFail(rebar?.Id ?? ElementId.InvalidElementId, rebar?.Id ?? ElementId.InvalidElementId, null, ex.Message);
            }
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

        private (IndependentTag tag, XYZ segStart, XYZ segEnd) TrySelectFreeEndTagWithLastSegment(IReadOnlyList<IndependentTag> tags)
        {
            try
            {
                if (tags == null || tags.Count == 0)
                {
                    return (null, null, null);
                }

                foreach (var tag in tags)
                {
                    if (tag == null)
                    {
                        continue;
                    }

                    if (!tag.HasLeader)
                    {
                        continue;
                    }

                    var (ok, s, e, _) = TryGetLeaderLastSegment(tag);
                    if (ok)
                    {
                        return (tag, s, e);
                    }
                }

                return (null, null, null);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (null, null, null);
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
                        var end = tag.GetLeaderEnd(r);
                        if (end == null)
                        {
                            errors.Add("GetLeaderEnd が null を返しました。");
                            continue;
                        }

                        var elbow = TryGetLeaderElbow(tag, r);
                        var head = tag.TagHeadPosition;

                        // 最後の線分: elbow->end、elbow無しなら head->end
                        var start = elbow ?? head;
                        if (start == null)
                        {
                            return (false, null, null, "始点を取得できません。");
                        }

                        return (true, start, end, string.Empty);
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

                var (hasBasis, right, up) = TryGetViewBasis(view);
                if (!hasBasis)
                {
                    return (false, null, null, "ビュー座標系を取得できません。");
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

                foreach (var e in bendingDetails)
                {
                    var (ok, t) = TryIntersectRayWithBoundingBox2D(e, uo, vo, du, dv, right, up);
                    if (!ok)
                    {
                        continue;
                    }

                    if (t < bestT)
                    {
                        bestT = t;
                        best = e;
                        bestPoint = rayOrigin + (rayDir * t);
                    }
                }

                if (best == null)
                {
                    return (false, null, null, "引き出し線の延長と曲げ詳細の交点を取得できません。");
                }

                return (true, best, bestPoint, string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null, null, $"{ex.GetType().Name}: {ex.Message}");
            }
        }

        private (bool ok, XYZ right, XYZ up) TryGetViewBasis(View view)
        {
            try
            {
                if (view == null)
                {
                    return (false, null, null);
                }

                // View の向き（2D座標系）
                var right = view.RightDirection;
                var up = view.UpDirection;
                if (right == null || up == null)
                {
                    return (false, null, null);
                }

                return (true, right, up);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null, null);
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

        private (bool ok, double tEnter) TryIntersectRayWithBoundingBox2D(
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
                    return (false, 0);
                }

                var bb = element.get_BoundingBox(null);
                if (bb == null || bb.Min == null || bb.Max == null)
                {
                    return (false, 0);
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
                    return (false, 0);
                }

                var tmin = 0.0;
                var tmax = double.MaxValue;

                if (!UpdateSlab(uo, du, uMin, uMax, ref tmin, ref tmax)) return (false, 0);
                if (!UpdateSlab(vo, dv, vMin, vMax, ref tmin, ref tmax)) return (false, 0);

                if (tmax < 0)
                {
                    return (false, 0);
                }

                return (true, tmin);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, 0);
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

