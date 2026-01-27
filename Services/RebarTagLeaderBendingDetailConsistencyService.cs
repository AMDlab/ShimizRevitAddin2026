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

        public RebarTagLeaderBendingDetailConsistencyService(RebarDependentTagCollector dependentTagCollector)
        {
            _dependentTagCollector = dependentTagCollector;
        }

        public IReadOnlyList<RebarTagLeaderBendingDetailCheckItem> Check(Document doc, Rebar rebar, View view)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (rebar == null) throw new ArgumentNullException(nameof(rebar));
            if (view == null) throw new ArgumentNullException(nameof(view));

            var tagIds = CollectBendingDetailTagIds(doc, rebar, view);
            var bendingDetails = CollectBendingDetailElementsInView(doc, view);
            return BuildResults(doc, tagIds, bendingDetails);
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

        private IReadOnlyList<RebarTagLeaderBendingDetailCheckItem> BuildResults(
            Document doc,
            IReadOnlyList<ElementId> tagIds,
            IReadOnlyList<Element> bendingDetails)
        {
            var results = new List<RebarTagLeaderBendingDetailCheckItem>();
            if (tagIds == null || tagIds.Count == 0)
            {
                return results;
            }

            foreach (var tagId in tagIds)
            {
                var item = BuildResultForTag(doc, tagId, bendingDetails);
                results.Add(item);
            }

            return results;
        }

        private RebarTagLeaderBendingDetailCheckItem BuildResultForTag(
            Document doc,
            ElementId tagId,
            IReadOnlyList<Element> bendingDetails)
        {
            try
            {
                var tag = doc.GetElement(tagId) as IndependentTag;
                if (tag == null)
                {
                    return CreateFail(tagId, ElementId.InvalidElementId, null, "タグが見つかりません。");
                }

                var (hasTaggedRebar, taggedRebarId) = TryResolveTaggedRebarId(doc, tag);
                if (!hasTaggedRebar)
                {
                    return CreateFail(tagId, taggedRebarId, null, "タグに紐づく鉄筋を取得できません。");
                }

                var (hasLeaderEnd, leaderEnd) = TryGetLeaderEnd(tag);
                if (!hasLeaderEnd)
                {
                    return CreateFail(tagId, taggedRebarId, leaderEnd, "LeaderEnd を取得できません（リーダー無し、または無効）。");
                }

                var (hasNearest, nearest, dist) = TryFindNearestBendingDetail(bendingDetails, leaderEnd);
                if (!hasNearest || nearest == null)
                {
                    return CreateFail(tagId, taggedRebarId, leaderEnd, "矢印が指す曲げ詳細を特定できません。");
                }

                var (hasHost, hostRebarId) = TryResolveHostRebarIdFromBendingDetail(doc, nearest);
                if (!hasHost)
                {
                    return new RebarTagLeaderBendingDetailCheckItem(
                        tagId,
                        taggedRebarId,
                        leaderEnd,
                        nearest.Id,
                        ElementId.InvalidElementId,
                        false,
                        "曲げ詳細から鉄筋を特定できません。");
                }

                if (hostRebarId == taggedRebarId)
                {
                    return new RebarTagLeaderBendingDetailCheckItem(
                        tagId,
                        taggedRebarId,
                        leaderEnd,
                        nearest.Id,
                        hostRebarId,
                        true,
                        $"一致（距離={FormatLength(doc, dist)}）。");
                }

                return new RebarTagLeaderBendingDetailCheckItem(
                    tagId,
                    taggedRebarId,
                    leaderEnd,
                    nearest.Id,
                    hostRebarId,
                    false,
                    $"不一致（距離={FormatLength(doc, dist)}）。");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return CreateFail(tagId, ElementId.InvalidElementId, null, ex.Message);
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

                // 複数参照タグの場合もあるので、最初の要素を採用する（必要なら後で拡張）
                var ids = tag.GetTaggedLocalElementIds();
                if (ids == null || ids.Count == 0)
                {
                    return (false, ElementId.InvalidElementId);
                }

                var id = ids.FirstOrDefault();
                if (id == null || id == ElementId.InvalidElementId)
                {
                    return (false, ElementId.InvalidElementId);
                }

                var e = doc.GetElement(id);
                if (e is Rebar)
                {
                    return (true, id);
                }

                return (false, id);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, ElementId.InvalidElementId);
            }
        }

        private (bool ok, XYZ leaderEnd) TryGetLeaderEnd(IndependentTag tag)
        {
            try
            {
                if (tag == null)
                {
                    return (false, null);
                }

                if (!tag.HasLeader)
                {
                    return (false, null);
                }

                var (hasRef, taggedRef) = TryGetFirstTaggedReference(tag);
                if (!hasRef || taggedRef == null)
                {
                    return (false, null);
                }

                // Revit 2026: LeaderEnd はプロパティではなく GetLeaderEnd(Reference)
                var end = tag.GetLeaderEnd(taggedRef);
                return (end != null, end);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null);
            }
        }

        private (bool ok, Reference reference) TryGetFirstTaggedReference(IndependentTag tag)
        {
            try
            {
                if (tag == null)
                {
                    return (false, null);
                }

                // API バージョン差を吸収するため、まずは GetTaggedReferences() を反射で取得する
                var m = tag.GetType().GetMethod("GetTaggedReferences", Type.EmptyTypes);
                if (m == null)
                {
                    return (false, null);
                }

                var obj = m.Invoke(tag, null);
                if (obj is IEnumerable<Reference> refs)
                {
                    var first = refs.FirstOrDefault();
                    return (first != null, first);
                }

                if (obj is IList<Reference> list)
                {
                    var first = list.FirstOrDefault();
                    return (first != null, first);
                }

                return (false, null);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null);
            }
        }

        private (bool ok, Element element, double distance) TryFindNearestBendingDetail(
            IReadOnlyList<Element> bendingDetails,
            XYZ leaderEnd)
        {
            try
            {
                if (bendingDetails == null || bendingDetails.Count == 0 || leaderEnd == null)
                {
                    return (false, null, double.MaxValue);
                }

                // リーダー端点付近の曲げ詳細を最近傍で決める（半径は 50mm）
                var radius = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Millimeters);
                Element best = null;
                var bestDist = double.MaxValue;

                foreach (var e in bendingDetails)
                {
                    var (ok, dist) = TryDistanceToBoundingBox(e, leaderEnd);
                    if (!ok)
                    {
                        continue;
                    }

                    if (dist < bestDist)
                    {
                        bestDist = dist;
                        best = e;
                    }
                }

                if (best == null || bestDist > radius)
                {
                    return (false, null, bestDist);
                }

                return (true, best, bestDist);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, null, double.MaxValue);
            }
        }

        private (bool ok, double distance) TryDistanceToBoundingBox(Element element, XYZ point)
        {
            try
            {
                if (element == null || point == null)
                {
                    return (false, double.MaxValue);
                }

                var bb = element.get_BoundingBox(null);
                if (bb == null || bb.Min == null || bb.Max == null)
                {
                    return (false, double.MaxValue);
                }

                var closest = ClampPointToBox(point, bb.Min, bb.Max);
                return (true, closest.DistanceTo(point));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, double.MaxValue);
            }
        }

        private XYZ ClampPointToBox(XYZ p, XYZ min, XYZ max)
        {
            // ボックス外の点でも最近傍点を求める
            var x = Clamp(p.X, min.X, max.X);
            var y = Clamp(p.Y, min.Y, max.Y);
            var z = Clamp(p.Z, min.Z, max.Z);
            return new XYZ(x, y, z);
        }

        private double Clamp(double v, double min, double max)
        {
            if (v < min) return min;
            if (v > max) return max;
            return v;
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

        private string FormatLength(Document doc, double internalLength)
        {
            try
            {
                if (doc == null)
                {
                    return internalLength.ToString("0.###");
                }

                return UnitFormatUtils.Format(doc.GetUnits(), SpecTypeId.Length, internalLength, false);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return internalLength.ToString("0.###");
            }
        }
    }
}

