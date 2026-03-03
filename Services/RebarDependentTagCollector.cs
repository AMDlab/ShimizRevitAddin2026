using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using ShimizRevitAddin2026.Model;

namespace ShimizRevitAddin2026.Services
{
    internal class RebarDependentTagCollector
    {
        private readonly RebarTagNameMatcher _nameMatcher;

        public RebarDependentTagCollector(RebarTagNameMatcher nameMatcher)
        {
            _nameMatcher = nameMatcher;
        }

        public RebarTag Collect(Document doc, Rebar rebar, View activeView)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (rebar == null) throw new ArgumentNullException(nameof(rebar));
            if (activeView == null) throw new ArgumentNullException(nameof(activeView));

            // 鉄筋に関連付けられた注釈要素を取得して、アクティブビューに限定する
            var dependentIds = GetDependentAnnotationIds(rebar);
            var elementsInActiveView = GetAnnotationElementsInView(doc, dependentIds, activeView.Id);
            var result = BuildResult(doc, rebar.Id, activeView.Id, elementsInActiveView);

            // GetDependentElements で構造タグが0件のとき、ビュー内タグを走査して「この鉄筋を付けたタグ」を逆引きする（Leader オフ等で dependent に含まれないケース対策）
            if (result != null && result.StructureTagIds != null && result.StructureTagIds.Count == 0)
            {
                var fallbackStructureTags = CollectStructureTagsByScanningView(doc, rebar, activeView);
                if (fallbackStructureTags != null && fallbackStructureTags.Count > 0)
                {
                    var mergedMatched = new List<ElementId>(result.MatchedTagIds ?? Array.Empty<ElementId>());
                    mergedMatched.AddRange(fallbackStructureTags);
                    var mergedStructure = new List<ElementId>(fallbackStructureTags);
                    var bending = result.BendingDetailTagIds ?? new List<ElementId>();
                    result = new RebarTag(rebar.Id, activeView.Id, mergedMatched, mergedStructure, bending);
                }
            }

            return result;
        }

        /// <summary>
        /// ビュー内の構造鉄筋タグを走査し、当該鉄筋を TaggedLocal に持つタグの ID を返す。GetDependentElements で拾えない場合の補完用。
        /// 鉄筋が Set（容器）内の 1 本の場合、タグは Set 側に付くため ContainerId も照合する。
        /// </summary>
        private IReadOnlyList<ElementId> CollectStructureTagsByScanningView(Document doc, Rebar rebar, View activeView)
        {
            if (doc == null || rebar == null || activeView == null || _nameMatcher == null)
                return new List<ElementId>();

            try
            {
                var rebarIdsToMatch = GetRebarIdsForTagMatching(rebar);
                if (rebarIdsToMatch == null || rebarIdsToMatch.Count == 0) return new List<ElementId>();

                var tagsInView = new FilteredElementCollector(doc, activeView.Id)
                    .OfClass(typeof(IndependentTag))
                    .Cast<IndependentTag>()
                    .Where(t => t != null)
                    .ToList();

                var result = new List<ElementId>();
                var matchSet = new HashSet<long>(rebarIdsToMatch.Where(id => id != null && id != ElementId.InvalidElementId).Select(id => id.Value));
                foreach (var tag in tagsInView)
                {
                    if (!_nameMatcher.IsStructureRebarTag(doc, tag)) continue;
                    var taggedIds = tag.GetTaggedLocalElementIds();
                    if (taggedIds == null) continue;
                    if (taggedIds.Any(id => id != null && id != ElementId.InvalidElementId && matchSet.Contains(id.Value)))
                    {
                        result.Add(tag.Id);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ElementId>();
            }
        }

        /// <summary>
        /// タグ照合用に、この鉄筋に対応しうる ElementId のリストを返す。自身の Id に加え、Set の場合は ContainerId を含める。
        /// </summary>
        private IReadOnlyList<ElementId> GetRebarIdsForTagMatching(Rebar rebar)
        {
            if (rebar == null) return new List<ElementId>();
            var list = new List<ElementId>();
            var selfId = rebar.Id;
            if (selfId != null && selfId != ElementId.InvalidElementId)
                list.Add(selfId);

            var containerId = TryGetRebarContainerId(rebar);
            if (containerId != null && containerId != ElementId.InvalidElementId && !list.Any(id => id.Value == containerId.Value))
                list.Add(containerId);

            return list;
        }

        /// <summary>
        /// Rebar が Set（容器）内の 1 本の場合、その容器の ElementId を返す。該当しない場合は InvalidElementId。
        /// </summary>
        private ElementId TryGetRebarContainerId(Rebar rebar)
        {
            try
            {
                if (rebar == null) return ElementId.InvalidElementId;
                // Rebar が Set 内の Bar の場合、GetRebarContainerId() で容器 ID が取れる（Revit バージョンによる）
                var method = rebar.GetType().GetMethod("GetRebarContainerId", Type.EmptyTypes);
                if (method == null) return ElementId.InvalidElementId;
                var id = method.Invoke(rebar, null);
                return id as ElementId ?? ElementId.InvalidElementId;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return ElementId.InvalidElementId;
            }
        }

        private IReadOnlyList<ElementId> GetDependentAnnotationIds(Element element)
        {
            try
            {
                var filter = BuildDependentFilter();
                return element.GetDependentElements(filter).ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ElementId>();
            }
        }

        private ElementFilter BuildDependentFilter()
        {
            // IndependentTag と MultiReferenceAnnotation を対象にする
            var filters = new List<ElementFilter>
            {
                new ElementClassFilter(typeof(IndependentTag)),
                new ElementClassFilter(typeof(MultiReferenceAnnotation)),
            };

            return new LogicalOrFilter(filters);
        }

        private IReadOnlyList<Element> GetAnnotationElementsInView(
            Document doc,
            IReadOnlyList<ElementId> dependentIds,
            ElementId viewId)
        {
            var result = new List<Element>();

            if (dependentIds == null)
            {
                return result;
            }

            foreach (var id in dependentIds)
            {
                var e = doc.GetElement(id);
                if (e == null)
                {
                    continue;
                }

                var (ok, ownerViewId) = TryGetOwnerViewId(e);
                if (!ok || ownerViewId != viewId)
                {
                    continue;
                }

                result.Add(e);
            }

            return result;
        }

        private (bool ok, ElementId viewId) TryGetOwnerViewId(Element e)
        {
            try
            {
                if (e is IndependentTag tag)
                {
                    return (true, tag.OwnerViewId);
                }

                if (e is MultiReferenceAnnotation mra)
                {
                    return (true, mra.OwnerViewId);
                }

                return (false, ElementId.InvalidElementId);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return (false, ElementId.InvalidElementId);
            }
        }

        private RebarTag BuildResult(
            Document doc,
            ElementId rebarId,
            ElementId viewId,
            IReadOnlyList<Element> tagElements)
        {
            var matched = new List<ElementId>();
            var structureTags = new List<ElementId>();
            var bendingDetailTags = new List<ElementId>();

            if (tagElements != null)
            {
                foreach (var e in tagElements)
                {
                    if (_nameMatcher != null && _nameMatcher.IsStructureRebarTag(doc, e))
                    {
                        structureTags.Add(e.Id);
                        matched.Add(e.Id);
                        continue;
                    }

                    if (_nameMatcher != null && _nameMatcher.IsBendingDetailTag(doc, e))
                    {
                        bendingDetailTags.Add(e.Id);
                        matched.Add(e.Id);
                        continue;
                    }
                }
            }

            return new RebarTag(rebarId, viewId, matched, structureTags, bendingDetailTags);
        }
    }
}

