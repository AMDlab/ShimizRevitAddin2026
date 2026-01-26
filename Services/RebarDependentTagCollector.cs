using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
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
            return BuildResult(doc, rebar.Id, activeView.Id, elementsInActiveView);
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

