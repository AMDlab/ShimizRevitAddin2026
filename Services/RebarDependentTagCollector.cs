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

            var dependentIds = GetDependentTagIds(rebar);
            var tagsInActiveView = GetIndependentTagsInView(doc, dependentIds, activeView.Id);
            return BuildResult(doc, rebar.Id, activeView.Id, tagsInActiveView);
        }

        private IReadOnlyList<ElementId> GetDependentTagIds(Element element)
        {
            try
            {
                var filter = new ElementClassFilter(typeof(IndependentTag));
                return element.GetDependentElements(filter).ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<ElementId>();
            }
        }

        private IReadOnlyList<IndependentTag> GetIndependentTagsInView(
            Document doc,
            IReadOnlyList<ElementId> dependentIds,
            ElementId viewId)
        {
            var result = new List<IndependentTag>();

            if (dependentIds == null)
            {
                return result;
            }

            foreach (var id in dependentIds)
            {
                var tag = doc.GetElement(id) as IndependentTag;
                if (tag == null)
                {
                    continue;
                }

                if (tag.OwnerViewId != viewId)
                {
                    continue;
                }

                result.Add(tag);
            }

            return result;
        }

        private RebarTag BuildResult(
            Document doc,
            ElementId rebarId,
            ElementId viewId,
            IReadOnlyList<IndependentTag> tags)
        {
            var matched = new List<ElementId>();
            var structureTags = new List<ElementId>();
            var bendingDetailTags = new List<ElementId>();

            if (tags != null)
            {
                foreach (var tag in tags)
                {
                    if (_nameMatcher != null && _nameMatcher.IsStructureRebarTag(doc, tag))
                    {
                        structureTags.Add(tag.Id);
                        matched.Add(tag.Id);
                        continue;
                    }

                    if (_nameMatcher != null && _nameMatcher.IsBendingDetailTag(doc, tag))
                    {
                        bendingDetailTags.Add(tag.Id);
                        matched.Add(tag.Id);
                        continue;
                    }
                }
            }

            return new RebarTag(rebarId, viewId, matched, structureTags, bendingDetailTags);
        }
    }
}

