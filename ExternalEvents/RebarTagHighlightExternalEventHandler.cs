using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using ShimizRevitAddin2026.Services;
using ShimizRevitAddin2026.UI.Models;

namespace ShimizRevitAddin2026.ExternalEvents
{
    internal class RebarTagHighlightExternalEventHandler : IExternalEventHandler
    {
        private readonly RebarTagHighlighter _highlighter;

        private ElementId _rebarId = ElementId.InvalidElementId;
        private ElementId _viewId = ElementId.InvalidElementId;
        private Action<RebarTagCheckResult> _onCompleted;

        public RebarTagHighlightExternalEventHandler(RebarTagHighlighter highlighter)
        {
            _highlighter = highlighter;
        }

        public void SetRequest(ElementId rebarId, ElementId viewId, Action<RebarTagCheckResult> onCompleted)
        {
            _rebarId = rebarId ?? ElementId.InvalidElementId;
            _viewId = viewId ?? ElementId.InvalidElementId;
            _onCompleted = onCompleted;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                var uidoc = app?.ActiveUIDocument;
                var doc = uidoc?.Document;
                if (uidoc == null || doc == null)
                {
                    NotifyCompleted(new RebarTagCheckResult(new List<string>(), new List<string>()));
                    return;
                }

                var (rebar, view) = ResolveTargets(doc);
                if (rebar == null || view == null)
                {
                    NotifyCompleted(new RebarTagCheckResult(new List<string>(), new List<string>()));
                    return;
                }

                var model = _highlighter.Highlight(uidoc, rebar, view);
                var result = BuildResult(doc, model);
                NotifyCompleted(result);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                try
                {
                    TaskDialog.Show("RebarTag", ex.ToString());
                }
                catch (Exception dialogEx)
                {
                    Debug.WriteLine(dialogEx);
                }

                NotifyCompleted(new RebarTagCheckResult(new List<string>(), new List<string>()));
            }
        }

        public string GetName()
        {
            return "RebarTagHighlightExternalEventHandler";
        }

        private (Rebar rebar, View view) ResolveTargets(Document doc)
        {
            var rebar = doc.GetElement(_rebarId) as Rebar;
            var view = doc.GetElement(_viewId) as View;
            return (rebar, view);
        }

        private RebarTagCheckResult BuildResult(Document doc, ShimizRevitAddin2026.Model.RebarTag model)
        {
            if (doc == null || model == null)
            {
                return new RebarTagCheckResult(new List<string>(), new List<string>());
            }

            var structure = BuildTagContentList(doc, model.StructureTagIds);
            var bending = BuildTagContentList(doc, model.BendingDetailTagIds);
            return new RebarTagCheckResult(structure, bending);
        }

        private IReadOnlyList<string> BuildTagContentList(Document doc, IReadOnlyList<ElementId> ids)
        {
            if (doc == null || ids == null)
            {
                return new List<string>();
            }

            return ids
                .Select(id => BuildTagContent(doc, id))
                .Where(x => x != null)
                .ToList();
        }

        private string BuildTagContent(Document doc, ElementId id)
        {
            if (doc == null || id == null)
            {
                return string.Empty;
            }

            try
            {
                var e = doc.GetElement(id);
                if (e is IndependentTag tag)
                {
                    return NormalizeContent(tag.TagText, id);
                }

                return id.Value.ToString();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return id.Value.ToString();
            }
        }

        private string NormalizeContent(string content, ElementId id)
        {
            if (!string.IsNullOrWhiteSpace(content))
            {
                return content;
            }

            return id == null ? string.Empty : id.Value.ToString();
        }

        private void NotifyCompleted(RebarTagCheckResult result)
        {
            try
            {
                _onCompleted?.Invoke(result);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }
    }
}

