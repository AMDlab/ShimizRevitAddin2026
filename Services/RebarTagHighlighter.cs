using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using ShimizRevitAddin2026.Model;

namespace ShimizRevitAddin2026.Services
{
    internal class RebarTagHighlighter
    {
        private readonly RebarDependentTagCollector _collector;

        public RebarTagHighlighter(RebarDependentTagCollector collector)
        {
            _collector = collector;
        }

        public RebarTag Highlight(UIDocument uidoc, Rebar rebar, View activeView)
        {
            return Highlight(uidoc, rebar, activeView, new List<ElementId>());
        }

        public RebarTag Highlight(UIDocument uidoc, Rebar rebar, View activeView, IEnumerable<ElementId> extraElementIds)
        {
            if (uidoc == null) throw new ArgumentNullException(nameof(uidoc));
            if (rebar == null) throw new ArgumentNullException(nameof(rebar));
            if (activeView == null) throw new ArgumentNullException(nameof(activeView));

            // シート上から実行した場合でも、対象Viewに切り替えてから表示する
            TrySetActiveView(uidoc, activeView);

            var model = CollectModel(uidoc.Document, rebar, activeView);
            var idsToHighlight = BuildSelectionIds(rebar.Id, model.MatchedTagIds, extraElementIds);

            uidoc.Selection.SetElementIds(idsToHighlight);
            uidoc.ShowElements(idsToHighlight);

            return model;
        }

        private void TrySetActiveView(UIDocument uidoc, View view)
        {
            if (uidoc == null || view == null)
            {
                return;
            }

            try
            {
                uidoc.ActiveView = view;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private RebarTag CollectModel(Document doc, Rebar rebar, View activeView)
        {
            if (_collector == null)
            {
                // 依存関係が未設定の場合は空結果を返す
                return new RebarTag(rebar.Id, activeView.Id, new List<ElementId>(), new List<ElementId>(), new List<ElementId>());
            }

            return _collector.Collect(doc, rebar, activeView);
        }

        private ICollection<ElementId> BuildSelectionIds(
            ElementId rebarId,
            IReadOnlyList<ElementId> tagIds,
            IEnumerable<ElementId> extraElementIds)
        {
            var result = new List<ElementId> { rebarId };
            if (tagIds == null)
            {
                return result;
            }

            result.AddRange(tagIds);
            if (extraElementIds != null)
            {
                result.AddRange(extraElementIds.Where(x => x != null && x != ElementId.InvalidElementId));
            }
            return result.Distinct().ToList();
        }
    }
}

