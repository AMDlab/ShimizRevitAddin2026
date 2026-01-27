using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace ShimizRevitAddin2026.Services
{
    internal class RebarCollectorInView
    {
        public IReadOnlyList<Rebar> Collect(Document doc, View view)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));

            return CollectInView(doc, view.Id);
        }

        private IReadOnlyList<Rebar> CollectInView(Document doc, ElementId viewId)
        {
            // 表示中の鉄筋だけを一覧化する
            var collector = new FilteredElementCollector(doc, viewId);
            return collector
                .OfClass(typeof(Rebar))
                .Cast<Rebar>()
                .OrderBy(r => r.Id.Value)
                .ToList();
        }
    }
}

