using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.Services
{
    internal class SheetPlacedViewCollector
    {
        public IReadOnlyList<View> Collect(Document doc, ViewSheet sheet)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));

            return CollectPlacedViews(doc, sheet.Id);
        }

        private IReadOnlyList<View> CollectPlacedViews(Document doc, ElementId sheetId)
        {
            // シートに配置されているViewportからViewを取得する
            try
            {
                var collector = new FilteredElementCollector(doc, sheetId);
                return collector
                    .OfClass(typeof(Viewport))
                    .Cast<Viewport>()
                    .Select(vp => ResolveView(doc, vp?.ViewId))
                    .Where(v => v != null)
                    .GroupBy(v => v.Id)
                    .Select(g => g.First())
                    .OrderBy(v => v.Name ?? string.Empty)
                    .ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return new List<View>();
            }
        }

        private View ResolveView(Document doc, ElementId viewId)
        {
            if (doc == null || viewId == null || viewId == ElementId.InvalidElementId)
            {
                return null;
            }

            return doc.GetElement(viewId) as View;
        }
    }
}

