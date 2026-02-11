using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using ShimizRevitAddin2026.Selection;
using ShimizRevitAddin2026.Services;

namespace ShimizRevitAddin2026.Commands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class HighlightRebarTagsInActiveViewCommand : IExternalCommand
    {
        private const string BendingDetailTagName = "曲げ加工詳細";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiapp = commandData.Application;
                var uidoc = uiapp.ActiveUIDocument;
                var doc = uidoc.Document;
                var activeView = doc.ActiveView;

                var rebar = PickRebar(uidoc);
                if (rebar == null)
                {
                    return Result.Cancelled;
                }

                var matcher = new RebarTagNameMatcher(BendingDetailTagName);
                var collector = new RebarDependentTagCollector(matcher);
                var model = collector.Collect(doc, rebar, activeView);

                var idsToHighlight = BuildSelectionIds(rebar.Id, model.MatchedTagIds);
                if (idsToHighlight.Count == 1)
                {
                    TaskDialog.Show("RebarTag", "アクティブビュー内で対象タグが見つかりませんでした。");
                    uidoc.Selection.SetElementIds(idsToHighlight);
                    uidoc.ShowElements(idsToHighlight);
                    return Result.Succeeded;
                }

                uidoc.Selection.SetElementIds(idsToHighlight);
                uidoc.ShowElements(idsToHighlight);
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RebarTag", ex.ToString());
                message = ex.Message;
                return Result.Failed;
            }
        }

        private Rebar PickRebar(UIDocument uidoc)
        {
            var filter = new RebarSelectionFilter();
            Reference pickedRef = uidoc.Selection.PickObject(ObjectType.Element, filter, "鉄筋を選択してください。");
            if (pickedRef == null)
            {
                return null;
            }

            return uidoc.Document.GetElement(pickedRef) as Rebar;
        }

        private ICollection<ElementId> BuildSelectionIds(ElementId rebarId, IReadOnlyList<ElementId> tagIds)
        {
            var result = new List<ElementId> { rebarId };
            if (tagIds == null)
            {
                return result;
            }

            result.AddRange(tagIds);
            return result.Distinct().ToList();
        }
    }
}

