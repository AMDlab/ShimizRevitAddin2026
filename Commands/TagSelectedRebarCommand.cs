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
    /// <summary>
    /// ビュー内で鉄筋を複数選択し、「Tag rebar_2」族タイプ「標準」で端部にタグを付けるコマンド。
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class TagSelectedRebarCommand : IExternalCommand
    {
        public Result Execute(
     ExternalCommandData commandData,
     ref string message,
     ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            IList<Reference> pickedRefs;

            try
            {
                pickedRefs = uidoc.Selection.PickObjects(
                    Autodesk.Revit.UI.Selection.ObjectType.Subelement,
                    "TABで個別バーを選択してください");
            }
            catch
            {
                return Result.Cancelled;
            }

            var selection = new List<(Reference Reference, Rebar Rebar)>();

            foreach (var r in pickedRefs)
            {
                var rebar = doc.GetElement(r.ElementId) as Rebar;
                if (rebar == null) continue;

                selection.Add((r, rebar));
            }

            var service = new RebarTagBySelectionService();

            int taggedCount;
            string errorMessage;

            using (Transaction tx = new Transaction(doc, "Tag Rebars"))
            {
                tx.Start();

                var result = service.TagRebarsInView(doc, view, selection);
                taggedCount = result.taggedCount;
                errorMessage = result.errorMessage;

                tx.Commit();
            }

            TaskDialog.Show(
                "Rebar Tag",
                "タグ作成数: " + taggedCount +
                (string.IsNullOrEmpty(errorMessage) ? "" : "\n" + errorMessage));

            return Result.Succeeded;
        }
    }
}
