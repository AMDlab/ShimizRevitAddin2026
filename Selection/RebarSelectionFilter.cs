using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Selection;

namespace ShimizRevitAddin2026.Selection
{
    /// <summary>
    /// 鉄筋（単体）または鉄筋 Set（容器）を選択可能にするフィルター。
    /// Set の場合は後続で組内の 1 本を取得してタグに使う。
    /// </summary>
    internal class RebarSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem == null) return false;
            if (elem is Rebar) return true;
            // 鉄筋 Set（容器）も選択可能にする。カテゴリが鉄筋の要素を許可（Set は同じカテゴリのことが多い）
            var cat = elem.Category;
            if (cat != null && cat.Id != null && cat.Id.Value == (int)BuiltInCategory.OST_Rebar)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }
}

