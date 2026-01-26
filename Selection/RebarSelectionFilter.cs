using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Selection;

namespace ShimizRevitAddin2026.Selection
{
    internal class RebarSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is Rebar;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }
}

