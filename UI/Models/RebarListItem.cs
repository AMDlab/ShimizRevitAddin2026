using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.UI.Models
{
    internal class RebarListItem
    {
        public ElementId RebarId { get; }
        public string DisplayText { get; }

        public RebarListItem(ElementId rebarId, string displayText)
        {
            RebarId = rebarId;
            DisplayText = displayText ?? string.Empty;
        }

        public override string ToString()
        {
            return DisplayText;
        }
    }
}

