using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.UI.Models
{
    internal class RebarListItem
    {
        public ElementId RebarId { get; }
        public string DisplayText { get; }
        public bool IsLeaderMismatch { get; }

        public RebarListItem(ElementId rebarId, string displayText)
            : this(rebarId, displayText, false)
        {
        }

        public RebarListItem(ElementId rebarId, string displayText, bool isLeaderMismatch)
        {
            RebarId = rebarId;
            DisplayText = displayText ?? string.Empty;
            IsLeaderMismatch = isLeaderMismatch;
        }

        public override string ToString()
        {
            return DisplayText;
        }
    }
}

