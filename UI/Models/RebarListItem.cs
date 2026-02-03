using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.UI.Models
{
    internal class RebarListItem
    {
        public ElementId RebarId { get; }
        public string DisplayText { get; }
        public bool IsLeaderMismatch { get; }
        public bool IsNoTagAndNoBendingDetail { get; }

        public RebarListItem(ElementId rebarId, string displayText)
            : this(rebarId, displayText, false, false)
        {
        }

        public RebarListItem(ElementId rebarId, string displayText, bool isLeaderMismatch)
            : this(rebarId, displayText, isLeaderMismatch, false)
        {
        }

        public RebarListItem(ElementId rebarId, string displayText, bool isLeaderMismatch, bool isNoTagAndNoBendingDetail)
        {
            RebarId = rebarId;
            DisplayText = displayText ?? string.Empty;
            IsLeaderMismatch = isLeaderMismatch;
            IsNoTagAndNoBendingDetail = isNoTagAndNoBendingDetail;
        }

        public override string ToString()
        {
            return DisplayText;
        }
    }
}

