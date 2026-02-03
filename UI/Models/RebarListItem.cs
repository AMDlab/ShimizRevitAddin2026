using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.UI.Models
{
    internal class RebarListItem
    {
        public ElementId RebarId { get; }
        public ElementId ViewId { get; }
        public string DisplayText { get; }
        public bool IsLeaderMismatch { get; }
        public bool IsNoTagAndNoBendingDetail { get; }

        public RebarListItem(ElementId rebarId, ElementId viewId, string displayText)
            : this(rebarId, viewId, displayText, false, false)
        {
        }

        public RebarListItem(ElementId rebarId, ElementId viewId, string displayText, bool isLeaderMismatch)
            : this(rebarId, viewId, displayText, isLeaderMismatch, false)
        {
        }

        public RebarListItem(ElementId rebarId, ElementId viewId, string displayText, bool isLeaderMismatch, bool isNoTagAndNoBendingDetail)
        {
            RebarId = rebarId;
            ViewId = viewId ?? ElementId.InvalidElementId;
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

