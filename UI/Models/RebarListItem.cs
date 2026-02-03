using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.UI.Models
{
    internal class RebarListItem
    {
        public ElementId RebarId { get; }
        public ElementId ViewId { get; }
        public string DisplayText { get; }
        public bool IsBendingDetailMismatch { get; }
        public bool IsLeaderLineNotFound { get; }
        public bool IsNoTagAndNoBendingDetail { get; }

        public RebarListItem(ElementId rebarId, ElementId viewId, string displayText)
            : this(rebarId, viewId, displayText, false, false, false)
        {
        }

        public RebarListItem(ElementId rebarId, ElementId viewId, string displayText, bool isBendingDetailMismatch, bool isLeaderLineNotFound)
            : this(rebarId, viewId, displayText, isBendingDetailMismatch, isLeaderLineNotFound, false)
        {
        }

        public RebarListItem(
            ElementId rebarId,
            ElementId viewId,
            string displayText,
            bool isBendingDetailMismatch,
            bool isLeaderLineNotFound,
            bool isNoTagAndNoBendingDetail)
        {
            RebarId = rebarId;
            ViewId = viewId ?? ElementId.InvalidElementId;
            DisplayText = displayText ?? string.Empty;
            IsBendingDetailMismatch = isBendingDetailMismatch;
            IsLeaderLineNotFound = isLeaderLineNotFound;
            IsNoTagAndNoBendingDetail = isNoTagAndNoBendingDetail;
        }

        public override string ToString()
        {
            return DisplayText;
        }
    }
}

