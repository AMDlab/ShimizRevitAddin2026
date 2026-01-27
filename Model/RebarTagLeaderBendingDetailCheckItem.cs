using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.Model
{
    internal class RebarTagLeaderBendingDetailCheckItem
    {
        public ElementId TagId { get; }
        public ElementId TaggedRebarId { get; }
        public XYZ LeaderEnd { get; }
        public ElementId PointedBendingDetailId { get; }
        public ElementId PointedRebarId { get; }
        public bool IsMatch { get; }
        public string Message { get; }

        public RebarTagLeaderBendingDetailCheckItem(
            ElementId tagId,
            ElementId taggedRebarId,
            XYZ leaderEnd,
            ElementId pointedBendingDetailId,
            ElementId pointedRebarId,
            bool isMatch,
            string message)
        {
            TagId = tagId ?? ElementId.InvalidElementId;
            TaggedRebarId = taggedRebarId ?? ElementId.InvalidElementId;
            LeaderEnd = leaderEnd;
            PointedBendingDetailId = pointedBendingDetailId ?? ElementId.InvalidElementId;
            PointedRebarId = pointedRebarId ?? ElementId.InvalidElementId;
            IsMatch = isMatch;
            Message = message ?? string.Empty;
        }
    }
}

