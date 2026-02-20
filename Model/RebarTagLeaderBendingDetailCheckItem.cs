using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.Model
{
    internal class RebarTagLeaderBendingDetailCheckItem
    {
        public ElementId TagId { get; }
        public ElementId TaggedRebarId { get; }
        public XYZ LeaderEnd { get; }
        public ElementId PointedBendingDetailId { get; }

        /// <summary>曲げ詳細経由で特定した鉄筋ID（グループ④用）</summary>
        public ElementId PointedRebarId { get; }

        public bool IsMatch { get; }
        public string Message { get; }

        /// <summary>引き出し線が鉄筋モデルを直接指している（グループ③）</summary>
        public bool IsLeaderPointingRebarDirectly { get; }

        /// <summary>引き出し線が直接指している鉄筋のID（グループ③用）</summary>
        public ElementId PointedDirectRebarId { get; }

        public RebarTagLeaderBendingDetailCheckItem(
            ElementId tagId,
            ElementId taggedRebarId,
            XYZ leaderEnd,
            ElementId pointedBendingDetailId,
            ElementId pointedRebarId,
            bool isMatch,
            string message,
            bool isLeaderPointingRebarDirectly = false,
            ElementId pointedDirectRebarId = null)
        {
            TagId = tagId ?? ElementId.InvalidElementId;
            TaggedRebarId = taggedRebarId ?? ElementId.InvalidElementId;
            LeaderEnd = leaderEnd;
            PointedBendingDetailId = pointedBendingDetailId ?? ElementId.InvalidElementId;
            PointedRebarId = pointedRebarId ?? ElementId.InvalidElementId;
            IsMatch = isMatch;
            Message = message ?? string.Empty;
            IsLeaderPointingRebarDirectly = isLeaderPointingRebarDirectly;
            PointedDirectRebarId = pointedDirectRebarId ?? ElementId.InvalidElementId;
        }
    }
}
