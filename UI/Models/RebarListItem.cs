using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.UI.Models
{
    internal class RebarListItem
    {
        public ElementId RebarId { get; }
        public ElementId ViewId { get; }
        public string DisplayText { get; }

        /// <summary>①-1 構造鉄筋タグが見つからない</summary>
        public bool IsStructureTagNotFound { get; }

        /// <summary>①-2 曲げ詳細が見つからない</summary>
        public bool IsBendingDetailNotFound { get; }

        /// <summary>① タグまたは曲げの詳細がホストされてない（鉄筋モデルに曲げ詳細・タグが存在しない）</summary>
        public bool IsNoTagOrBendingDetail { get; }

        /// <summary>② 自由な端点のタグがホストされてない（有効な自由端タグが見つからない）</summary>
        public bool IsFreeEndTagNotFound { get; }

        /// <summary>③ 自由な端点のタグが鉄筋モデルを指している</summary>
        public bool IsLeaderPointingRebar { get; }

        /// <summary>③ の中でHOST不一致（赤表示）</summary>
        public bool IsLeaderPointingRebarMismatch { get; }

        /// <summary>④ 自由な端点の先にある曲げの詳細を指している</summary>
        public bool IsLeaderPointingBendingDetail { get; }

        /// <summary>④ の中でHOST不一致（赤表示）</summary>
        public bool IsLeaderPointingBendingDetailMismatch { get; }

        public RebarListItem(ElementId rebarId, ElementId viewId, string displayText)
            : this(rebarId, viewId, displayText, false, false, false, false, false, false)
        {
        }

        public RebarListItem(
            ElementId rebarId,
            ElementId viewId,
            string displayText,
            bool isNoTagOrBendingDetail,
            bool isFreeEndTagNotFound,
            bool isLeaderPointingRebar,
            bool isLeaderPointingRebarMismatch,
            bool isLeaderPointingBendingDetail,
            bool isLeaderPointingBendingDetailMismatch)
            : this(
                rebarId,
                viewId,
                displayText,
                isStructureTagNotFound: false,
                isBendingDetailNotFound: false,
                isNoTagOrBendingDetail: isNoTagOrBendingDetail,
                isFreeEndTagNotFound: isFreeEndTagNotFound,
                isLeaderPointingRebar: isLeaderPointingRebar,
                isLeaderPointingRebarMismatch: isLeaderPointingRebarMismatch,
                isLeaderPointingBendingDetail: isLeaderPointingBendingDetail,
                isLeaderPointingBendingDetailMismatch: isLeaderPointingBendingDetailMismatch)
        {
        }

        public RebarListItem(
            ElementId rebarId,
            ElementId viewId,
            string displayText,
            bool isStructureTagNotFound,
            bool isBendingDetailNotFound,
            bool isNoTagOrBendingDetail,
            bool isFreeEndTagNotFound,
            bool isLeaderPointingRebar,
            bool isLeaderPointingRebarMismatch,
            bool isLeaderPointingBendingDetail,
            bool isLeaderPointingBendingDetailMismatch)
        {
            RebarId = rebarId;
            ViewId = viewId ?? ElementId.InvalidElementId;
            DisplayText = displayText ?? string.Empty;
            IsStructureTagNotFound = isStructureTagNotFound;
            IsBendingDetailNotFound = isBendingDetailNotFound;
            IsNoTagOrBendingDetail = isNoTagOrBendingDetail;
            IsFreeEndTagNotFound = isFreeEndTagNotFound;
            IsLeaderPointingRebar = isLeaderPointingRebar;
            IsLeaderPointingRebarMismatch = isLeaderPointingRebarMismatch;
            IsLeaderPointingBendingDetail = isLeaderPointingBendingDetail;
            IsLeaderPointingBendingDetailMismatch = isLeaderPointingBendingDetailMismatch;
        }

        public override string ToString()
        {
            return DisplayText;
        }
    }
}
