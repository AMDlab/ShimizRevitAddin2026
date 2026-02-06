using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace ShimizRevitAddin2026.UI.Models
{
    internal enum RebarTagCheckExecutionMode
    {
        ActiveView = 0,
        KeywordSheets = 1,
    }

    internal enum RebarTagCheckExecutionTargetKind
    {
        None = 0,
        SingleView = 1,
        SheetWithPlacedViews = 2,
        MultipleSheets = 3,
    }

    internal class RebarTagCheckExecutionResult
    {
        public bool IsSucceeded { get; }
        public string ErrorMessage { get; }

        public RebarTagCheckExecutionTargetKind TargetKind { get; }

        public View TargetView { get; }
        public ViewSheet TargetSheet { get; }
        public IReadOnlyList<View> PlacedViews { get; }

        public IReadOnlyList<ViewSheet> Sheets { get; }
        public int ViewCount { get; }

        public IReadOnlyList<RebarListItem> Rebars { get; }
        public int RebarCount { get; }

        public RebarTagCheckExecutionResult(
            bool isSucceeded,
            string errorMessage,
            RebarTagCheckExecutionTargetKind targetKind,
            View targetView,
            ViewSheet targetSheet,
            IReadOnlyList<View> placedViews,
            IReadOnlyList<ViewSheet> sheets,
            int viewCount,
            IReadOnlyList<RebarListItem> rebars,
            int rebarCount)
        {
            IsSucceeded = isSucceeded;
            ErrorMessage = errorMessage ?? string.Empty;
            TargetKind = targetKind;
            TargetView = targetView;
            TargetSheet = targetSheet;
            PlacedViews = placedViews ?? new List<View>();
            Sheets = sheets ?? new List<ViewSheet>();
            ViewCount = viewCount < 0 ? 0 : viewCount;
            Rebars = rebars ?? new List<RebarListItem>();
            RebarCount = rebarCount < 0 ? 0 : rebarCount;
        }
    }
}

