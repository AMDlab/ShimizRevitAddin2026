using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Autodesk.Revit.DB;
using ShimizRevitAddin2026.Model;
using ShimizRevitAddin2026.UI.Models;

namespace ShimizRevitAddin2026.UI.ViewModels
{
    internal class RebarTagCheckViewModel : INotifyPropertyChanged
    {
        // ① タグまたは曲げの詳細がホストされてない
        public ObservableCollection<RebarListItem> Group1Rebars { get; } = new ObservableCollection<RebarListItem>();

        // ② 自由な端点のタグがホストされてない
        public ObservableCollection<RebarListItem> Group2Rebars { get; } = new ObservableCollection<RebarListItem>();

        // ③ 自由な端点のタグが鉄筋モデルを指している（赤=HOST不一致）
        public ObservableCollection<RebarListItem> Group3Rebars { get; } = new ObservableCollection<RebarListItem>();

        // ④ 自由な端点の先にある曲げの詳細を指している（赤=HOST不一致）
        public ObservableCollection<RebarListItem> Group4Rebars { get; } = new ObservableCollection<RebarListItem>();

        public ObservableCollection<RebarTagPairRow> Rows { get; } = new ObservableCollection<RebarTagPairRow>();

        private string _keyword = string.Empty;
        public string Keyword
        {
            get => _keyword;
            set { if (_keyword == value) return; _keyword = value ?? string.Empty; OnPropertyChanged(); }
        }

        private string _ngReasonText = string.Empty;
        public string NgReasonText
        {
            get => _ngReasonText;
            private set { if (_ngReasonText == value) return; _ngReasonText = value ?? string.Empty; OnPropertyChanged(); }
        }

        private string _targetViewText = string.Empty;
        public string TargetViewText
        {
            get => _targetViewText;
            private set { if (_targetViewText == value) return; _targetViewText = value ?? string.Empty; OnPropertyChanged(); }
        }

        private string _targetSheetText = string.Empty;
        public string TargetSheetText
        {
            get => _targetSheetText;
            private set { if (_targetSheetText == value) return; _targetSheetText = value ?? string.Empty; OnPropertyChanged(); }
        }

        private string _targetViewsText = string.Empty;
        public string TargetViewsText
        {
            get => _targetViewsText;
            private set { if (_targetViewsText == value) return; _targetViewsText = value ?? string.Empty; OnPropertyChanged(); }
        }

        private string _rebarCountText = string.Empty;
        public string RebarCountText
        {
            get => _rebarCountText;
            private set { if (_rebarCountText == value) return; _rebarCountText = value ?? string.Empty; OnPropertyChanged(); }
        }

        private string _group1CountText = "0";
        public string Group1CountText
        {
            get => _group1CountText;
            private set { if (_group1CountText == value) return; _group1CountText = value ?? "0"; OnPropertyChanged(); }
        }

        private string _group2CountText = "0";
        public string Group2CountText
        {
            get => _group2CountText;
            private set { if (_group2CountText == value) return; _group2CountText = value ?? "0"; OnPropertyChanged(); }
        }

        private string _group3CountText = "0";
        public string Group3CountText
        {
            get => _group3CountText;
            private set { if (_group3CountText == value) return; _group3CountText = value ?? "0"; OnPropertyChanged(); }
        }

        private string _group4CountText = "0";
        public string Group4CountText
        {
            get => _group4CountText;
            private set { if (_group4CountText == value) return; _group4CountText = value ?? "0"; OnPropertyChanged(); }
        }

        public void SetTargetView(View view)
        {
            var viewName = view == null ? string.Empty : (view.Name ?? string.Empty);
            TargetSheetText = string.Empty;
            TargetViewsText = viewName;
            TargetViewText = viewName;
        }

        public void SetTargetSheetAndViews(ViewSheet sheet, IEnumerable<View> views)
        {
            var sheetNumber = sheet == null ? string.Empty : (sheet.SheetNumber ?? string.Empty);
            var sheetNameOnly = sheet == null ? string.Empty : (sheet.Name ?? string.Empty);
            var sheetName = string.IsNullOrWhiteSpace(sheetNumber)
                ? sheetNameOnly
                : (string.IsNullOrWhiteSpace(sheetNameOnly) ? sheetNumber : $"{sheetNumber} {sheetNameOnly}");
            var viewNames = views == null
                ? new List<string>()
                : views.Where(v => v != null).Select(v => v.Name ?? string.Empty).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();

            if (viewNames.Count == 0)
            {
                TargetSheetText = sheetName;
                TargetViewsText = string.Empty;
                TargetViewText = sheetName;
                return;
            }

            var joined = string.Join(" / ", viewNames);
            TargetSheetText = sheetName;
            TargetViewsText = joined;
            TargetViewText = string.IsNullOrWhiteSpace(sheetName) ? joined : $"{sheetName}  {joined}";
        }

        public void SetTargetSheetsAndViewCount(IReadOnlyList<ViewSheet> sheets, int viewCount)
        {
            var sheetCount = sheets == null ? 0 : sheets.Count;
            var sheetSummary = BuildSheetSummaryText(sheets, 3);
            TargetSheetText = string.IsNullOrWhiteSpace(sheetSummary) ? $"シート数：{sheetCount}" : $"{sheetSummary}（{sheetCount}）";
            TargetViewsText = viewCount < 0 ? "ビュー数：0" : $"ビュー数：{viewCount}";
            TargetViewText = $"{TargetSheetText}  {TargetViewsText}";
        }

        private string BuildSheetSummaryText(IReadOnlyList<ViewSheet> sheets, int maxCount)
        {
            if (sheets == null || sheets.Count == 0) return string.Empty;
            if (maxCount < 1) maxCount = 1;

            var names = sheets
                .Where(s => s != null)
                .Take(maxCount)
                .Select(BuildSheetDisplayText)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            if (names.Count == 0) return string.Empty;
            var joined = string.Join(" / ", names);
            return sheets.Count <= maxCount ? joined : $"{joined} / ...";
        }

        private string BuildSheetDisplayText(ViewSheet sheet)
        {
            if (sheet == null) return string.Empty;
            var no = sheet.SheetNumber ?? string.Empty;
            var name = sheet.Name ?? string.Empty;
            if (string.IsNullOrWhiteSpace(no)) return name;
            if (string.IsNullOrWhiteSpace(name)) return no;
            return $"{no} {name}";
        }

        public void SetRebarCount(int count)
        {
            if (count < 0) count = 0;
            RebarCountText = count.ToString();
        }

        public void SetRebars(IEnumerable<RebarListItem> items)
        {
            ClearRebarGroups();
            if (items == null)
            {
                UpdateGroupCounts();
                return;
            }

            foreach (var item in BuildOrderedRebarItems(items))
            {
                AddToGroup(item);
            }

            UpdateGroupCounts();
        }

        private IEnumerable<RebarListItem> BuildOrderedRebarItems(IEnumerable<RebarListItem> items)
        {
            return items
                .OrderBy(GetSafeDisplayText)
                .ThenBy(GetSafeRebarIdValue);
        }

        private void ClearRebarGroups()
        {
            Group1Rebars.Clear();
            Group2Rebars.Clear();
            Group3Rebars.Clear();
            Group4Rebars.Clear();
        }

        private void AddToGroup(RebarListItem item)
        {
            if (item == null) return;

            // ① タグまたは曲げの詳細がホストされてない
            if (item.IsNoTagOrBendingDetail)
            {
                Group1Rebars.Add(item);
                return;
            }

            // ② 自由な端点のタグがホストされてない
            if (item.IsFreeEndTagNotFound)
            {
                Group2Rebars.Add(item);
                return;
            }

            // ③ 自由な端点のタグが鉄筋モデルを指している（赤=HOST不一致）
            if (item.IsLeaderPointingRebar)
            {
                Group3Rebars.Add(item);
                return;
            }

            // ④ 自由な端点の先にある曲げの詳細を指している（赤=HOST不一致）
            Group4Rebars.Add(item);
        }

        private void UpdateGroupCounts()
        {
            Group1CountText = Group1Rebars.Count.ToString();
            Group2CountText = Group2Rebars.Count.ToString();
            Group3CountText = Group3Rebars.Count.ToString();
            Group4CountText = Group4Rebars.Count.ToString();
        }

        private string GetSafeDisplayText(RebarListItem item)
        {
            return item == null ? string.Empty : (item.DisplayText ?? string.Empty);
        }

        private long GetSafeRebarIdValue(RebarListItem item)
        {
            if (item == null || item.RebarId == null) return long.MaxValue;
            return item.RebarId.Value;
        }

        public void UpdateRowsFromModel(RebarTag model)
        {
            if (model == null) { ClearRows(); return; }
            var left = ToDisplayStrings(model.StructureTagIds);
            var right = ToDisplayStrings(model.BendingDetailTagIds);
            UpdateRows(left, right);
        }

        public void UpdateRows(IReadOnlyList<string> structure, IReadOnlyList<string> bending)
        {
            ClearRows();
            foreach (var row in BuildRows(structure, bending))
                Rows.Add(row);
        }

        public void UpdateNgReason(string text)
        {
            NgReasonText = text ?? string.Empty;
        }

        private void ClearRows()
        {
            Rows.Clear();
        }

        private IReadOnlyList<string> ToDisplayStrings(IReadOnlyList<ElementId> ids)
        {
            if (ids == null) return new List<string>();
            return ids.Select(x => x == null ? string.Empty : x.Value.ToString()).ToList();
        }

        private IReadOnlyList<RebarTagPairRow> BuildRows(IReadOnlyList<string> structure, IReadOnlyList<string> bending)
        {
            var result = new List<RebarTagPairRow>();
            var max = new[] { structure?.Count ?? 0, bending?.Count ?? 0 }.Max();

            for (int i = 0; i < max; i++)
            {
                var s = structure != null && i < structure.Count ? structure[i] : string.Empty;
                var b = bending != null && i < bending.Count ? bending[i] : string.Empty;
                result.Add(new RebarTagPairRow(s, b));
            }

            if (result.Count == 0)
                result.Add(new RebarTagPairRow(string.Empty, string.Empty));

            return result;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
