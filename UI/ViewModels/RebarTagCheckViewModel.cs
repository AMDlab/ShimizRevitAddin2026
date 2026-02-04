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
        public ObservableCollection<RebarListItem> RedRebars { get; } = new ObservableCollection<RebarListItem>();
        public ObservableCollection<RebarListItem> YellowRebars { get; } = new ObservableCollection<RebarListItem>();
        public ObservableCollection<RebarListItem> BlueRebars { get; } = new ObservableCollection<RebarListItem>();
        public ObservableCollection<RebarListItem> BlackRebars { get; } = new ObservableCollection<RebarListItem>();
        public ObservableCollection<RebarTagPairRow> Rows { get; } = new ObservableCollection<RebarTagPairRow>();

        private string _ngReasonText = string.Empty;
        public string NgReasonText
        {
            get => _ngReasonText;
            private set
            {
                if (_ngReasonText == value) return;
                _ngReasonText = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        private string _targetViewText = string.Empty;
        public string TargetViewText
        {
            get => _targetViewText;
            private set
            {
                if (_targetViewText == value) return;
                _targetViewText = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        private string _targetSheetText = string.Empty;
        public string TargetSheetText
        {
            get => _targetSheetText;
            private set
            {
                if (_targetSheetText == value) return;
                _targetSheetText = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        private string _targetViewsText = string.Empty;
        public string TargetViewsText
        {
            get => _targetViewsText;
            private set
            {
                if (_targetViewsText == value) return;
                _targetViewsText = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        private string _rebarCountText = string.Empty;
        public string RebarCountText
        {
            get => _rebarCountText;
            private set
            {
                if (_rebarCountText == value) return;
                _rebarCountText = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        private string _redCountText = "0";
        public string RedCountText
        {
            get => _redCountText;
            private set
            {
                if (_redCountText == value) return;
                _redCountText = value ?? "0";
                OnPropertyChanged();
            }
        }

        private string _yellowCountText = "0";
        public string YellowCountText
        {
            get => _yellowCountText;
            private set
            {
                if (_yellowCountText == value) return;
                _yellowCountText = value ?? "0";
                OnPropertyChanged();
            }
        }

        private string _blueCountText = "0";
        public string BlueCountText
        {
            get => _blueCountText;
            private set
            {
                if (_blueCountText == value) return;
                _blueCountText = value ?? "0";
                OnPropertyChanged();
            }
        }

        private string _blackCountText = "0";
        public string BlackCountText
        {
            get => _blackCountText;
            private set
            {
                if (_blackCountText == value) return;
                _blackCountText = value ?? "0";
                OnPropertyChanged();
            }
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
            // 複数シート対象時の表示（長くなりすぎないように件数を主に出す）
            var sheetCount = sheets == null ? 0 : sheets.Count;
            var sheetSummary = BuildSheetSummaryText(sheets, 3);

            TargetSheetText = string.IsNullOrWhiteSpace(sheetSummary) ? $"シート数：{sheetCount}" : $"{sheetSummary}（{sheetCount}）";
            TargetViewsText = viewCount < 0 ? "ビュー数：0" : $"ビュー数：{viewCount}";
            TargetViewText = $"{TargetSheetText}  {TargetViewsText}";
        }

        private string BuildSheetSummaryText(IReadOnlyList<ViewSheet> sheets, int maxCount)
        {
            if (sheets == null || sheets.Count == 0)
            {
                return string.Empty;
            }

            if (maxCount < 1)
            {
                maxCount = 1;
            }

            var names = sheets
                .Where(s => s != null)
                .Take(maxCount)
                .Select(BuildSheetDisplayText)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            if (names.Count == 0)
            {
                return string.Empty;
            }

            var joined = string.Join(" / ", names);
            if (sheets.Count <= maxCount)
            {
                return joined;
            }

            return $"{joined} / ...";
        }

        private string BuildSheetDisplayText(ViewSheet sheet)
        {
            if (sheet == null)
            {
                return string.Empty;
            }

            var no = sheet.SheetNumber ?? string.Empty;
            var name = sheet.Name ?? string.Empty;
            if (string.IsNullOrWhiteSpace(no))
            {
                return name;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                return no;
            }

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
            RedRebars.Clear();
            YellowRebars.Clear();
            BlueRebars.Clear();
            BlackRebars.Clear();
        }

        private void AddToGroup(RebarListItem item)
        {
            if (item == null)
            {
                return;
            }

            // 赤：曲げ詳細を取得できて不一致
            if (item.IsBendingDetailMismatch)
            {
                RedRebars.Add(item);
                return;
            }

            // 黄：自由端タグの線分を取得できない
            if (item.IsLeaderLineNotFound)
            {
                YellowRebars.Add(item);
                return;
            }

            // 青：構造タグなし・曲げ詳細なし（鉄筋のみ）
            if (item.IsNoTagAndNoBendingDetail)
            {
                BlueRebars.Add(item);
                return;
            }

            // 黒：上記以外（一致など）
            BlackRebars.Add(item);
        }

        private void UpdateGroupCounts()
        {
            RedCountText = RedRebars.Count.ToString();
            YellowCountText = YellowRebars.Count.ToString();
            BlueCountText = BlueRebars.Count.ToString();
            BlackCountText = BlackRebars.Count.ToString();
        }

        private string GetSafeDisplayText(RebarListItem item)
        {
            return item == null ? string.Empty : (item.DisplayText ?? string.Empty);
        }

        private long GetSafeRebarIdValue(RebarListItem item)
        {
            if (item == null || item.RebarId == null)
            {
                return long.MaxValue;
            }

            return item.RebarId.Value;
        }

        public void UpdateRowsFromModel(RebarTag model)
        {
            if (model == null)
            {
                ClearRows();
                return;
            }

            var left = ToDisplayStrings(model.StructureTagIds);
            var right = ToDisplayStrings(model.BendingDetailTagIds);

            UpdateRows(left, right);
        }

        public void UpdateRows(IReadOnlyList<string> structure, IReadOnlyList<string> bending)
        {
            ClearRows();

            foreach (var row in BuildRows(structure, bending))
            {
                Rows.Add(row);
            }
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
            {
                // 空表示用の1行
                result.Add(new RebarTagPairRow(string.Empty, string.Empty));
            }

            return result;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

