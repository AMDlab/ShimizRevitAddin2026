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
        public ObservableCollection<RebarListItem> Rebars { get; } = new ObservableCollection<RebarListItem>();
        public ObservableCollection<RebarTagPairRow> Rows { get; } = new ObservableCollection<RebarTagPairRow>();

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

        public void SetTargetView(View view)
        {
            TargetViewText = view == null ? string.Empty : view.Name;
        }

        public void SetRebars(IEnumerable<RebarListItem> items)
        {
            Rebars.Clear();
            if (items == null) return;

            foreach (var item in items)
            {
                Rebars.Add(item);
            }
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

