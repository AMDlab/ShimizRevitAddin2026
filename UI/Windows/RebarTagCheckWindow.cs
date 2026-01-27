using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using ShimizRevitAddin2026.ExternalEvents;
using ShimizRevitAddin2026.Services;
using ShimizRevitAddin2026.UI.Models;
using ShimizRevitAddin2026.UI.ViewModels;
using Wpf.Ui.Controls;
using Wpf.Ui.Markup;

namespace ShimizRevitAddin2026.UI.Windows
{
    internal class RebarTagCheckWindow : FluentWindow
    {
        private readonly UIDocument _uidoc;
        private readonly View _targetView;
        private readonly RebarTagHighlightExternalEventService _externalEventService;
        private readonly RebarTagCheckViewModel _vm;

        private ListBox _rebarListBox;
        private System.Windows.Controls.DataGrid _resultGrid;

        public RebarTagCheckWindow(
            UIDocument uidoc,
            View targetView,
            RebarTagHighlightExternalEventService externalEventService,
            IReadOnlyList<RebarListItem> rebars)
        {
            _uidoc = uidoc;
            _targetView = targetView;
            _externalEventService = externalEventService;

            _vm = new RebarTagCheckViewModel();
            _vm.SetTargetView(targetView);
            _vm.SetRebars(rebars);

            DataContext = _vm;
            InitializeWindow();
            ApplyWpfUiResources();
            BuildLayout();
        }

        private void InitializeWindow()
        {
            Title = "検証結果";
            Width = 1100;
            Height = 650;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        private void ApplyWpfUiResources()
        {
            try
            {
                // Revit環境ではApp.xamlが無いので、明示的にResourceDictionaryを追加する
                AddDictionaryIfMissing<Wpf.Ui.Markup.ControlsDictionary>(new Wpf.Ui.Markup.ControlsDictionary());

                var theme = new Wpf.Ui.Markup.ThemesDictionary
                {
                    Theme = Wpf.Ui.Appearance.ApplicationTheme.Light
                };
                AddDictionaryIfMissing<Wpf.Ui.Markup.ThemesDictionary>(theme);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void AddDictionaryIfMissing<T>(ResourceDictionary dictionary) where T : ResourceDictionary
        {
            if (dictionary == null)
            {
                return;
            }

            if (Resources == null || Resources.MergedDictionaries == null)
            {
                return;
            }

            foreach (var d in Resources.MergedDictionaries)
            {
                if (d is T)
                {
                    return;
                }
            }

            Resources.MergedDictionaries.Add(dictionary);
        }

        private void BuildLayout()
        {
            var root = CreateRootGrid();
            root.Children.Add(CreateHeader());
            root.Children.Add(CreateBody());
            Content = root;
        }

        private System.Windows.Controls.Grid CreateRootGrid()
        {
            var root = new System.Windows.Controls.Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            return root;
        }

        private UIElement CreateHeader()
        {
            var panel = new DockPanel { LastChildFill = true, Margin = new Thickness(16, 12, 16, 8) };

            var label = new System.Windows.Controls.TextBlock
            {
                Text = "対象View：",
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 16,
                Margin = new Thickness(0, 0, 8, 0)
            };
            DockPanel.SetDock(label, Dock.Left);

            var value = new System.Windows.Controls.TextBlock
            {
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 16,
                FontWeight = FontWeights.SemiBold
            };
            value.SetBinding(System.Windows.Controls.TextBlock.TextProperty, new System.Windows.Data.Binding(nameof(RebarTagCheckViewModel.TargetViewText)));

            panel.Children.Add(label);
            panel.Children.Add(value);

            System.Windows.Controls.Grid.SetRow(panel, 0);
            return panel;
        }

        private UIElement CreateBody()
        {
            var grid = new System.Windows.Controls.Grid { Margin = new Thickness(16, 8, 16, 16) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(360) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var left = CreateRebarListCard();
            var right = CreateResultCard();

            System.Windows.Controls.Grid.SetColumn(left, 0);
            System.Windows.Controls.Grid.SetColumn(right, 2);
            grid.Children.Add(left);
            grid.Children.Add(right);

            System.Windows.Controls.Grid.SetRow(grid, 1);
            return grid;
        }

        private UIElement CreateRebarListCard()
        {
            var card = new Card { Padding = new Thickness(12) };
            var stack = new StackPanel();

            stack.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = "鉄筋List",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8)
            });

            _rebarListBox = new ListBox
            {
                MinHeight = 400,
                HorizontalContentAlignment = HorizontalAlignment.Stretch
            };
            _rebarListBox.SetBinding(ItemsControl.ItemsSourceProperty, new System.Windows.Data.Binding(nameof(RebarTagCheckViewModel.Rebars)));
            _rebarListBox.SelectionChanged += OnRebarSelectionChanged;

            stack.Children.Add(_rebarListBox);
            card.Content = stack;
            return card;
        }

        private UIElement CreateResultCard()
        {
            var card = new Card { Padding = new Thickness(12) };
            var stack = new StackPanel();

            stack.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = "検証結果",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8)
            });

            _resultGrid = CreateResultGrid();
            _resultGrid.SetBinding(ItemsControl.ItemsSourceProperty, new System.Windows.Data.Binding(nameof(RebarTagCheckViewModel.Rows)));

            stack.Children.Add(_resultGrid);
            card.Content = stack;
            return card;
        }

        private System.Windows.Controls.DataGrid CreateResultGrid()
        {
            var grid = new System.Windows.Controls.DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                CanUserAddRows = false,
                HeadersVisibility = DataGridHeadersVisibility.Column,
                MinHeight = 400
            };

            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "鉄筋タグ",
                Binding = new System.Windows.Data.Binding(nameof(RebarTagPairRow.StructureRebarTag)),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star)
            });

            grid.Columns.Add(new DataGridTextColumn
            {
                Header = "曲げ詳細の鉄筋タグ",
                Binding = new System.Windows.Data.Binding(nameof(RebarTagPairRow.BendingDetailRebarTag)),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star)
            });

            return grid;
        }

        private void OnRebarSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var item = _rebarListBox?.SelectedItem as RebarListItem;
                if (item == null)
                {
                    return;
                }

                var rebar = ResolveRebar(item.RebarId);
                if (rebar == null)
                {
                    return;
                }

                RequestHighlight(rebar.Id);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // 操作キャンセル
                Debug.WriteLine("OperationCanceledException");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RebarTag", ex.ToString());
            }
        }

        private Rebar ResolveRebar(ElementId rebarId)
        {
            if (_uidoc == null || _uidoc.Document == null)
            {
                return null;
            }

            var e = _uidoc.Document.GetElement(rebarId);
            return e as Rebar;
        }

        private void RequestHighlight(ElementId rebarId)
        {
            if (_externalEventService == null || _targetView == null)
            {
                return;
            }

            var viewId = _targetView.Id;
            _externalEventService.Request(rebarId, viewId, OnHighlightCompleted);
        }

        private void OnHighlightCompleted(RebarTagCheckResult result)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    var structure = result?.Structure ?? new List<string>();
                    var bending = result?.Bending ?? new List<string>();
                    _vm.UpdateRows(structure, bending);
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }
    }
}

