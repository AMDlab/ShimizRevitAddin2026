using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
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
    internal class RebarTagCheckWindow : System.Windows.Window
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
            IReadOnlyList<RebarListItem> rebars,
            int rebarCount)
        {
            _uidoc = uidoc;
            _targetView = targetView;
            _externalEventService = externalEventService;

            _vm = new RebarTagCheckViewModel();
            _vm.SetTargetView(targetView);
            _vm.SetRebars(rebars);
            _vm.SetRebarCount(rebarCount);

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

            // 通常のウィンドウ枠を使用して、移動・最小化・閉じるを有効化する
            WindowStyle = WindowStyle.SingleBorderWindow;
            ResizeMode = ResizeMode.CanResize;
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
            var root = new StackPanel { Margin = new Thickness(16, 12, 16, 8) };

            var panel = new DockPanel { LastChildFill = true };
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

            root.Children.Add(panel);

            var countPanel = new DockPanel { LastChildFill = true, Margin = new Thickness(0, 6, 0, 0) };
            var countLabel = new System.Windows.Controls.TextBlock
            {
                Text = "鉄筋数：",
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 14,
                Margin = new Thickness(0, 0, 8, 0)
            };
            DockPanel.SetDock(countLabel, Dock.Left);

            var countValue = new System.Windows.Controls.TextBlock
            {
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold
            };
            countValue.SetBinding(System.Windows.Controls.TextBlock.TextProperty, new System.Windows.Data.Binding(nameof(RebarTagCheckViewModel.RebarCountText)));

            countPanel.Children.Add(countLabel);
            countPanel.Children.Add(countValue);
            root.Children.Add(countPanel);

            System.Windows.Controls.Grid.SetRow(root, 0);
            return root;
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

        private Card CreateCard()
        {
            return new Card
            {
                Padding = new Thickness(12),
                VerticalAlignment = VerticalAlignment.Stretch,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                SnapsToDevicePixels = true
            };
        }

        private UIElement CreateRebarListCard()
        {
            var card = CreateCard();

            var grid = new System.Windows.Controls.Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            var title = new System.Windows.Controls.TextBlock
            {
                Text = "鉄筋List",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8)
            };
            System.Windows.Controls.Grid.SetRow(title, 0);
            grid.Children.Add(title);

            _rebarListBox = new ListBox
            {
                MinHeight = 400,
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };
            _rebarListBox.ItemContainerStyle = BuildRebarListItemStyle();
            ScrollViewer.SetVerticalScrollBarVisibility(_rebarListBox, ScrollBarVisibility.Auto);
            ScrollViewer.SetHorizontalScrollBarVisibility(_rebarListBox, ScrollBarVisibility.Disabled);
            _rebarListBox.SetBinding(ItemsControl.ItemsSourceProperty, new System.Windows.Data.Binding(nameof(RebarTagCheckViewModel.Rebars)));
            _rebarListBox.SelectionChanged += OnRebarSelectionChanged;

            System.Windows.Controls.Grid.SetRow(_rebarListBox, 1);
            grid.Children.Add(_rebarListBox);

            card.Content = grid;
            return card;
        }

        private Style BuildRebarListItemStyle()
        {
            // 不一致がある鉄筋を視覚的に強調する
            var style = new Style(typeof(ListBoxItem));
            style.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, Brushes.Black));
            style.Setters.Add(new Setter(System.Windows.Controls.Control.FontWeightProperty, FontWeights.Normal));

            var noTagNoBendingDetailTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding(nameof(UI.Models.RebarListItem.IsNoTagAndNoBendingDetail)),
                Value = true
            };
            // 構造タグなし・曲げ詳細なし（鉄筋のみ）を青で表示
            noTagNoBendingDetailTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, Brushes.Blue));
            noTagNoBendingDetailTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.FontWeightProperty, FontWeights.SemiBold));
            noTagNoBendingDetailTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(227, 242, 253))));

            var mismatchTrigger = new MultiDataTrigger();
            mismatchTrigger.Conditions.Add(new Condition(new System.Windows.Data.Binding(nameof(UI.Models.RebarListItem.IsLeaderMismatch)), true));
            mismatchTrigger.Conditions.Add(new Condition(new System.Windows.Data.Binding(nameof(UI.Models.RebarListItem.IsNoTagAndNoBendingDetail)), false));
            mismatchTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, Brushes.Red));
            mismatchTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.FontWeightProperty, FontWeights.SemiBold));
            mismatchTrigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 235, 238))));

            style.Triggers.Add(noTagNoBendingDetailTrigger);
            style.Triggers.Add(mismatchTrigger);
            return style;
        }

        private UIElement CreateResultCard()
        {
            var card = CreateCard();

            var grid = new System.Windows.Controls.Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var title = new System.Windows.Controls.TextBlock
            {
                Text = "検証結果",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8)
            };
            System.Windows.Controls.Grid.SetRow(title, 0);
            grid.Children.Add(title);

            _resultGrid = CreateResultGrid();
            _resultGrid.SetBinding(ItemsControl.ItemsSourceProperty, new System.Windows.Data.Binding(nameof(RebarTagCheckViewModel.Rows)));

            System.Windows.Controls.Grid.SetRow(_resultGrid, 1);
            grid.Children.Add(_resultGrid);

            var ngPanel = CreateNgReasonPanel();
            System.Windows.Controls.Grid.SetRow(ngPanel, 2);
            grid.Children.Add(ngPanel);

            card.Content = grid;
            return card;
        }

        private UIElement CreateNgReasonPanel()
        {
            var panel = new System.Windows.Controls.Grid { Margin = new Thickness(0, 12, 0, 0) };
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            panel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(140) });

            var label = new System.Windows.Controls.TextBlock
            {
                Text = "メッセージボックス",
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 6)
            };
            System.Windows.Controls.Grid.SetRow(label, 0);
            panel.Children.Add(label);

            var box = new System.Windows.Controls.TextBox
            {
                IsReadOnly = true,
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
            };
            box.SetBinding(System.Windows.Controls.TextBox.TextProperty, new System.Windows.Data.Binding(nameof(RebarTagCheckViewModel.NgReasonText)));
            System.Windows.Controls.Grid.SetRow(box, 1);
            panel.Children.Add(box);

            return panel;
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
                Header = "曲げ詳細",
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
                    _vm.UpdateNgReason(result?.NgReason ?? string.Empty);
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }
    }
}

