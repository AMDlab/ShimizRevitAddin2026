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
        private readonly RebarTagHighlightExternalEventService _externalEventService;
        private readonly RebarTagCheckViewModel _vm;

        private ListBox _redListBox;
        private ListBox _yellowListBox;
        private ListBox _blueListBox;
        private System.Windows.Controls.DataGrid _resultGrid;

        public RebarTagCheckWindow(
            UIDocument uidoc,
            View targetView,
            RebarTagHighlightExternalEventService externalEventService,
            IReadOnlyList<RebarListItem> rebars,
            int rebarCount)
        {
            _uidoc = uidoc;
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

        public RebarTagCheckWindow(
            UIDocument uidoc,
            ViewSheet sheet,
            IReadOnlyList<View> views,
            RebarTagHighlightExternalEventService externalEventService,
            IReadOnlyList<RebarListItem> rebars,
            int rebarCount)
        {
            _uidoc = uidoc;
            _externalEventService = externalEventService;

            _vm = new RebarTagCheckViewModel();
            _vm.SetTargetSheetAndViews(sheet, views);
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

            root.Children.Add(CreateTargetSheetRow());
            root.Children.Add(CreateTargetViewRow());

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

        private UIElement CreateTargetSheetRow()
        {
            var panel = new DockPanel { LastChildFill = true };
            var label = new System.Windows.Controls.TextBlock
            {
                Text = "対象シート：",
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 16,
                Margin = new Thickness(0, 0, 8, 0)
            };
            DockPanel.SetDock(label, Dock.Left);

            var value = new System.Windows.Controls.TextBlock
            {
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 18,
                FontWeight = FontWeights.SemiBold,
                TextWrapping = TextWrapping.Wrap
            };
            value.SetBinding(System.Windows.Controls.TextBlock.TextProperty, new System.Windows.Data.Binding(nameof(RebarTagCheckViewModel.TargetSheetText)));

            panel.Children.Add(label);
            panel.Children.Add(value);
            return panel;
        }

        private UIElement CreateTargetViewRow()
        {
            var panel = new DockPanel { LastChildFill = true, Margin = new Thickness(0, 4, 0, 0) };
            var label = new System.Windows.Controls.TextBlock
            {
                Text = "対象ビュー：",
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 14,
                Margin = new Thickness(0, 0, 8, 0)
            };
            DockPanel.SetDock(label, Dock.Left);

            var value = new System.Windows.Controls.TextBlock
            {
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 14,
                FontWeight = FontWeights.Normal,
                TextWrapping = TextWrapping.Wrap
            };
            value.SetBinding(System.Windows.Controls.TextBlock.TextProperty, new System.Windows.Data.Binding(nameof(RebarTagCheckViewModel.TargetViewsText)));

            panel.Children.Add(label);
            panel.Children.Add(value);
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
            // 折りたたみ時でも上寄せで表示する
            card.VerticalContentAlignment = VerticalAlignment.Top;
            card.HorizontalContentAlignment = HorizontalAlignment.Stretch;

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

            var body = CreateRebarGroupPanel();
            var scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                VerticalContentAlignment = VerticalAlignment.Top,
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                Content = body
            };
            System.Windows.Controls.Grid.SetRow(scroll, 1);
            grid.Children.Add(scroll);

            card.Content = grid;
            return card;
        }

        private UIElement CreateRebarGroupPanel()
        {
            var panel = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Stretch
            };

            panel.Children.Add(CreateRebarExpander("赤（不一致）", nameof(RebarTagCheckViewModel.RedCountText), out _redListBox, nameof(RebarTagCheckViewModel.RedRebars), Brushes.Red, System.Windows.Media.Color.FromRgb(255, 235, 238)));
            panel.Children.Add(CreateRebarExpander("黄（線分未取得）", nameof(RebarTagCheckViewModel.YellowCountText), out _yellowListBox, nameof(RebarTagCheckViewModel.YellowRebars), Brushes.DarkGoldenrod, System.Windows.Media.Color.FromRgb(255, 249, 196)));
            panel.Children.Add(CreateRebarExpander("青（鉄筋のみ）", nameof(RebarTagCheckViewModel.BlueCountText), out _blueListBox, nameof(RebarTagCheckViewModel.BlueRebars), Brushes.Blue, System.Windows.Media.Color.FromRgb(227, 242, 253)));

            return panel;
        }

        private UIElement CreateRebarExpander(
            string headerText,
            string countBindingName,
            out ListBox listBox,
            string itemsBindingName,
            Brush foreground,
            System.Windows.Media.Color backgroundColor)
        {
            var expander = new Expander
            {
                IsExpanded = true,
                Margin = new Thickness(0, 0, 0, 6)
            };

            var header = new DockPanel { LastChildFill = true };
            var title = new System.Windows.Controls.TextBlock
            {
                Text = headerText,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            };
            DockPanel.SetDock(title, Dock.Left);

            var count = new System.Windows.Controls.TextBlock
            {
                FontSize = 13,
                Foreground = Brushes.Gray,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8, 0, 0, 0)
            };
            count.SetBinding(System.Windows.Controls.TextBlock.TextProperty, new System.Windows.Data.Binding(countBindingName));

            header.Children.Add(title);
            header.Children.Add(count);
            expander.Header = header;

            listBox = new ListBox
            {
                MinHeight = 120,
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch
            };
            listBox.ItemContainerStyle = BuildFixedColorListItemStyle(foreground, backgroundColor);
            ScrollViewer.SetVerticalScrollBarVisibility(listBox, ScrollBarVisibility.Auto);
            ScrollViewer.SetHorizontalScrollBarVisibility(listBox, ScrollBarVisibility.Disabled);
            listBox.SetBinding(ItemsControl.ItemsSourceProperty, new System.Windows.Data.Binding(itemsBindingName));
            listBox.SelectionChanged += OnRebarSelectionChanged;

            expander.Content = listBox;
            return expander;
        }

        private Style BuildFixedColorListItemStyle(Brush foreground, System.Windows.Media.Color backgroundColor)
        {
            var style = new Style(typeof(ListBoxItem));
            style.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, foreground));
            style.Setters.Add(new Setter(System.Windows.Controls.Control.FontWeightProperty, FontWeights.SemiBold));
            style.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(backgroundColor)));
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
                var selected = (sender as ListBox)?.SelectedItem as RebarListItem;
                if (selected == null)
                {
                    return;
                }

                ClearOtherSelections(sender as ListBox);

                var item = selected;
                if (item == null)
                {
                    return;
                }

                var rebar = ResolveRebar(item.RebarId);
                if (rebar == null)
                {
                    return;
                }

                RequestHighlight(rebar.Id, item.ViewId);
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

        private void ClearOtherSelections(ListBox sender)
        {
            // 別のリストで選択が残ると混乱するため、選択元以外は解除する
            try
            {
                if (!ReferenceEquals(sender, _redListBox) && _redListBox != null) _redListBox.SelectedItem = null;
                if (!ReferenceEquals(sender, _yellowListBox) && _yellowListBox != null) _yellowListBox.SelectedItem = null;
                if (!ReferenceEquals(sender, _blueListBox) && _blueListBox != null) _blueListBox.SelectedItem = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
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

        private void RequestHighlight(ElementId rebarId, ElementId viewId)
        {
            if (_externalEventService == null)
            {
                return;
            }

            var safeViewId = viewId ?? ElementId.InvalidElementId;
            if (safeViewId == ElementId.InvalidElementId)
            {
                safeViewId = _uidoc?.ActiveView?.Id ?? ElementId.InvalidElementId;
            }

            _externalEventService.Request(rebarId, safeViewId, OnHighlightCompleted);
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

