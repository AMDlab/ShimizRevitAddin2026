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
        private readonly RebarTagCheckExecuteExternalEventService _checkExecuteExternalEventService;
        private readonly RebarTagCheckViewModel _vm;

        private readonly double _leftColumnWidth = 360;

        private readonly View _initialView;

        private System.Windows.Controls.TextBox _keywordTextBox;
        private System.Windows.Controls.Button _executeButton;
        private System.Windows.Controls.Button _resetButton;

        private ListBox _group1ListBox;
        private ListBox _group2ListBox;
        private ListBox _group3ListBox;
        private ListBox _group4ListBox;
        private System.Windows.Controls.DataGrid _resultGrid;

        public RebarTagCheckWindow(
            UIDocument uidoc,
            View initialView,
            RebarTagHighlightExternalEventService externalEventService,
            RebarTagCheckExecuteExternalEventService checkExecuteExternalEventService)
        {
            _uidoc = uidoc;
            _initialView = initialView;
            _externalEventService = externalEventService;
            _checkExecuteExternalEventService = checkExecuteExternalEventService;

            _vm = new RebarTagCheckViewModel();
            _vm.SetTargetView(initialView);
            _vm.SetRebars(new List<RebarListItem>());
            _vm.SetRebarCount(0);

            DataContext = _vm;
            InitializeWindow();
            ApplyWpfUiResources();
            BuildLayout();
        }

        public RebarTagCheckWindow(
            UIDocument uidoc,
            View targetView,
            RebarTagHighlightExternalEventService externalEventService,
            IReadOnlyList<RebarListItem> rebars,
            int rebarCount)
        {
            _uidoc = uidoc;
            _externalEventService = externalEventService;
            _checkExecuteExternalEventService = null;
            _initialView = targetView;

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
            _checkExecuteExternalEventService = null;
            _initialView = sheet;

            _vm = new RebarTagCheckViewModel();
            _vm.SetTargetSheetAndViews(sheet, views);
            _vm.SetRebars(rebars);
            _vm.SetRebarCount(rebarCount);

            DataContext = _vm;
            InitializeWindow();
            ApplyWpfUiResources();
            BuildLayout();
        }

        public RebarTagCheckWindow(
            UIDocument uidoc,
            IReadOnlyList<ViewSheet> sheets,
            int viewCount,
            RebarTagHighlightExternalEventService externalEventService,
            IReadOnlyList<RebarListItem> rebars,
            int rebarCount)
        {
            _uidoc = uidoc;
            _externalEventService = externalEventService;
            _checkExecuteExternalEventService = null;
            _initialView = null;

            _vm = new RebarTagCheckViewModel();
            _vm.SetTargetSheetsAndViewCount(sheets, viewCount);
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

            root.Children.Add(CreateKeywordAndActionRow());

            System.Windows.Controls.Grid.SetRow(root, 0);
            return root;
        }

        private UIElement CreateKeywordAndActionRow()
        {
            // Keywordの幅を左カラムに合わせ、右側にボタンを寄せて配置する
            var outer = new System.Windows.Controls.Grid
            {
                Margin = new Thickness(0, 6, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Left
            };

            var keywordGrid = new System.Windows.Controls.Grid
            {
                Width = _leftColumnWidth,
                HorizontalAlignment = HorizontalAlignment.Left
            };
            keywordGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            keywordGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var label = new System.Windows.Controls.TextBlock
            {
                Text = "Keyword：",
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 14,
                Margin = new Thickness(0, 0, 8, 0)
            };
            System.Windows.Controls.Grid.SetColumn(label, 0);

            var box = new System.Windows.Controls.TextBox
            {
                // フォントが切れないように、余白と高さを少し確保する
                Height = 32,
                MinHeight = 32,
                Padding = new Thickness(6, 3, 6, 3),
                VerticalContentAlignment = VerticalAlignment.Center
            };
            _keywordTextBox = box;
            box.SetBinding(System.Windows.Controls.TextBox.TextProperty, new System.Windows.Data.Binding(nameof(RebarTagCheckViewModel.Keyword))
            {
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            });
            System.Windows.Controls.Grid.SetColumn(box, 1);

            keywordGrid.Children.Add(label);
            keywordGrid.Children.Add(box);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(12, 0, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Left
            };

            var execute = new System.Windows.Controls.Button
            {
                Content = "照査実行",
                Width = 120,
                Height = 32,
                Margin = new Thickness(0, 0, 8, 0)
            };
            execute.Click += OnExecuteButtonClick;
            _executeButton = execute;

            var reset = new System.Windows.Controls.Button
            {
                Content = "リセット",
                Width = 120,
                Height = 32
            };
            reset.Click += OnResetButtonClick;
            _resetButton = reset;

            buttons.Children.Add(execute);
            buttons.Children.Add(reset);

            outer.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            outer.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            System.Windows.Controls.Grid.SetColumn(keywordGrid, 0);
            System.Windows.Controls.Grid.SetColumn(buttons, 1);
            outer.Children.Add(keywordGrid);
            outer.Children.Add(buttons);

            return outer;
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
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(_leftColumnWidth) });
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

            // ① タグまたは曲げの詳細がホストされてない
            panel.Children.Add(CreateRebarExpander(
                "①タグまたは曲げ詳細なし",
                nameof(RebarTagCheckViewModel.Group1CountText),
                out _group1ListBox,
                nameof(RebarTagCheckViewModel.Group1Rebars),
                Brushes.Blue,
                System.Windows.Media.Color.FromRgb(227, 242, 253)));

            // ② 自由な端点のタグがホストされてない
            panel.Children.Add(CreateRebarExpander(
                "②自由端タグなし",
                nameof(RebarTagCheckViewModel.Group2CountText),
                out _group2ListBox,
                nameof(RebarTagCheckViewModel.Group2Rebars),
                Brushes.DarkGoldenrod,
                System.Windows.Media.Color.FromRgb(255, 249, 196)));

            // ③ 自由な端点のタグが鉄筋モデルを指している（赤=HOST不一致）
            panel.Children.Add(CreateRebarExpander(
                "③鉄筋モデルを指している",
                nameof(RebarTagCheckViewModel.Group3CountText),
                out _group3ListBox,
                nameof(RebarTagCheckViewModel.Group3Rebars),
                BuildGroup3ListItemStyle()));

            // ④ 自由な端点の先にある曲げの詳細を指している（赤=HOST不一致）
            panel.Children.Add(CreateRebarExpander(
                "④曲げ詳細を指している",
                nameof(RebarTagCheckViewModel.Group4CountText),
                out _group4ListBox,
                nameof(RebarTagCheckViewModel.Group4Rebars),
                BuildGroup4ListItemStyle()));

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
            return CreateRebarExpander(headerText, countBindingName, out listBox, itemsBindingName,
                BuildFixedColorListItemStyle(foreground, backgroundColor));
        }

        private UIElement CreateRebarExpander(
            string headerText,
            string countBindingName,
            out ListBox listBox,
            string itemsBindingName,
            Style itemStyle)
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
            listBox.ItemContainerStyle = itemStyle;
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

        // ③ 自由な端点のタグが鉄筋モデルを指している
        // IsLeaderPointingRebarMismatch=true の場合は赤、それ以外は緑
        private Style BuildGroup3ListItemStyle()
        {
            var style = new Style(typeof(ListBoxItem));
            style.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, Brushes.DarkGreen));
            style.Setters.Add(new Setter(System.Windows.Controls.Control.FontWeightProperty, FontWeights.SemiBold));
            style.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(232, 245, 233))));

            var trigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding(nameof(RebarListItem.IsLeaderPointingRebarMismatch)),
                Value = true
            };
            trigger.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, Brushes.Red));
            trigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 235, 238))));
            style.Triggers.Add(trigger);

            return style;
        }

        // ④ 自由な端点の先にある曲げの詳細を指している
        // IsLeaderPointingBendingDetailMismatch=true の場合は赤、それ以外は黒
        private Style BuildGroup4ListItemStyle()
        {
            var style = new Style(typeof(ListBoxItem));
            style.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, Brushes.Black));
            style.Setters.Add(new Setter(System.Windows.Controls.Control.FontWeightProperty, FontWeights.SemiBold));
            style.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(245, 245, 245))));

            var trigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding(nameof(RebarListItem.IsLeaderPointingBendingDetailMismatch)),
                Value = true
            };
            trigger.Setters.Add(new Setter(System.Windows.Controls.Control.ForegroundProperty, Brushes.Red));
            trigger.Setters.Add(new Setter(System.Windows.Controls.Control.BackgroundProperty, new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 235, 238))));
            style.Triggers.Add(trigger);

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

        private void OnExecuteButtonClick(object sender, RoutedEventArgs e)
        {
            TryExecuteCheck();
        }

        private void OnResetButtonClick(object sender, RoutedEventArgs e)
        {
            ResetToInitialState();
        }

        private enum ExecuteScope
        {
            Cancelled = 0,
            ActiveView = 1,
            KeywordSheets = 2,
        }

        private void TryExecuteCheck()
        {
            if (_checkExecuteExternalEventService == null)
            {
                TaskDialog.Show("RebarTag", "照査実行サービスが初期化されていません。");
                return;
            }

            var scope = SelectExecuteScope();
            if (scope == ExecuteScope.Cancelled)
            {
                return;
            }

            if (scope == ExecuteScope.KeywordSheets && IsKeywordEmpty())
            {
                RequireKeywordInput();
                return;
            }

            SetActionButtonsEnabled(false);

            var mode = scope == ExecuteScope.ActiveView
                ? RebarTagCheckExecutionMode.ActiveView
                : RebarTagCheckExecutionMode.KeywordSheets;

            _checkExecuteExternalEventService.Request(mode, _vm?.Keyword ?? string.Empty, OnCheckCompleted);
        }

        private ExecuteScope SelectExecuteScope()
        {
            // 実行時に対象範囲を選択する
            var td = new TaskDialog("RebarTag")
            {
                MainInstruction = "対象範囲を選択してください。",
                MainContent = "現在のビュー、またはKeywordを含むシートを対象に検証します。",
                CommonButtons = TaskDialogCommonButtons.Cancel
            };

            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "現在のビュー（Sheetの場合はそのSheetに配置されたビュー）");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Keyword（シート名にKeywordを含む）");

            var r = td.Show();
            if (r == TaskDialogResult.CommandLink1)
            {
                return ExecuteScope.ActiveView;
            }

            if (r == TaskDialogResult.CommandLink2)
            {
                return ExecuteScope.KeywordSheets;
            }

            return ExecuteScope.Cancelled;
        }

        private bool IsKeywordEmpty()
        {
            var k = _vm == null ? string.Empty : (_vm.Keyword ?? string.Empty);
            return string.IsNullOrWhiteSpace(k);
        }

        private void RequireKeywordInput()
        {
            TaskDialog.Show("RebarTag", "Keyword を入力してください。");
            FocusKeywordTextBox();
        }

        private void FocusKeywordTextBox()
        {
            try
            {
                _keywordTextBox?.Focus();
                _keywordTextBox?.SelectAll();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void OnCheckCompleted(RebarTagCheckExecutionResult result)
        {
            try
            {
                Dispatcher.Invoke(() => ApplyCheckResult(result));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                SetActionButtonsEnabled(true);
            }
        }

        private void ApplyCheckResult(RebarTagCheckExecutionResult result)
        {
            try
            {
                if (result == null || !result.IsSucceeded)
                {
                    var msg = result == null ? "結果が null です。" : (result.ErrorMessage ?? string.Empty);
                    TaskDialog.Show("RebarTag", string.IsNullOrWhiteSpace(msg) ? "照査に失敗しました。" : msg);
                    return;
                }

                ClearCurrentResultPanel();
                ClearAllSelections();

                UpdateTargetHeaderByResult(result);
                _vm.SetRebars(result.Rebars);
                _vm.SetRebarCount(result.RebarCount);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                TaskDialog.Show("RebarTag", ex.ToString());
            }
            finally
            {
                SetActionButtonsEnabled(true);
            }
        }

        private void UpdateTargetHeaderByResult(RebarTagCheckExecutionResult result)
        {
            if (result == null || _vm == null)
            {
                return;
            }

            if (result.TargetKind == RebarTagCheckExecutionTargetKind.SingleView)
            {
                _vm.SetTargetView(result.TargetView);
                return;
            }

            if (result.TargetKind == RebarTagCheckExecutionTargetKind.SheetWithPlacedViews)
            {
                _vm.SetTargetSheetAndViews(result.TargetSheet, result.PlacedViews);
                return;
            }

            if (result.TargetKind == RebarTagCheckExecutionTargetKind.MultipleSheets)
            {
                _vm.SetTargetSheetsAndViewCount(result.Sheets, result.ViewCount);
            }
        }

        private void ClearCurrentResultPanel()
        {
            // 直前選択の結果が残ると混乱するため、実行/リセットでクリアする
            try
            {
                _vm?.UpdateRows(new List<string>(), new List<string>());
                _vm?.UpdateNgReason(string.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void ClearAllSelections()
        {
            try
            {
                if (_group1ListBox != null) _group1ListBox.SelectedItem = null;
                if (_group2ListBox != null) _group2ListBox.SelectedItem = null;
                if (_group3ListBox != null) _group3ListBox.SelectedItem = null;
                if (_group4ListBox != null) _group4ListBox.SelectedItem = null;
                if (_resultGrid != null) _resultGrid.SelectedItem = null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void ResetToInitialState()
        {
            try
            {
                SetActionButtonsEnabled(true);
                ClearAllSelections();
                ClearCurrentResultPanel();

                if (_vm != null)
                {
                    _vm.Keyword = string.Empty;
                    _vm.SetRebars(new List<RebarListItem>());
                    _vm.SetRebarCount(0);

                    if (_initialView != null)
                    {
                        _vm.SetTargetView(_initialView);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void SetActionButtonsEnabled(bool enabled)
        {
            try
            {
                if (_executeButton != null) _executeButton.IsEnabled = enabled;
                if (_resetButton != null) _resetButton.IsEnabled = enabled;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
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
                if (!ReferenceEquals(sender, _group1ListBox) && _group1ListBox != null) _group1ListBox.SelectedItem = null;
                if (!ReferenceEquals(sender, _group2ListBox) && _group2ListBox != null) _group2ListBox.SelectedItem = null;
                if (!ReferenceEquals(sender, _group3ListBox) && _group3ListBox != null) _group3ListBox.SelectedItem = null;
                if (!ReferenceEquals(sender, _group4ListBox) && _group4ListBox != null) _group4ListBox.SelectedItem = null;
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

