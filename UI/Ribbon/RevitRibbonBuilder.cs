using System;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.UI;
using ShimizRevitAddin2026.Commands;

namespace ShimizRevitAddin2026.UI.Ribbon
{
    internal class RevitRibbonBuilder
    {
        private readonly RibbonIconFactory _iconFactory;

        public RevitRibbonBuilder()
        {
            _iconFactory = new RibbonIconFactory();
        }

        public void Build(UIControlledApplication application)
        {
            if (application == null)
            {
                return;
            }

            var tabName = "Shimiz";
            CreateTabIfNeeded(application, tabName);

            var panel = CreateOrGetPanel(application, tabName, "鉄筋");
            AddRebarCheckButton(panel);
        }

        private void CreateTabIfNeeded(UIControlledApplication application, string tabName)
        {
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                // タブが既に存在する場合は何もしない
            }
        }

        private RibbonPanel CreateOrGetPanel(UIControlledApplication application, string tabName, string panelName)
        {
            try
            {
                var existing = application
                    .GetRibbonPanels(tabName)
                    .FirstOrDefault(p => string.Equals(p.Name, panelName, StringComparison.Ordinal));

                return existing ?? application.CreateRibbonPanel(tabName, panelName);
            }
            catch (Exception ex)
            {
                // パネル作成失敗時は例外内容を表示して呼び出し側に伝播
                TaskDialog.Show("ShimizRevitAddin2026", ex.ToString());
                throw;
            }
        }

        private void AddRebarCheckButton(RibbonPanel panel)
        {
            if (panel == null)
            {
                return;
            }

            var data = BuildRebarCheckButtonData();
            var created = panel.AddItem(data) as PushButton;
            if (created == null)
            {
                return;
            }

            ConfigureRebarCheckButton(created);
        }

        private PushButtonData BuildRebarCheckButtonData()
        {
            var assemblyPath = Assembly.GetExecutingAssembly().Location;
            var commandType = typeof(OpenRebarTagCheckWindowCommand);
            var commandFullName = commandType.FullName ?? commandType.Name;

            return new PushButtonData(
                "Shimiz.RebarTagCheck",
                "鉄筋照査",
                assemblyPath,
                commandFullName);
        }

        private void ConfigureRebarCheckButton(PushButton button)
        {
            button.ToolTip = "鉄筋タグ（曲げ加工詳細）を照査して一覧表示します。";
            button.LongDescription = "アクティブビュー内の鉄筋を集計し、タグの不整合（引出線・曲げ加工詳細）を確認します。";

            button.Image = _iconFactory.CreateIcon16();
            button.LargeImage = _iconFactory.CreateIcon32();
        }
    }
}

