using System;
using System.Windows.Interop;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ShimizRevitAddin2026.ExternalEvents;
using ShimizRevitAddin2026.Services;
using ShimizRevitAddin2026.UI.Windows;

namespace ShimizRevitAddin2026.Commands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class OpenRebarTagCheckWindowCommand : IExternalCommand
    {
        private const string BendingDetailTagName = "曲げ加工詳細";
        private static readonly RebarTagCheckWindowStore _windowStore = new RebarTagCheckWindowStore();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (_windowStore.TryActivateExisting())
                {
                    return Result.Succeeded;
                }

                var uiapp = commandData.Application;
                var uidoc = uiapp.ActiveUIDocument;
                var doc = uidoc.Document;
                var activeView = doc.ActiveView;

                var dependentTagCollector = BuildDependentTagCollector();
                var consistencyService = BuildConsistencyService(dependentTagCollector);

                var highlighter = BuildHighlighter(dependentTagCollector);
                var externalEventService = new RebarTagHighlightExternalEventService(highlighter, consistencyService);

                // ウィンドウは「開くだけ」にして、照査はウィンドウ内ボタンで実行する
                var executionService = new RebarTagCheckExecutionService(dependentTagCollector, consistencyService);
                var checkExecuteService = new RebarTagCheckExecuteExternalEventService(executionService);

                var window = new RebarTagCheckWindow(uidoc, activeView, externalEventService, checkExecuteService);
                SetOwner(uiapp, window);
                _windowStore.SetCurrent(window);
                window.Show();

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RebarTag", ex.ToString());
                message = ex.Message;
                return Result.Failed;
            }
        }

        private RebarDependentTagCollector BuildDependentTagCollector()
        {
            var matcher = new RebarTagNameMatcher(BendingDetailTagName);
            return new RebarDependentTagCollector(matcher);
        }

        private RebarTagLeaderBendingDetailConsistencyService BuildConsistencyService(RebarDependentTagCollector dependentTagCollector)
        {
            return new RebarTagLeaderBendingDetailConsistencyService(dependentTagCollector);
        }

        private RebarTagHighlighter BuildHighlighter(RebarDependentTagCollector collector)
        {
            return new RebarTagHighlighter(collector);
        }

        private void SetOwner(UIApplication uiapp, System.Windows.Window window)
        {
            if (uiapp == null || window == null)
            {
                return;
            }

            try
            {
                var helper = new WindowInteropHelper(window);
                helper.Owner = uiapp.MainWindowHandle;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("RebarTag", ex.ToString());
            }
        }
    }
}

