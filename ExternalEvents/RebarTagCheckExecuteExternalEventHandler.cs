using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ShimizRevitAddin2026.Services;
using ShimizRevitAddin2026.UI.Models;

namespace ShimizRevitAddin2026.ExternalEvents
{
    internal class RebarTagCheckExecuteExternalEventHandler : IExternalEventHandler
    {
        private readonly RebarTagCheckExecutionService _executionService;

        private RebarTagCheckExecutionMode _mode = RebarTagCheckExecutionMode.ActiveView;
        private string _keyword = string.Empty;
        private Action<RebarTagCheckExecutionResult> _onCompleted;

        public RebarTagCheckExecuteExternalEventHandler(RebarTagCheckExecutionService executionService)
        {
            _executionService = executionService;
        }

        public void SetRequest(RebarTagCheckExecutionMode mode, string keyword, Action<RebarTagCheckExecutionResult> onCompleted)
        {
            _mode = mode;
            _keyword = keyword ?? string.Empty;
            _onCompleted = onCompleted;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                var uidoc = app?.ActiveUIDocument;
                var doc = uidoc?.Document;
                var activeView = doc?.ActiveView;

                if (_executionService == null || doc == null || activeView == null)
                {
                    NotifyCompleted(BuildFailed("実行サービス、または Document/View が取得できません。"));
                    return;
                }

                if (_mode == RebarTagCheckExecutionMode.ActiveView)
                {
                    NotifyCompleted(_executionService.ExecuteForActiveContext(doc, activeView));
                    return;
                }

                NotifyCompleted(_executionService.ExecuteForKeywordSheets(doc, _keyword));
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // 操作キャンセル
                NotifyCompleted(BuildFailed("キャンセルされました。"));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                try
                {
                    TaskDialog.Show("RebarTag", ex.ToString());
                }
                catch (Exception dialogEx)
                {
                    Debug.WriteLine(dialogEx);
                }

                NotifyCompleted(BuildFailed(ex.ToString()));
            }
        }

        public string GetName()
        {
            return "RebarTagCheckExecuteExternalEventHandler";
        }

        private RebarTagCheckExecutionResult BuildFailed(string message)
        {
            return new RebarTagCheckExecutionResult(
                false,
                message ?? string.Empty,
                RebarTagCheckExecutionTargetKind.None,
                null,
                null,
                new List<View>(),
                new List<ViewSheet>(),
                0,
                new List<RebarListItem>(),
                0);
        }

        private void NotifyCompleted(RebarTagCheckExecutionResult result)
        {
            try
            {
                _onCompleted?.Invoke(result);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }
    }
}

