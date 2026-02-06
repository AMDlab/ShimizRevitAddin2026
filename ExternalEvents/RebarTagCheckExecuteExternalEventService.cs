using System;
using System.Diagnostics;
using Autodesk.Revit.UI;
using ShimizRevitAddin2026.Services;
using ShimizRevitAddin2026.UI.Models;

namespace ShimizRevitAddin2026.ExternalEvents
{
    internal class RebarTagCheckExecuteExternalEventService
    {
        private readonly RebarTagCheckExecuteExternalEventHandler _handler;
        private readonly ExternalEvent _externalEvent;

        public RebarTagCheckExecuteExternalEventService(RebarTagCheckExecutionService executionService)
        {
            _handler = new RebarTagCheckExecuteExternalEventHandler(executionService);
            _externalEvent = ExternalEvent.Create(_handler);
        }

        public void Request(RebarTagCheckExecutionMode mode, string keyword, Action<RebarTagCheckExecutionResult> onCompleted)
        {
            try
            {
                _handler.SetRequest(mode, keyword, onCompleted);
                _externalEvent.Raise();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                try
                {
                    onCompleted?.Invoke(new RebarTagCheckExecutionResult(false, ex.ToString(), RebarTagCheckExecutionTargetKind.None, null, null, null, null, 0, null, 0));
                }
                catch (Exception callbackEx)
                {
                    Debug.WriteLine(callbackEx);
                }
            }
        }
    }
}

