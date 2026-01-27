using System;
using System.Diagnostics;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ShimizRevitAddin2026.Services;
using ShimizRevitAddin2026.UI.Models;

namespace ShimizRevitAddin2026.ExternalEvents
{
    internal class RebarTagHighlightExternalEventService
    {
        private readonly RebarTagHighlightExternalEventHandler _handler;
        private readonly ExternalEvent _externalEvent;

        public RebarTagHighlightExternalEventService(RebarTagHighlighter highlighter)
        {
            _handler = new RebarTagHighlightExternalEventHandler(highlighter);
            _externalEvent = ExternalEvent.Create(_handler);
        }

        public void Request(ElementId rebarId, ElementId viewId, Action<RebarTagCheckResult> onCompleted)
        {
            try
            {
                _handler.SetRequest(rebarId, viewId, onCompleted);
                _externalEvent.Raise();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                try
                {
                    onCompleted?.Invoke(new RebarTagCheckResult(null, null));
                }
                catch (Exception callbackEx)
                {
                    Debug.WriteLine(callbackEx);
                }
            }
        }
    }
}

