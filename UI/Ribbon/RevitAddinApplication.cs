using System;
using Autodesk.Revit.UI;

namespace ShimizRevitAddin2026.UI.Ribbon
{
    public class RevitAddinApplication : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                var builder = new RevitRibbonBuilder();
                builder.Build(application);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ShimizRevitAddin2026", ex.ToString());
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}

