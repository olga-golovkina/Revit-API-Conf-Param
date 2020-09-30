using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Reflection;

namespace ConfParam
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class App : IExternalApplication
    {
        public static string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
        public Result OnStartup(UIControlledApplication app)
        {
            string tabName = "Configuration Parameters";
            try
            {
                app.CreateRibbonTab(tabName);
            }
            catch (Exception) { return Result.Failed; }

            PushButtonData pushButtonData01 = null;
            PushButton pushButton01 = null;
            RibbonPanel ribbonPanel = null;
            try
            {
                ribbonPanel = app.CreateRibbonPanel(tabName, "Configuration Parameters");
            }
            catch (Exception) { return Result.Failed; }

            thisAssemblyPath = Path.Combine(Assembly.GetExecutingAssembly().Location);

            pushButtonData01 = new PushButtonData("ConfParam", "\n\n\nStart Code", thisAssemblyPath, "ConfParam.ConfParamClass");
            pushButton01 = ribbonPanel.AddItem(pushButtonData01) as PushButton;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }
}
