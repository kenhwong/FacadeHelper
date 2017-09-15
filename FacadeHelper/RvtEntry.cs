using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Media.Imaging;

namespace FacadeHelper
{
    /// <remarks>
    /// This application's main class. The class must be Public.
    /// </remarks>
    public class RvtEntry : IExternalApplication
    {
        // Both OnStartup and OnShutdown must be implemented as public method
        public Result OnStartup(UIControlledApplication application)
        {
            application.ControlledApplication.DocumentOpened += new EventHandler<DocumentOpenedEventArgs>(Application_DocumentOpened);
            // Add a new ribbon panel
            RibbonPanel rpanel = application.CreateRibbonPanel("Facade Helper");

            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData bndata_process_elements = new PushButtonData("cmdProcessElements", "构件处理", thisAssemblyPath, "FacadeHelper.ICommand_Document_Process_Elements");
            PushButtonData bndata_zone = new PushButtonData("cmdZone", "分区处理", thisAssemblyPath, "FacadeHelper.ICommand_Document_Zone");
            PushButtonData bndata_zone4d = new PushButtonData("cmdZone4D", "4D分区处理", thisAssemblyPath, "FacadeHelper.ICommand_Document_Zone4D");
            bndata_process_elements.LargeImage = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources", "level32.png")));
            bndata_zone.LargeImage = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources", "se32.png")));
            bndata_zone.LargeImage = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources", "level32.png")));
            //rpanel.AddItem(bndata_process_elements);
            rpanel.AddItem(bndata_zone);
            rpanel.AddItem(bndata_zone4d);

            PushButtonData bndata_test = new PushButtonData("cmdTEST", "TEST", thisAssemblyPath, "FacadeHelper.ICommand_Document_TEST");
            //rpanel.AddItem(bndata_test);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // nothing to clean up in this simple case
            application.ControlledApplication.DocumentOpened -= new EventHandler<DocumentOpenedEventArgs>(Application_DocumentOpened);
            return Result.Succeeded;
        }

        public void Application_DocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            /** 數據庫模塊入口
            Document doc = e.Document;
            Global.DocContent = new DocumentContent();
            Global.DataFile = Path.Combine(Path.GetDirectoryName(doc.PathName), $"{Path.GetFileNameWithoutExtension(doc.PathName)}.data");
            Global.SQLDataFile = Path.Combine(Path.GetDirectoryName(doc.PathName), $"{Path.GetFileNameWithoutExtension(doc.PathName)}.db");
            Global.DocContent.CurrentDBContext = new SQLContext($"Data Source={Global.SQLDataFile}");
            Global.DocContent.CurrentDBContext.Database.Create();
            Global.DocContent.CurrentDBContext.SaveChanges();
            **/
        }


    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ICommand_Document_TEST : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            var app = uiapp.Application;
            Document doc = uidoc.Document;

            using (Transaction trans = new Transaction(doc, "CreateProjectParameters"))
            {
                trans.Start();
                //Gets the families
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                IList<Element> collection = collector.OfClass(typeof(Family)).ToElements();

                foreach (ElementId id in uidoc.Selection.GetElementIds())
                {
                    Element ele = doc.GetElement(id);
                    var ff = collection.Where(x => x.Name == (ele as FamilyInstance).Symbol.Family.Name).FirstOrDefault() as Family;
                    ff.Name += "CHANGED";
                }


                trans.Commit();
            }

            return Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ICommand_Document_Zone : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;

            try
            {
                Zone ucpe = new Zone(commandData);
                Window winaddin = new Window();
                ucpe.parentWin = winaddin;
                winaddin.Content = ucpe;
                //winaddin.WindowStyle = WindowStyle.None;
                winaddin.Padding = new Thickness(0);
                Global.winhelper = new System.Windows.Interop.WindowInteropHelper(winaddin);
                winaddin.ShowDialog();
            }
            catch (Exception e)
            {
                TaskDialog.Show("Error", e.ToString());
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ICommand_Document_Zone4D : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;

            try
            {
                Zone4D ucpe4d = new Zone4D(commandData);
                Window winaddin = new Window();
                ucpe4d.parentWin = winaddin;
                winaddin.Content = ucpe4d;
                //winaddin.WindowStyle = WindowStyle.None;
                winaddin.Padding = new Thickness(0);
                Global.winhelper = new System.Windows.Interop.WindowInteropHelper(winaddin);
                winaddin.ShowDialog();
            }
            catch (Exception e)
            {
                TaskDialog.Show("Error", e.ToString());
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }

}
