using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Configuration;
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
            RibbonPanel rpanel = application.CreateRibbonPanel("Facade Helper");

            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData bndata_appconfig = new PushButtonData("cmdConfig", "全局设定", thisAssemblyPath, "FacadeHelper.Config_Command");
            PushButtonData bndata_process_elements = new PushButtonData("cmdProcessElements", "构件处理", thisAssemblyPath, "FacadeHelper.ICommand_Document_Process_Elements");
            PushButtonData bndata_zone = new PushButtonData("cmdZone", "分区处理", thisAssemblyPath, "FacadeHelper.ICommand_Document_Zone");
            bndata_appconfig.LargeImage = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources", "config32.png")));
            bndata_process_elements.LargeImage = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources", "level32.png")));
            bndata_zone.LargeImage = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources", "se32.png")));
            rpanel.AddItem(bndata_appconfig);
            rpanel.AddItem(bndata_zone);
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

    [Transaction(TransactionMode.Manual)]
    public class Config_Command : IExternalCommand
    {
        //http://blog.rodhowarth.com/2009/07/how-to-use-appconfig-file-in-dll-plugin.html
        //http://adndevblog.typepad.com/aec/2014/09/get-value-from-appconfig-file-and-sur-la-france.html

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            var app = uiapp.Application;
            Document doc = uidoc.Document;

            Global.UpdateAppConfig("TestProperty", "Test Value");
            Configuration config = ConfigurationManager.OpenExeConfiguration (Assembly.GetExecutingAssembly().Location);
            //string value = config.AppSettings.Settings["pdfOutput"].Value;
            config.AppSettings.Settings["TestProperty"].Value = "Test Value";
            config.Save(ConfigurationSaveMode.Full);
            //TaskDialog.Show("App config Value", value);

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
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

    [Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ICommand_Document_Zone : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;

            try
            {
                Zone ucpe = new Zone(commandData);
                Window winaddin = new Window();
                ucpe.ParentWin = winaddin;
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



}
