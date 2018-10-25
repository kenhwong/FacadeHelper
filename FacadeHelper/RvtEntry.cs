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
            Global.UpdateAppConfig("lang_2052", "/Resources/Langs/2052.xaml");
            Global.UpdateAppConfig("lang_1033", "/Resources/Langs/1033.xaml");
            
            //Global.SetCurrentLanguage(System.Threading.Thread.CurrentThread.CurrentCulture.LCID);

            application.ControlledApplication.DocumentOpened += new EventHandler<DocumentOpenedEventArgs>(Application_DocumentOpened);
            RibbonPanel rpanel = application.CreateRibbonPanel("Facade Helper");

            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData bndataAppconfig = new PushButtonData("cmdConfig", "全局设置", thisAssemblyPath, "FacadeHelper.Config_Command");
            PushButtonData bndataProcessElements = new PushButtonData("cmdProcessElements", "构件处理", thisAssemblyPath, "FacadeHelper.ICommand_Document_Process_Elements");
            PushButtonData bndataZone = new PushButtonData("cmdZone", "分区处理", thisAssemblyPath, "FacadeHelper.ICommand_Document_Zone");
            PushButtonData bndataFixzone = new PushButtonData("cmdFixZone", "分区修复", thisAssemblyPath, "FacadeHelper.ICommand_Document_FixZone");
            PushButtonData bndataFamilyParam = new PushButtonData("cmdFamilyParam", "族参初始", thisAssemblyPath, "FacadeHelper.ICommand_Document_AddFamilyParameters");
            PushButtonData bndataSelectfilter = new PushButtonData("cmdSelectFilter", "选择过滤", thisAssemblyPath, "FacadeHelper.ICommand_Document_SelectFilter");
            PushButtonData bndataCodeMapping = new PushButtonData("cmdCodeMapping", "编码映射", thisAssemblyPath, "FacadeHelper.ICommand_Document_CodeMapping");
            bndataAppconfig.LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.config32.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            bndataProcessElements.LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.level32.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            bndataZone.LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.se32.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            bndataFixzone.LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.fix32.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            bndataFamilyParam.LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.sv32.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            bndataSelectfilter.LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.filter32.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            bndataCodeMapping.LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.code32.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            rpanel.AddItem(bndataAppconfig);
            rpanel.AddItem(bndataZone);
            rpanel.AddItem(bndataFixzone);
            rpanel.AddSeparator();
            rpanel.AddItem(bndataFamilyParam);
            rpanel.AddItem(bndataSelectfilter);
            rpanel.AddItem(bndataCodeMapping);
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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;

            try
            {
                FacadeConfig ucpe = new FacadeConfig();
                Window winopt = new Window();
                ucpe.ParentWin = winopt;
                winopt.Content = ucpe;
                winopt.SizeToContent = SizeToContent.WidthAndHeight;
                //winaddin.WindowStyle = WindowStyle.None;
                winopt.Padding = new Thickness(0);
                Global.winhelper = new System.Windows.Interop.WindowInteropHelper(winopt);
                winopt.ShowDialog();
            }
            catch (Exception e)
            {
                TaskDialog.Show("Error", e.ToString());
                return Result.Failed;
            }

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

    [Transaction(TransactionMode.Manual)]
    public class ICommand_Document_SelectFilter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            commandData.Application.Application.SharedParametersFilename = Global.GetAppConfig("SharedParametersFile");

            try
            {
                SelectFilter ucpe = new SelectFilter(commandData);
                Window winaddin = new Window();
                ucpe.ParentWin = winaddin;
                winaddin.Content = ucpe;
                winaddin.SizeToContent = SizeToContent.WidthAndHeight;
                winaddin.WindowStartupLocation = WindowStartupLocation.CenterScreen;
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


    [Transaction(TransactionMode.Manual)]
    public class ICommand_Document_CodeMapping : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            commandData.Application.Application.SharedParametersFilename = Global.GetAppConfig("SharedParametersFile");

            try
            {
                FacadeCodeMapping ucpe = new FacadeCodeMapping(commandData);
                //ucpe.Resources.MergedDictionaries.Where(d => d.Source.OriginalString.Contains("Langs")).ToList().ForEach(d => ucpe.Resources.MergedDictionaries.Remove(d));
                //ucpe.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(ConfigurationManager.AppSettings.Get($"lang_{System.Threading.Thread.CurrentThread.CurrentCulture.LCID}"), UriKind.RelativeOrAbsolute) });
                Window winaddin = new Window();
                ucpe.ParentWin = winaddin;
                winaddin.Content = ucpe;
                winaddin.SizeToContent = SizeToContent.WidthAndHeight;
                winaddin.WindowStartupLocation = WindowStartupLocation.CenterScreen;
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

    [Transaction(TransactionMode.Manual)]
    public class ICommand_Document_AddFamilyParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            commandData.Application.Application.SharedParametersFilename = Global.GetAppConfig("SharedParametersFile");

            try
            {
                FamilyHelper ucpe = new FamilyHelper(commandData);
                Window winaddin = new Window();
                ucpe.ParentWin = winaddin;
                winaddin.Content = ucpe;
                winaddin.SizeToContent = SizeToContent.WidthAndHeight;
                winaddin.WindowStartupLocation = WindowStartupLocation.CenterScreen;
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

    public class ICommand_Document_FixZone : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;

            try
            {
                CodeFix ucpe = new CodeFix(commandData);
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
