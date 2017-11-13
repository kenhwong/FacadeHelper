using AqlaSerializer;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using RvtApplication = Autodesk.Revit.ApplicationServices.Application;

// The Building Coder Samples https://github.com/jeremytammik/the_building_coder_samples
//http://www.revitapidocs.com

//https://forums.autodesk.com/t5/revit-api-forum/project-parameter-id-amp-guid/td-p/5834037
//https://forums.autodesk.com/autodesk/attachments/autodesk/160/5205/1/GetProjectParameterGUIDs.txt
//https://forums.autodesk.com/t5/revit-api-forum/reporting-on-project-parameter-definitions-need-guids/td-p/4684297

namespace FacadeHelper
{


    public class ProjectParameterData
    {
        public Definition Definition = null;
        public ElementBinding Binding = null;
        public bool IsSharedStatusKnown = false;  // Will probably always be true when the data is gathered
        public bool IsShared = false;
        public string GUID = null;
    }

    /// <summary>
    /// Provides static functions to convert unit
    /// </summary>
    static class Unit
    {
        #region Methods
        /// <summary>
        /// Convert the value get from RevitAPI to the value indicated by DisplayUnitType
        /// </summary>
        /// <param name="to">DisplayUnitType indicates unit of target value</param>
        /// <param name="value">value get from RevitAPI</param>
        /// <returns>Target value</returns>
        public static double CovertFromAPI(DisplayUnitType to, double value)
        {
            return value *= ImperialDutRatio(to);
        }

        /// <summary>
        /// Convert a value indicated by DisplayUnitType to the value used by RevitAPI
        /// </summary>
        /// <param name="value">Value to be converted</param>
        /// <param name="from">DisplayUnitType indicates the unit of the value to be converted</param>
        /// <returns>Target value</returns>
        public static double CovertToAPI(double value, DisplayUnitType from)
        {
            return value /= ImperialDutRatio(from);
        }

        /// <summary>
        /// Get ratio between value in RevitAPI and value to display indicated by DisplayUnitType
        /// </summary>
        /// <param name="dut">DisplayUnitType indicates display unit type</param>
        /// <returns>Ratio </returns>
        private static double ImperialDutRatio(DisplayUnitType dut)
        {
            switch (dut)
            {
                case DisplayUnitType.DUT_DECIMAL_FEET: return 1;
                case DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES: return 1;
                case DisplayUnitType.DUT_DECIMAL_INCHES: return 12;
                case DisplayUnitType.DUT_FRACTIONAL_INCHES: return 12;
                case DisplayUnitType.DUT_METERS: return 0.3048;
                case DisplayUnitType.DUT_CENTIMETERS: return 30.48;
                case DisplayUnitType.DUT_MILLIMETERS: return 304.8;
                case DisplayUnitType.DUT_METERS_CENTIMETERS: return 0.3048;
                default: return 1;
            }
        }
        #endregion
    }
    public static class ParameterHelper
    {

        #region RawProjectParameterInfo

        public class RawProjectParameterInfo
        {
            public static string FileName { get; set; }

            public string Name { get; set; }
            public BuiltInParameterGroup Group { get; set; }
            public ParameterType Type { get; set; }
            public bool ReadOnly { get; set; }
            public bool BoundToInstance { get; set; }
            public string[] BoundCategories { get; set; }

            public bool FromShared { get; set; }
            public string GUID { get; set; }
            public string Owner { get; set; }
            public bool Visible { get; set; }
        }

        public static List<T> RawConvertSetToList<T>(IEnumerable set)
        {
            List<T> list = (from T p in set select p).ToList<T>();
            return list;
        }

        public static List<RawProjectParameterInfo> RawGetProjectParametersInfo(Document doc)
        {
            RawProjectParameterInfo.FileName = doc.Title;
            List<RawProjectParameterInfo> paramList = new List<RawProjectParameterInfo>();

            BindingMap map = doc.ParameterBindings;
            DefinitionBindingMapIterator it = map.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                ElementBinding eleBinding = it.Current as ElementBinding;
                InstanceBinding insBinding = eleBinding as InstanceBinding;
                Definition def = it.Key;
                if (def != null)
                {
                    ExternalDefinition extDef = def as ExternalDefinition;
                    bool shared = extDef != null;
                    RawProjectParameterInfo param = new RawProjectParameterInfo
                    {
                        Name = def.Name,
                        Group = def.ParameterGroup,
                        Type = def.ParameterType,
                        ReadOnly = true, // def.IsReadOnly, def.IsReadOnly NOT working in either 2015 or 2014 but working in 2013
                        BoundToInstance = insBinding != null,
                        BoundCategories = RawConvertSetToList<Category>(eleBinding.Categories).Select(c => c.Name).ToArray(),

                        FromShared = shared,
                        GUID = shared ? extDef.GUID.ToString() : string.Empty,
                        Owner = shared ? extDef.OwnerGroup.Name : string.Empty,
                        Visible = shared ? extDef.Visible : true,
                    };

                    paramList.Add(param);
                }
            }

            return paramList;
        }

        /// <summary>
        /// convert the informative object List into a single string
        /// </summary>
        /// <param name="infoList"></param>
        /// <param name="title"></param>
        /// <returns></returns>
        public static string RawParametersInfoToCSVString(List<RawProjectParameterInfo> infoList, ref string title)
        {
            StringBuilder sb = new StringBuilder();
            PropertyInfo[] propInfoArray = typeof(RawProjectParameterInfo).GetProperties();
            foreach (PropertyInfo pi in propInfoArray)
            {
                title += pi.Name + ",";
            }
            title = title.Remove(title.Length - 1);

            foreach (RawProjectParameterInfo info in infoList)
            {
                foreach (PropertyInfo pi in propInfoArray)
                {
                    object obj = info.GetType().InvokeMember(pi.Name, BindingFlags.GetProperty, null, info, null);
                    IList list = obj as IList;
                    if (list != null)
                    {
                        string str = string.Empty;
                        foreach (object e in list)
                        {
                            str += e.ToString() + ";";
                        }
                        str = str.Remove(str.Length - 1);

                        sb.Append(str + ",");
                    }
                    else
                    {
                        sb.Append((obj == null ? string.Empty : obj.ToString()) + ",");
                    }
                }
                sb.Remove(sb.Length - 1, 1).Append(Environment.NewLine);
            }

            return sb.ToString();
        }

        //Then we can write all the project parameter information of a Revit Document to a CSV file.

        /***
        …
        List<RawProjectParameterInfo> paramsInfo = RawGetProjectParametersInfo(CachedDoc);
        using (StreamWriter sw = new StreamWriter(@"c:\temp\ProjectParametersInfo.csv"))
        {
            string title = string.Empty;
            string rows = RawParametersInfoToCSVString(paramsInfo, ref title);
            sw.WriteLine(title);
            sw.Write(rows);
        }
        …
         * **/


        #endregion


        #region Create Project Parameter

        public static void RawCreateProjectParameterFromExistingSharedParameter(RvtApplication app, string name, CategorySet cats, BuiltInParameterGroup group, bool inst)
        {
            DefinitionFile defFile = app.OpenSharedParameterFile();
            if (defFile == null) throw new Exception("No SharedParameter File!");

            var v = (from DefinitionGroup dg in defFile.Groups
                     from ExternalDefinition d in dg.Definitions
                     where d.Name == name
                     select d);
            if (v == null || v.Count() < 1) throw new Exception("Invalid Name Input!");

            ExternalDefinition def = v.First();

            Autodesk.Revit.DB.Binding binding = app.Create.NewTypeBinding(cats);
            if (inst) binding = app.Create.NewInstanceBinding(cats);

            BindingMap map = (new UIApplication(app)).ActiveUIDocument.Document.ParameterBindings;
            map.Insert(def, binding, group);
        }

        public static void RawCreateProjectParameterFromNewSharedParameter(RvtApplication app, string defGroup, string name, ParameterType type, bool visible, CategorySet cats, BuiltInParameterGroup paramGroup, bool inst)
        {
            DefinitionFile defFile = app.OpenSharedParameterFile();
            if (defFile == null) throw new Exception("No SharedParameter File!");

            ExternalDefinition def = app.OpenSharedParameterFile().Groups.Create(defGroup).Definitions.Create(name, type, visible) as ExternalDefinition;

            Autodesk.Revit.DB.Binding binding = app.Create.NewTypeBinding(cats);
            if (inst) binding = app.Create.NewInstanceBinding(cats);

            BindingMap map = (new UIApplication(app)).ActiveUIDocument.Document.ParameterBindings;
            map.Insert(def, binding, paramGroup);
        }

        public static void RawCreateProjectParameter(RvtApplication app, string name, ParameterType type, bool visible, CategorySet cats, BuiltInParameterGroup group, bool inst)
        {
            //InternalDefinition def = new InternalDefinition();
            //Definition def = new Definition();

            string oriFile = app.SharedParametersFilename;
            string tempFile = Path.GetTempFileName() + ".txt";
            using (File.Create(tempFile)) { }
            app.SharedParametersFilename = tempFile;

            ExternalDefinition def = app.OpenSharedParameterFile().Groups.Create("TemporaryDefintionGroup").Definitions.Create(name, type, visible) as ExternalDefinition;

            app.SharedParametersFilename = oriFile;
            File.Delete(tempFile);

            Autodesk.Revit.DB.Binding binding = app.Create.NewTypeBinding(cats);
            if (inst) binding = app.Create.NewInstanceBinding(cats);

            BindingMap map = (new UIApplication(app)).ActiveUIDocument.Document.ParameterBindings;
            map.Insert(def, binding, group);
        }

        #endregion

        #region project parameter guid
        /// <summary>
        /// This class contains information discovered about a (shared or non-shared) project parameter 
        /// </summary>


        // ================= HELPER METHODS ======================================================================================



        /// <summary>
        /// Returns a list of the objects containing references to the project parameter definitions
        /// </summary>
        /// <param name="projectDocument">The project document being quereied</param>
        /// <returns></returns>
        public static List<ProjectParameterData> GetProjectParameterData(Document projectDocument)
        {
            // Following good SOA practices, first validate incoming parameters
            if (projectDocument == null)
            {
                throw new ArgumentNullException("projectDocument");
            }

            if (projectDocument.IsFamilyDocument)
            {
                throw new Exception("projectDocument can not be a family document.");
            }

            List<ProjectParameterData> result = new List<ProjectParameterData>();
            BindingMap map = projectDocument.ParameterBindings;
            DefinitionBindingMapIterator it = map.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                ProjectParameterData newProjectParameterData = new ProjectParameterData();
                newProjectParameterData.Definition = it.Key;
                newProjectParameterData.Binding = it.Current as ElementBinding;
                result.Add(newProjectParameterData);
            }

            return result;
        }



        /// <summary>
        /// This method takes a category and information about a project parameter and 
        /// adds a binding to the category for the parameter.  It will throw an exception if the parameter
        /// is already bound to the desired category.  It returns whether or not the API reports that it
        /// successfully bound the parameter to the desired category.
        /// </summary>
        /// <param name="projectDocument">The project document in which the project parameter has been defined</param>
        /// <param name="projectParameterData">Information about the project parameter</param>
        /// <param name="category">The additional category to which to bind the project parameter</param>
        /// <returns></returns>
        public static bool AddProjectParameterBinding(Document projectDocument,
                                                       ProjectParameterData projectParameterData,
                                                       Category category)
        {
            // Following good SOA practices, first validate incoming parameters
            if (projectDocument == null)
            {
                throw new ArgumentNullException("projectDocument");
            }

            if (projectDocument.IsFamilyDocument)
            {
                throw new Exception("projectDocument can not be a family document.");
            }

            if (projectParameterData == null)
            {
                throw new ArgumentNullException("projectParameterData");
            }

            if (category == null)
            {
                throw new ArgumentNullException("category");
            }

            bool result = false;
            CategorySet cats = projectParameterData.Binding.Categories;

            if (cats.Contains(category))
            {
                // It's already bound to the desired category.  Nothing to do.
                string errorMessage = string.Format("The project parameter '{0}' is already bound to the '{1}' category.",
                                                    projectParameterData.Definition.Name,
                                                    category.Name);

                throw new Exception(errorMessage);
            }

            cats.Insert(category);

            // See if the parameter is an instance or type parameter.
            InstanceBinding instanceBinding = projectParameterData.Binding as InstanceBinding;

            if (instanceBinding != null)
            {
                // Is an Instance parameter
                InstanceBinding newInstanceBinding = projectDocument.Application.Create.NewInstanceBinding(cats);
                if (projectDocument.ParameterBindings.ReInsert(projectParameterData.Definition, newInstanceBinding))
                {
                    result = true;
                }
            }
            else
            {
                // Is a type parameter
                TypeBinding typeBinding = projectDocument.Application.Create.NewTypeBinding(cats);
                if (projectDocument.ParameterBindings.ReInsert(projectParameterData.Definition, typeBinding))
                {
                    result = true;
                }
            }

            return result;
        }




        /// <summary>
        /// This method populates the appropriate values on a ProjectParameterData object with information from
        /// the given Parameter object.
        /// </summary>
        /// <param name="parameter">The Parameter object with source information</param>
        /// <param name="projectParameterDataToFill">The ProjectParameterData object to fill</param>
        public static void PopulateProjectParameterData(Parameter parameter,
                                                         ProjectParameterData projectParameterDataToFill)
        {
            // Following good SOA practices, validat incoming parameters first.
            if (parameter == null)
            {
                throw new ArgumentNullException("parameter");
            }

            if (projectParameterDataToFill == null)
            {
                throw new ArgumentNullException("projectParameterDataToFill");
            }

            projectParameterDataToFill.IsSharedStatusKnown = true;
            projectParameterDataToFill.IsShared = parameter.IsShared;
            if (parameter.IsShared)
            {
                if (parameter.GUID != null)
                {
                    projectParameterDataToFill.GUID = parameter.GUID.ToString();
                }
            }

        }  // end of PopulateProjectParameterData

        #endregion

        #region 初始化项目参数
        public static void InitProjectParameters(ref Document doc)
        {
            #region 设置项目参数

            using (Transaction trans = new Transaction(doc, "CreateProjectParameters"))
            {
                trans.Start();
                #region 设置项目参数：立面朝向
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "立面朝向"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "立面朝向", ParameterType.Text, true, _catset, BuiltInParameterGroup.PG_DATA, true);
                }
                #endregion
                #region 设置项目参数：立面系统
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "立面系统"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "立面系统", ParameterType.Text, true, _catset, BuiltInParameterGroup.PG_DATA, true);
                }
                #endregion
                #region 设置项目参数：立面楼层
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "立面楼层"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "立面楼层", ParameterType.Integer, true, _catset, BuiltInParameterGroup.PG_DATA, true);
                }
                #endregion
                #region 设置项目参数：构件分项
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "构件分项"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "构件分项", ParameterType.Integer, true, _catset, BuiltInParameterGroup.PG_DATA, true);
                }
                #endregion
                #region 设置项目参数：构件子项
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "构件子项"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "构件子项", ParameterType.Text, true, _catset, BuiltInParameterGroup.PG_DATA, true);
                }
                #endregion
                #region 设置项目参数：加工图号
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "加工图号"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "加工图号", ParameterType.Text, true, _catset, BuiltInParameterGroup.PG_DATA, true);
                }
                #endregion

                #region 设置项目参数：分区序号
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "分区序号"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "分区序号", ParameterType.Integer, true, _catset, BuiltInParameterGroup.PG_DATA, true);
                }
                #endregion

                #region 设置项目参数：分区区号
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "分区区号"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "分区区号", ParameterType.Text, true, _catset, BuiltInParameterGroup.PG_DATA, true);
                }
                #endregion

                #region 设置项目参数：分区编码
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "分区编码"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "分区编码", ParameterType.Text, true, _catset, BuiltInParameterGroup.PG_DATA, true);
                }
                #endregion

                #region 设置项目参数：進場時間
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "进场时间"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "进场时间", ParameterType.Text, true, _catset, BuiltInParameterGroup.PG_PHASING, true);
                }
                #endregion

                #region 设置项目参数：安裝開始
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "安装开始"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "安装开始", ParameterType.Text, true, _catset, BuiltInParameterGroup.PG_PHASING, true);
                }
                #endregion

                #region 设置项目参数：安裝結束
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "安装结束"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallMullions));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "安装结束", ParameterType.Text, true, _catset, BuiltInParameterGroup.PG_PHASING, true);
                }
                #endregion
                Global.DocContent.ParameterInfoList = ParameterHelper.RawGetProjectParametersInfo(doc);
                trans.Commit();

            }
            #endregion
        }

        #endregion

        public static void IsolateCategoriesAndZoom(ElementId eid, UIApplication uiapp)
        {

        }

    }

    public static class ZoneHelper
    {
        /// <summary>
        /// 加载ACADE导出的分区进度数据 - 4D简单模型
        /// </summary>
        /// <param name="ZoneScheduleDataFile">进度数据文件(*.txt)</param>
        #region 加载ACADE导出的分区进度数据 - 4D简单模型
        public static void FnLoadSimpleZoneScheduleData(string ZoneScheduleDataFile)
        {
            Global.DocContent.ZoneScheduleSimpleList.Clear();
            using (StreamReader reader = new StreamReader(ZoneScheduleDataFile))
            {
                string dataline;
                reader.ReadLine();
                while ((dataline = reader.ReadLine()) != null)
                {
                    string[] rowdata = dataline.Split('\t');
                    if (int.TryParse(rowdata[2].Substring(2, 2), out int lv2))
                        Global.DocContent.ZoneScheduleSimpleList.Add(new ZoneLayerInfo
                        {
                            ZoneCode = Regex.Replace(rowdata[2], @"Z-\d{2}-", @"Z-00-", RegexOptions.IgnoreCase),
                            ZoneLayer = lv2,
                            ZoneStart = DateTime.Parse($"{rowdata[3].Substring(0, 2)}/{rowdata[3].Substring(2, 2)}/{rowdata[3].Substring(4, 2)}"),
                            ZoneFinish = DateTime.Parse($"{rowdata[4].Substring(0, 2)}/{rowdata[4].Substring(2, 2)}/{rowdata[4].Substring(4, 2)}"),
                        });
                    if (int.TryParse(rowdata[5].Substring(2, 2), out int lv5))
                        Global.DocContent.ZoneScheduleSimpleList.Add(new ZoneLayerInfo
                        {
                            ZoneCode = Regex.Replace(rowdata[5], @"Z-\d{2}-", @"Z-00-", RegexOptions.IgnoreCase),
                            ZoneLayer = lv5,
                            ZoneStart = DateTime.Parse($"{rowdata[6].Substring(0, 2)}/{rowdata[6].Substring(2, 2)}/{rowdata[6].Substring(4, 2)}"),
                            ZoneFinish = DateTime.Parse($"{rowdata[7].Substring(0, 2)}/{rowdata[7].Substring(2, 2)}/{rowdata[7].Substring(4, 2)}"),
                        });
                }
            }
        }

        #endregion

        /// <summary>
        /// 加载ACADE导出的分区进度数据 - 4D设计模型
        /// </summary>
        /// <param name="ZoneScheduleDataFile">进度数据文件(*.txt)</param>
        #region 加载ACADE导出的分区进度数据 - 4D设计模型
        public static void FnLoadZoneScheduleData(string ZoneScheduleDataFile)
        {
            Global.DocContent.ZoneLayerList.Clear();

            using (StreamReader reader = new StreamReader(ZoneScheduleDataFile))
            {
                string dataline;
                reader.ReadLine();
                ZoneLayerInfo zli = new ZoneLayerInfo();
                string linehandleid;
                string linezonecode = string.Empty;
                while ((dataline = reader.ReadLine()) != null)
                {
                    string[] rowdata = dataline.Split('\t');
                    linehandleid = rowdata[0].Replace(@"'", ""); //AutoCAD Attribute Block Handle ID.

                    if ((rowdata.Length > 5 && rowdata[2] == rowdata[5]) || (rowdata.Length > 8 && (rowdata[2] == rowdata[8] || rowdata[5] == rowdata[8])))
                    {
                        MessageBox.Show($"当前行[{rowdata[0]}]的分区编号有重复，不能继续读取分析数据。", "错误 - 分区进度数据");
                        return;
                    }

                    if (rowdata.Length > 2)
                        if (int.TryParse(rowdata[2].Substring(2, 2), out int lv2))
                        {
                            if (lv2 > 3 || lv2 < 1)
                            {
                                MessageBox.Show($"当前行[{rowdata[0]}]的分区[{rowdata[2]}]的工序层{lv2}小于预设的1层下限或大于3层上限，不能继续读取分析数据。", "错误 - 分区进度数据");
                                return;
                            }
                            linezonecode = Regex.Replace(rowdata[2], @"Z-\d{2}-", @"Z-00-", RegexOptions.IgnoreCase);
                            Global.DocContent.ZoneLayerList.Add(new ZoneLayerInfo()
                            {
                                HandleId = linehandleid,
                                ZoneLayer = lv2 - 1,
                                ZoneCode = linezonecode,
                                ZoneStart = DateTime.Parse($"{rowdata[3].Substring(0, 2)}/{rowdata[3].Substring(2, 2)}/{rowdata[3].Substring(4, 2)}"),
                                ZoneFinish = DateTime.Parse($"{rowdata[4].Substring(0, 2)}/{rowdata[4].Substring(2, 2)}/{rowdata[4].Substring(4, 2)}")
                            });
                        }
                    if (rowdata.Length > 5)
                        if (int.TryParse(rowdata[5].Substring(2, 2), out int lv5))
                        {
                            if (lv5 > 3 || lv5 < 1)
                            {
                                MessageBox.Show($"当前行[{rowdata[0]}]的分区[{rowdata[5]}]的工序层{lv5}小于预设的1层下限或大于3层上限，不能继续读取分析数据。", "错误 - 分区进度数据");
                                return;
                            }
                            if (linezonecode != Regex.Replace(rowdata[5], @"Z-\d{2}-", @"Z-00-", RegexOptions.IgnoreCase))
                            {
                                MessageBox.Show($"当前行[{rowdata[0]}]的分区[{linezonecode}]编号不统一，不能继续读取分析数据。", "错误 - 分区进度数据");
                                return;
                            }
                            Global.DocContent.ZoneLayerList.Add(new ZoneLayerInfo()
                            {
                                HandleId = linehandleid,
                                ZoneLayer = lv5 - 1,
                                ZoneCode = linezonecode,
                                ZoneStart = DateTime.Parse($"{rowdata[6].Substring(0, 2)}/{rowdata[6].Substring(2, 2)}/{rowdata[6].Substring(4, 2)}"),
                                ZoneFinish = DateTime.Parse($"{rowdata[7].Substring(0, 2)}/{rowdata[7].Substring(2, 2)}/{rowdata[7].Substring(4, 2)}")
                            });
                        }
                    if (rowdata.Length > 8)
                        if (int.TryParse(rowdata[8].Substring(2, 2), out int lv8))
                        {
                            if (lv8 > 3 || lv8 < 1)
                            {
                                MessageBox.Show($"当前行[{rowdata[0]}]的分区[{rowdata[8]}]的工序层{lv8}小于预设的1层下限或大于3层上限，不能继续读取分析数据。", "错误 - 分区进度数据");
                                return;
                            }
                            if (linezonecode != Regex.Replace(rowdata[8], @"Z-\d{2}-", @"Z-00-", RegexOptions.IgnoreCase))
                            {
                                MessageBox.Show($"当前行[{rowdata[0]}]的分区编号[{linezonecode}]不统一，不能继续读取分析数据。", "错误 - 分区进度数据");
                                return;
                            }
                            Global.DocContent.ZoneLayerList.Add(new ZoneLayerInfo()
                            {
                                HandleId = linehandleid,
                                ZoneLayer = lv8 - 1,
                                ZoneCode = linezonecode,
                                ZoneStart = DateTime.Parse($"{rowdata[9].Substring(0, 2)}/{rowdata[9].Substring(2, 2)}/{rowdata[9].Substring(4, 2)}"),
                                ZoneFinish = DateTime.Parse($"{rowdata[10].Substring(0, 2)}/{rowdata[10].Substring(2, 2)}/{rowdata[10].Substring(4, 2)}")
                            });
                        }
                }
            }
        }

        #endregion

        public static void FnSearch(UIDocument uidoc,
            string querystring,
            ref List<ZoneInfoBase> relistzone, ref List<CurtainPanelInfo> relistpanel, ref List<ScheduleElementInfo> relistelement,
            bool hasrangezone, bool hasrangepanel, bool hasrangeelement,
            ref ListBox listinfo)
        {
            var doc = uidoc.Document;
            listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - F: KEY/{querystring}...");

            if (hasrangezone && DateTime.TryParse(querystring, out DateTime querydatetime))
                relistzone.AddRange(Global.DocContent.ZoneList.Where(z => z.ZoneLayerStart.Equals(querydatetime) || z.ZoneLayerFinish.Equals(querydatetime)));

            if (hasrangezone) relistzone.AddRange(Global.DocContent.ZoneList.Where(z => Regex.IsMatch(z.ZoneCode, querystring, RegexOptions.IgnoreCase)));
            if (hasrangepanel) relistpanel.AddRange(Global.DocContent.CurtainPanelList.Where(p => Regex.IsMatch($"{p.INF_ElementId} # {p.INF_Code}", querystring, RegexOptions.IgnoreCase)));
            if (hasrangeelement) relistelement.AddRange(Global.DocContent.ScheduleElementList.Where(e => Regex.IsMatch($"{e.INF_ElementId} # {e.INF_Code}", querystring, RegexOptions.IgnoreCase)));
        }

        public static void FnSearch(UIDocument uidoc,
            string querystring,
            ref List<ZoneInfoBase> relistzone, ref List<CurtainPanelInfo> relistpanel, ref List<MullionInfo> relistmullion,
            bool hasrangezone, bool hasrangepanel, bool hasrangeelement,
            ref ListBox listinfo)
        {

        }

        /// <summary>
        /// 4D简单分区，嵌板和竖梃分析
        /// </summary>
        /// <param name="uidoc">UIDOC</param>
        /// <param name="zone">目标分区ZoneInfoBase对象引用</param>
        /// <param name="listinfo">输出状态信息控件引用</param>
        /// <param name="processinfo">当前操作信息控件引用</param>
        #region 4D简单分区，嵌板和竖梃分析
        internal static void FnResolveSimpleZone(UIDocument uidoc, ZoneInfoBase zone, ref ListBox listinfo, ref Label processinfo)
        {
            var doc = uidoc.Document;
            uidoc.Selection.Elements.Clear();
            IOrderedEnumerable<CurtainPanelInfo> _panelsinzone = null;
            IOrderedEnumerable<MullionInfo> _mullionsinzone = null;

            var zsi = Global.DocContent.ZoneScheduleSimpleList.FirstOrDefault(zs => zs.ZoneCode == zone.ZoneCode);
            int zonedays = (zsi.ZoneFinish - zsi.ZoneStart).Days + 1;
            int zonehours = zonedays * Global.OptionHoursPerDay;

            #region CurtainPanelList 排序
            if (zone.ZoneDirection == "S")
                _panelsinzone = Global.DocContent.CurtainPanelList
                    .Where(p => p.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                    .OrderBy(p1 => Math.Round(p1.INF_OriginZ_Metric / Constants.RVTPrecision))
                    .ThenBy(p2 => Math.Round(p2.INF_OriginX_Metric / Constants.RVTPrecision));
            if (zone.ZoneDirection == "N")
                _panelsinzone = Global.DocContent.CurtainPanelList
                    .Where(p => p.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                    .OrderBy(p1 => Math.Round(p1.INF_OriginZ_Metric / Constants.RVTPrecision))
                    .OrderByDescending(p2 => Math.Round(p2.INF_OriginX_Metric / Constants.RVTPrecision));
            if (zone.ZoneDirection == "E")
                _panelsinzone = Global.DocContent.CurtainPanelList
                    .Where(p => p.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                    .OrderBy(p1 => Math.Round(p1.INF_OriginZ_Metric / Constants.RVTPrecision))
                    .ThenBy(p2 => Math.Round(p2.INF_OriginY_Metric / Constants.RVTPrecision));
            if (zone.ZoneDirection == "W")
                _panelsinzone = Global.DocContent.CurtainPanelList
                    .Where(p => p.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                    .OrderBy(p1 => Math.Round(p1.INF_OriginZ_Metric / Constants.RVTPrecision))
                    .OrderByDescending(p2 => Math.Round(p2.INF_OriginY_Metric / Constants.RVTPrecision));

            #endregion

            double v_hours_per_panel = 1.0 * zonehours / _panelsinzone.Count();
            int pindex = 0;
            //確定分區內嵌板數據及排序
            foreach (CurtainPanelInfo _pi in _panelsinzone)
            {
                processinfo.Content = $"當前處理進度：[分區：{zone.ZoneCode}] - [幕墻嵌板：{_pi.INF_ElementId}({++pindex}/{_panelsinzone.Count()})]";
                listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - SORT: P/{_pi.INF_ElementId}...");
                _pi.INF_Index = pindex;
                _pi.INF_Code = $"CW-{_pi.INF_Type:00}-{_pi.INF_Level:00}-{_pi.INF_Direction}{_pi.INF_System}-{pindex:0000}";
                _pi.INF_TaskStart = GetDeadTime(zsi.ZoneStart, v_hours_per_panel * (pindex - 1));
                _pi.INF_TaskFinish = GetDeadTime(zsi.ZoneStart, v_hours_per_panel * pindex);
            }

            #region MullionList 排序
            if (zone.ZoneDirection == "S")
                _mullionsinzone = Global.DocContent.MullionList
                    .Where(m => m.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                    .OrderByDescending(m1 => m1.INF_Type) //先8立柱后7橫樑
                    .OrderBy(m2 => Math.Round(m2.INF_OriginZ_Metric / Constants.RVTPrecision))
                    .ThenBy(m3 => Math.Round(m3.INF_OriginX_Metric / Constants.RVTPrecision));
            if (zone.ZoneDirection == "N")
                _mullionsinzone = Global.DocContent.MullionList
                    .Where(m => m.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                    .OrderByDescending(m1 => m1.INF_Type) //先8立柱后7橫樑
                    .OrderBy(m2 => Math.Round(m2.INF_OriginZ_Metric / Constants.RVTPrecision))
                    .OrderByDescending(m3 => Math.Round(m3.INF_OriginX_Metric / Constants.RVTPrecision));
            if (zone.ZoneDirection == "E")
                _mullionsinzone = Global.DocContent.MullionList
                    .Where(m => m.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                    .OrderByDescending(m1 => m1.INF_Type) //先8立柱后7橫樑
                    .OrderBy(m2 => Math.Round(m2.INF_OriginZ_Metric / Constants.RVTPrecision))
                    .ThenBy(m3 => Math.Round(m3.INF_OriginY_Metric / Constants.RVTPrecision));
            if (zone.ZoneDirection == "W")
                _mullionsinzone = Global.DocContent.MullionList
                    .Where(m => m.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                    .OrderByDescending(m1 => m1.INF_Type) //先8立柱后7橫樑
                    .OrderBy(m2 => Math.Round(m2.INF_OriginZ_Metric / Constants.RVTPrecision))
                    .OrderByDescending(m3 => Math.Round(m3.INF_OriginY_Metric / Constants.RVTPrecision));

            #endregion

            double v_hours_per_mullion = 1.0 * zonehours / _mullionsinzone.Count();
            int tindex = 0;
            pindex = 0;
            //確定分區內V竪梃(8)數據及排序
            foreach (MullionInfo _mi in _mullionsinzone.Where(m => m.INF_Type == 8))
            {
                processinfo.Content = $"當前處理進度：[分區：{zone.ZoneCode}] - [V 幕墻竪梃：{_mi.INF_ElementId}({++pindex}:{++tindex}/{_mullionsinzone.Count()})]";
                listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - CALC: M/V{_mi.INF_ElementId}...");
                _mi.INF_Index = pindex;
                _mi.INF_Code = $"CW-{_mi.INF_Type:00}-{_mi.INF_Level:00}-{_mi.INF_Direction}{_mi.INF_System}-{tindex:0000}";

                _mi.INF_TaskStart = GetDeadTime(zsi.ZoneStart, v_hours_per_mullion * (tindex - 1));
                _mi.INF_TaskFinish = GetDeadTime(zsi.ZoneStart, v_hours_per_mullion * tindex);
            }
            pindex = 0;
            //確定分區內V竪梃(8)數據及排序
            foreach (MullionInfo _mi in _mullionsinzone.Where(m => m.INF_Type < 8))
            {
                _mi.INF_Type = 7; //临时处理
                processinfo.Content = $"當前處理進度：[分區：{zone.ZoneCode}] - [H 幕墻竪梃：{_mi.INF_ElementId}({++pindex}:{++tindex}/{_mullionsinzone.Count()})]";
                listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - CALC: M/H{_mi.INF_ElementId}...");
                _mi.INF_Index = pindex;
                _mi.INF_Code = $"CW-{_mi.INF_Type:00}-{_mi.INF_Level:00}-{_mi.INF_Direction}{_mi.INF_System}-{tindex:0000}";

                _mi.INF_TaskStart = GetDeadTime(zsi.ZoneStart, v_hours_per_mullion * (tindex - 1));
                _mi.INF_TaskFinish = GetDeadTime(zsi.ZoneStart, v_hours_per_mullion * tindex);
            }
        }

        #endregion

        #region 4D设计分区，嵌板和构件分析
        /// <summary>
        /// 设计分区，嵌板和构件分析
        /// </summary>
        /// <param name="uidoc">UIDOC</param>
        /// <param name="zone">目标分区ZoneInfoBase对象引用</param>
        /// <param name="txt_curr_ele">状态文字控件引用-当前处理对象</param>
        /// <param name="txt_curr_op">状态文字控件引用-当前操作</param>
        /// <param name="progbar_curr">当前状态进度条控件引用</param>
        /// <param name="listinfo">输出状态信息控件引用</param>
        public static void FnResolveZone(UIDocument uidoc, ZoneInfoBase zone, ref Label txt_curr_ele, ref Label txt_curr_op, ref ProgressBar progbar_curr, ref ListBox listinfo)
        {
            var doc = uidoc.Document;
            bool _haserror = false;
            uidoc.Selection.Elements.Clear();

            IOrderedEnumerable<CurtainPanelInfo> _panelsinzone = null;
            IOrderedEnumerable<ScheduleElementInfo>[] _sesinzone = new IOrderedEnumerable<ScheduleElementInfo>[3];

            var zsi = Global.DocContent.ZoneScheduleLayerList.FirstOrDefault(zs => zs.ZoneUniversalCode == zone.ZoneCode);

            List<ScheduleElementInfo> p_ScheduleElementList = new List<ScheduleElementInfo>();
            listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - CALC: Z/{zone.ZoneCode}.");
            System.Windows.Forms.Application.DoEvents();

            #region CurtainPanelList 排序
            switch (zone.ZoneDirection)
            {
                case "S":
                    listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - SORT: Z/{zone.ZoneCode}, X+, Z+...");
                    System.Windows.Forms.Application.DoEvents();
                    _panelsinzone = Global.DocContent.CurtainPanelList
                        .Where(p => p.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                        .OrderBy(p1 => Math.Round(p1.INF_OriginZ_Metric / Constants.RVTPrecision))
                        .ThenBy(p2 => Math.Round(p2.INF_OriginX_Metric / Constants.RVTPrecision));
                    break;
                case "N":
                    listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - SORT: Z/{zone.ZoneCode}, X-, Z+...");
                    System.Windows.Forms.Application.DoEvents();
                    _panelsinzone = Global.DocContent.CurtainPanelList
                        .Where(p => p.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                        .OrderBy(p1 => Math.Round(p1.INF_OriginZ_Metric / Constants.RVTPrecision))
                        .OrderByDescending(p2 => Math.Round(p2.INF_OriginX_Metric / Constants.RVTPrecision));
                    break;
                case "E":
                    listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - SORT: Z/{zone.ZoneCode}, Y+, Z+...");
                    System.Windows.Forms.Application.DoEvents();
                    _panelsinzone = Global.DocContent.CurtainPanelList
                        .Where(p => p.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                        .OrderBy(p1 => Math.Round(p1.INF_OriginZ_Metric / Constants.RVTPrecision))
                        .ThenBy(p2 => Math.Round(p2.INF_OriginY_Metric / Constants.RVTPrecision));
                    break;
                case "W":
                    listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - SORT: Z/{zone.ZoneCode}, Y-, Z+...");
                    System.Windows.Forms.Application.DoEvents();
                    _panelsinzone = Global.DocContent.CurtainPanelList
                        .Where(p => p.INF_HostZoneInfo.ZoneCode == zone.ZoneCode)
                        .OrderBy(p1 => Math.Round(p1.INF_OriginZ_Metric / Constants.RVTPrecision))
                        .OrderByDescending(p2 => Math.Round(p2.INF_OriginY_Metric / Constants.RVTPrecision));
                    break;
                default: break;
            }
            #endregion

            progbar_curr.Maximum = _panelsinzone.Count();
            progbar_curr.Value = 0;

            int pindex = 0;
            //確定分區內嵌板數據及排序
            foreach (CurtainPanelInfo _pi in _panelsinzone)
            {
                txt_curr_ele.Content = $"Z/{zone.ZoneCode}.P/{_pi.INF_ElementId}, {_pi.INF_Code}";
                txt_curr_op.Content = "RESOLVE/P";
                progbar_curr.Value++;
                System.Windows.Forms.Application.DoEvents();
                Autodesk.Revit.DB.Panel _p = doc.GetElement(new ElementId(_pi.INF_ElementId)) as Autodesk.Revit.DB.Panel;

                _pi.INF_Index = pindex;
                _pi.INF_Code = $"CW-{_pi.INF_Type:00}-{_pi.INF_Level:00}-{_pi.INF_Direction}{_pi.INF_System}-{pindex:0000}";

                //確定嵌板內明細構件數據
                var p_subs = _p.GetSubComponentIds();
                foreach (ElementId eid in p_subs)
                {
                    txt_curr_op.Content = "RESOLVE/P.E";
                    System.Windows.Forms.Application.DoEvents();

                    ScheduleElementInfo _sei = new ScheduleElementInfo();
                    Element __element = (doc.GetElement(eid));
                    _sei.INF_ElementId = eid.IntegerValue;
                    _sei.INF_Name = __element.Name;
                    _sei.INF_ZoneCode = zone.ZoneCode;
                    _sei.INF_ZoneIndex = zone.ZoneIndex;
                    _sei.INF_HostZoneInfo = zone;
                    _sei.INF_Level = zone.ZoneLevel;
                    _sei.INF_Direction = zone.ZoneDirection;
                    _sei.INF_System = zone.ZoneSystem;
                    _sei.INF_HostCurtainPanel = _pi;

                    #region 确定分项参数 + 工序层级
                    Parameter _parameter;
                    _parameter = __element.get_Parameter("分项");
                    if (_parameter != null)
                    {
                        if ((_parameter = __element.get_Parameter("分项")).HasValue)
                        {
                            if (int.TryParse(_parameter.AsString(), out int _type))
                            {
                                ElementClass __sec;
                                if ((__sec = Global.ElementClassList.Find(ec => ec.EClassIndex == _type && ec.IsScheduled)) != null)
                                {
                                    _sei.INF_Type = _type;
                                    _sei.INF_TaskLevel = __sec.ETaskLevel;
                                    _sei.INF_TaskSubLevel = __sec.ETaskSubLevel;
                                    _sei.INF_IsScheduled = true;
                                }
                                else
                                {
                                    _sei.INF_Type = -11;
                                    continue;
                                }
                            }
                            else
                            {
                                _sei.INF_Type = -1;
                                _haserror = true;
                                _sei.INF_ErrorInfo = "构件[分项]参数错误(INF_Type)(非整数值)";
                                listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: VALUE TYPE, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, E/{_sei.INF_ElementId}, {_sei.INF_Name}.");
                                uidoc.Selection.Elements.Add(__element);
                                continue;
                            }
                        }
                        else
                        {
                            _sei.INF_Type = -2;
                            _haserror = true;
                            _sei.INF_ErrorInfo = "构件[分项]参数无参数值(INF_Type)";
                            listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: NO VALUE TYPE, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, E/{_sei.INF_ElementId}, {_sei.INF_Name}.");
                            uidoc.Selection.Elements.Add(__element);
                            continue;
                        }
                    }
                    else
                    {
                        _sei.INF_Type = -3;
                        _haserror = true;
                        _sei.INF_ErrorInfo = "构件[分项]参数项未设置(INF_Type)";
                        listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: PARAM NOTSET, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, {_sei.INF_ElementId}.");
                        uidoc.Selection.Elements.Add(__element);
                        continue;
                    }
                    #endregion

                    #region 判断門窗
                    /****
                    if (_schedule_element.INF_Type > 20 && _schedule_element.INF_Type < 30) 
                    {
                        _schedule_element.INF_HasDeepElements = true;
                        _schedule_element.INF_DeepElements = new List<DeepElementInfo>();
                        var _deepids = ((FamilyInstance)_element).GetSubComponentIds();
                        foreach (ElementId sid in _deepids)
                        {
                            DeepElementInfo _deep_element = new DeepElementInfo();
                            Element _de = (doc.GetElement(sid));
                            _deep_element.INF_ElementId = sid.IntegerValue;
                            _deep_element.INF_Name = _de.Name;
                            _deep_element.INF_HostCurtainPanel = p;
                            _deep_element.INF_HostScheduleElement = _schedule_element;

                            Parameter _sparam;
                            if ((_sparam = _de.get_Parameter("分项")).HasValue)
                            {
                                if (int.TryParse(_sparam.AsString(), out int _type)) _deep_element.INF_Type = _type;
                                else
                                {
                                    _deep_element.INF_Type = -10;
                                    listinfo.SelectedIndex = listinfo.Items.Add($"[分項]參數錯誤");
                                    listinfo.SelectedIndex = listinfo.Items.Add($"[{_deep_element.INF_HostCurtainPanel.INF_ElementId}][{_deep_element.INF_ElementId}]:[分项]参数错误...");
                                }
                            }
                            else
                            {
                                _deep_element.INF_Type = -10;
                                listinfo.SelectedIndex = listinfo.Items.Add($"[分項]參數未設置");
                                listinfo.SelectedIndex = listinfo.Items.Add($"[{_deep_element.INF_HostCurtainPanel.INF_ElementId}][{_deep_element.INF_ElementId}]:[分项]参数错误...");
                            }

                            _deep_element.INF_Level = p.INF_Level;
                            _deep_element.INF_System = p.INF_System;
                            _deep_element.INF_Direction = p.INF_Direction;
                            _deep_element.INF_ZoneID = p.INF_ZoneID;
                            _deep_element.INF_ZoneCode = p.INF_ZoneCode;
                            _deep_element.INF_Type_ZoneCode = $"[{_deep_element.INF_ZoneCode}][{_deep_element.INF_Type:00}]";
                            _schedule_element.INF_DeepElements.Add(_deep_element);
                        }
                    }
                    ****/
                    #endregion

                    XYZ _xyzOrigin = ((FamilyInstance)__element).GetTotalTransform().Origin;
                    _sei.INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.X);
                    _sei.INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Y);
                    _sei.INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Z);

                    //processinfo.Content = $"當前處理進度：[分區：{zone.ZoneCode}] - [幕墻嵌板：{_pi.INF_ElementId}({pindex}/{_panelsinzone.Count()})] - [構件：{_sei.INF_ElementId}]";
                    System.Windows.Forms.Application.DoEvents();
                    p_ScheduleElementList.Add(_sei);
                    //Global.DocContent.ScheduleElementList.Add(_sei);
                }
            }

            #region 参数错误构件保存选择集
            if (!uidoc.Selection.Elements.IsEmpty)
            {
                using (Transaction trans = new Transaction(doc, "CreateZoneErrorElementGroup"))
                {
                    trans.Start();
                    var sfe = SelectionFilterElement.Create(doc, $"ERROR-{zone.ZoneCode}");
                    sfe.AddSet(uidoc.Selection.GetElementIds());
                    trans.Commit();
                    listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - WRITE: Z/[{zone.ZoneCode}, ERR#{zone.ZoneCode}.");
                }
            }
            #endregion

            if (_haserror)
            {
                if (MessageBox.Show($"有部分构件存在参数错误，是否继续处理数据？", "构件参数错误...", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No)
                {
                    return;
                }
            }


            #region 明细构件 排序
            var lvgroupsScheduleElement = p_ScheduleElementList.ToLookup(se => se.INF_TaskLevel);
            foreach (var lvgroup in lvgroupsScheduleElement)
            {
                switch (zone.ZoneDirection)
                {
                    case "S":
                        listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - SORT: Z/{zone.ZoneCode}@{lvgroup.Key}, X+, Z+...");
                        System.Windows.Forms.Application.DoEvents();
                        _sesinzone[lvgroup.Key] = lvgroup
                            .OrderBy(m1 => m1.INF_TaskSubLevel)
                            .ThenBy(e1 => Math.Round(e1.INF_HostCurtainPanel.INF_OriginZ_Metric / Constants.RVTPrecision))
                            .ThenBy(e2 => Math.Round(e2.INF_HostCurtainPanel.INF_OriginX_Metric / Constants.RVTPrecision))
                            .ThenBy(e3 => Math.Round(doc.GetElement(new ElementId(e3.INF_ElementId)).get_BoundingBox(doc.ActiveView).Min.Z / Constants.RVTPrecision))
                            .ThenBy(e4 => Math.Round(doc.GetElement(new ElementId(e4.INF_ElementId)).get_BoundingBox(doc.ActiveView).Min.X / Constants.RVTPrecision));
                        break;
                    case "N":
                        listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - SORT: Z/{zone.ZoneCode}@{lvgroup.Key}, X-, Z+...");
                        System.Windows.Forms.Application.DoEvents();
                        _sesinzone[lvgroup.Key] = lvgroup
                            .OrderByDescending(m1 => m1.INF_TaskSubLevel)
                            .ThenBy(e1 => Math.Round(e1.INF_HostCurtainPanel.INF_OriginZ_Metric / Constants.RVTPrecision))
                            .ThenByDescending(e2 => Math.Round(e2.INF_HostCurtainPanel.INF_OriginX_Metric / Constants.RVTPrecision))
                            .ThenBy(e3 => Math.Round(doc.GetElement(new ElementId(e3.INF_ElementId)).get_BoundingBox(doc.ActiveView).Min.Z / Constants.RVTPrecision))
                            .ThenByDescending(e4 => Math.Round(doc.GetElement(new ElementId(e4.INF_ElementId)).get_BoundingBox(doc.ActiveView).Min.X / Constants.RVTPrecision));
                        break;
                    case "E":
                        listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - SORT: Z/{zone.ZoneCode}@{lvgroup.Key}, Y+, Z+...");
                        System.Windows.Forms.Application.DoEvents();
                        _sesinzone[lvgroup.Key] = lvgroup
                            .OrderByDescending(m1 => m1.INF_TaskSubLevel)
                            .ThenBy(e1 => Math.Round(e1.INF_HostCurtainPanel.INF_OriginZ_Metric / Constants.RVTPrecision))
                            .ThenBy(e2 => Math.Round(e2.INF_HostCurtainPanel.INF_OriginY_Metric / Constants.RVTPrecision))
                            .ThenBy(e3 => Math.Round(doc.GetElement(new ElementId(e3.INF_ElementId)).get_BoundingBox(doc.ActiveView).Min.Z / Constants.RVTPrecision))
                            .ThenBy(e4 => Math.Round(doc.GetElement(new ElementId(e4.INF_ElementId)).get_BoundingBox(doc.ActiveView).Min.Y / Constants.RVTPrecision));
                        break;
                    case "W":
                        listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - SORT: Z/{zone.ZoneCode}@{lvgroup.Key}, Y-, Z+...");
                        System.Windows.Forms.Application.DoEvents();
                        _sesinzone[lvgroup.Key] = lvgroup
                            .OrderByDescending(m1 => m1.INF_TaskSubLevel)
                            .ThenBy(e1 => Math.Round(e1.INF_HostCurtainPanel.INF_OriginZ_Metric / Constants.RVTPrecision))
                            .ThenByDescending(e2 => Math.Round(e2.INF_HostCurtainPanel.INF_OriginY_Metric / Constants.RVTPrecision))
                            .ThenBy(e3 => Math.Round(doc.GetElement(new ElementId(e3.INF_ElementId)).get_BoundingBox(doc.ActiveView).Min.Z / Constants.RVTPrecision))
                            .ThenByDescending(e4 => Math.Round(doc.GetElement(new ElementId(e4.INF_ElementId)).get_BoundingBox(doc.ActiveView).Min.Y / Constants.RVTPrecision));
                        break;
                    default:
                        break;
                }

                int eindexinzone = 0;
                foreach (var ele in _sesinzone[lvgroup.Key])
                {
                    eindexinzone++;
                    ele.INF_Code = $"CW-{ele.INF_Type:00}-{ele.INF_Level:00}-{ele.INF_Direction}{ele.INF_System}-{eindexinzone:0000}";//构件编码
                }
                listinfo.SelectedIndex = listinfo.Items.Add($"{DateTime.Now:HH:mm:ss} - CODE: Z/{zone.ZoneCode}@{lvgroup.Key}, E/{_sesinzone[lvgroup.Key].Count()}...");
            }
            Global.DocContent.ScheduleElementList.AddRange(p_ScheduleElementList);
            #endregion

        }

        #endregion
        public static void FnContentSerialize()
        {
            if (File.Exists(Global.DataFile)) File.Delete(Global.DataFile);
            using (FileStream fs = new FileStream(Global.DataFile, FileMode.Create))
            {
                Serializer.Serialize(fs, Global.DocContent);
            }
        }

        public static void FnContentDeserialize(string datafilename)
        {
            if (File.Exists(datafilename))
            {
                using (Stream stream = new FileStream(datafilename, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    Global.DocContent = Serializer.Deserialize<DocumentContent>(stream);
                }
            }
        }

        public static void FnContentBackup() { if (File.Exists(Global.DataFile)) File.Move(Global.DataFile, $"{Global.DataFile}.{DateTime.Now:yyyyMMddHHmmss}.bak"); }

        public static void FnContentSerializeWithBackup()
        {
            FnContentBackup();
            using (FileStream fs = new FileStream(Global.DataFile, FileMode.Create))
            {
                Serializer.Serialize(fs, Global.DocContent);
            }
        }

        public static void AddRangeEx<T>(this List<T> SourceList, IEnumerable<T> SubList) where T : ElementInfoBase
        {
            foreach (var ei in SubList)
            {
                SourceList.RemoveAll(__e => __e.INF_ElementId == ei.INF_ElementId);
                SourceList.Add(ei);
            }
        }

        public static void AddEx<T>(this ObservableCollection<T> SourceList, T NewItem) where T : ZoneInfoBase
        {
            foreach (ZoneInfoBase z in SourceList)
                if (z.ZoneCode == NewItem.ZoneCode)
                {
                    SourceList.Remove((T)z);
                    break;
                }
            SourceList.Add(NewItem);
            return;
        }

        public static DateTime GetDeadTime(DateTime timestart, double durationhours)
        {
            int num_days = Convert.ToInt32(Math.Floor(durationhours / Global.OptionHoursPerDay));
            double rest_hours = durationhours - num_days * Global.OptionHoursPerDay;
            if (rest_hours > 4) rest_hours++;
            return (timestart + new TimeSpan(num_days, 0, 0, 0)).AddHours(8 + rest_hours);
        }
    }

    public static class StreamHelper
    {
        public static MemoryStream ByteArrayToMemeryStream(byte[] barray)
        {
            MemoryStream stream = new MemoryStream();
            stream.Write(barray, 0, barray.Length);
            return stream;
        }

        public static byte[] MemeryStreamToByteArray(this Stream input)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                input.CopyTo(ms);
                return ms.ToArray();
            }
        }
    }

    public class PathButton : Button
    {
        public static readonly DependencyProperty DataProperty = DependencyProperty.Register("Data", typeof(Geometry), typeof(PathButton), new PropertyMetadata(new PropertyChangedCallback(OnDataChanged)));

        private static void OnDataChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            PathButton button = d as PathButton;
            button.Content = new System.Windows.Shapes.Path()
            {
                Data = e.NewValue as Geometry,
                Fill = Brushes.Gray,
                Stretch = Stretch.Uniform,
                StrokeThickness = 1,
                Stroke = button.Foreground,
                HorizontalAlignment = HorizontalAlignment.Left,
                Width = 30,
                Height = 30
            };
        }
        public Geometry Data { get { return (Geometry)GetValue(DataProperty); } set { SetValue(DataProperty, value); } }
    }

}
