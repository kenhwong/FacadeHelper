using AqlaSerializer;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using RvtDB = Autodesk.Revit.DB;

namespace FacadeHelper
{
    /// <summary>
    /// Interaction logic for ProcessElements.xaml
    /// </summary>
    public partial class ProcessElements : UserControl
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Document doc;

        Progress<int> progress;
        private bool isRealTimeProgress = true;
        private UC_Process CurrentProcess = UC_Process.Other;

        private List<CurtainPanelInfo> SelectedCurtainPanelList = new List<CurtainPanelInfo>();

        private enum UC_Process : int
        {
            Other = 0,
            DataGrid_Loading_Level = 1,
            DataGrid_Loading_CurtainSystem = 2,
            DataGrid_Loading_Wall = 3,
            DataGrid_Loading_Panel = 4,
            DataGrid_Loading_SubElement = 5,
            DataGrid_Check_Host = 6
        }

        public ProcessElements(ExternalCommandData commandData)
        {
            InitializeComponent();
            InitializeCommand();
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;
            Global.DataFile = Path.Combine(Path.GetDirectoryName(doc.PathName), $"{Path.GetFileNameWithoutExtension(doc.PathName)}.data");

            checkRealTimeProgress.IsChecked = isRealTimeProgress;

            if (Global.DocContent is null)
            {
                Global.DocContent = new DocumentContent();
                if (File.Exists(Global.DataFile))
                {
                    using (Stream stream = new FileStream(Global.DataFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        Global.DocContent = Serializer.Deserialize<DocumentContent>(stream);
                    }
                    if (Global.DocContent.LevelDictionary.Count > 0) Global.DocContent.IsLevelsLoaded = true;
                    if (Global.DocContent.CurtainSystemList.Count > 0) Global.DocContent.IsCurtainsystemsLoaded = true;
                    if (Global.DocContent.WallList.Count > 0) Global.DocContent.IsWallsLoaded = true;
                    if (Global.DocContent.CurtainPanelList.Count > 0) Global.DocContent.IsCurtainPanelsLoaded = true;
                    if (Global.DocContent.ScheduleElementList.Count > 0) Global.DocContent.IsDeepElementsInPanelsLoaded = true;
                }
            }
            InitProjectInfo();
        }

        #region 初始化 Command
        private RoutedCommand cmdLoadLevel = new RoutedCommand();
        private RoutedCommand cmdLoadCurtainSystem = new RoutedCommand();
        private RoutedCommand cmdLoadWall = new RoutedCommand();
        private RoutedCommand cmdLoadPanel = new RoutedCommand();
        private RoutedCommand cmdLoadSubElement = new RoutedCommand();

        private RoutedCommand cmdNavLevel = new RoutedCommand();
        private RoutedCommand cmdNavCurtainSystem = new RoutedCommand();
        private RoutedCommand cmdNavWall = new RoutedCommand();
        private RoutedCommand cmdNavPanel = new RoutedCommand();
        private RoutedCommand cmdNavSubElement = new RoutedCommand();

        private RoutedCommand cmdApplyParameters = new RoutedCommand();

        private RoutedCommand cmdModelInit = new RoutedCommand();

        private void InitializeCommand()
        {
            CommandBinding cbLoadLevel = new CommandBinding(cmdLoadLevel, cbLoadlevel_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbLoadCurtainsystem = new CommandBinding(cmdLoadCurtainSystem, cbLoadCurtainSystem_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbLoadWall = new CommandBinding(cmdLoadWall, cbLoadWall_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbLoadPanel = new CommandBinding(cmdLoadPanel, cbLoadPanel_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbLoadSubelement = new CommandBinding(cmdLoadSubElement, cbLoadSubElement_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbNavLevel = new CommandBinding(cmdNavLevel,
                (sender, e) => { griddataLevel.ItemsSource = Global.DocContent.LevelList; },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbNavCurtainSystem = new CommandBinding(cmdNavCurtainSystem,
                (sender, e) => { griddataCurtainSystem.ItemsSource = e.Parameter as IGrouping<int, CurtainSystemInfo>; },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbNavwall = new CommandBinding(cmdNavWall,
                (sender, e) => { griddataWall.ItemsSource = e.Parameter as IGrouping<int, WallInfo>; },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbNavPanel = new CommandBinding(cmdNavPanel,
                (sender, e) => { griddataPanel.ItemsSource = e.Parameter as IGrouping<string, CurtainPanelInfo>; },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbNavSubelement = new CommandBinding(cmdNavSubElement,
                (sender, e) =>
                {
                    griddataSubelement.ItemsSource = (e.Parameter as IGrouping<string, GeneralElementInfo>).Join(
                        Global.DocContent.CurtainPanelList,
                        x => x.INF_HostCurtainPanel.INF_ElementId,
                        y => y.INF_ElementId,
                        (x, y) => new
                        {
                            Element_ID = x.INF_ElementId,
                            Element_Name = x.INF_Name,
                            Element_Type = x.INF_Type,
                            Element_System = x.INF_System,
                            Element_Direction = x.INF_Direction,
                            Element_Level = x.INF_Level,
                            Panel_ID = y.INF_ElementId,
                            Panel_ZoneCode = y.INF_ZoneCode,
                            Panel_Code = y.INF_Code,
                            Element_ZoneID = x.INF_ZoneID,
                            Element_ZoneCode = x.INF_ZoneCode,
                            Element_Code = x.INF_Code
                        });
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbApplyParameters = new CommandBinding(cmdApplyParameters, cbApplyParameters_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbModelInit = new CommandBinding(cmdModelInit, cbModelInit_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            bnDataLevel.Command = cmdLoadLevel;
            bnDataCurtainSystem.Command = cmdLoadCurtainSystem;
            bnDataWall.Command = cmdLoadWall;
            bnDataPanel.Command = cmdLoadPanel;
            bnDataSubElement.Command = cmdLoadSubElement;

            bnApplyParameters.Command = cmdApplyParameters;
            bnModelInit.Command = cmdModelInit;

            ProcEle.CommandBindings.AddRange(new CommandBinding[] { cbLoadLevel, cbLoadCurtainsystem, cbLoadWall, cbLoadPanel, cbLoadSubelement });
            ProcEle.CommandBindings.AddRange(new CommandBinding[] { cbNavLevel, cbNavCurtainSystem, cbNavwall, cbNavPanel, cbNavSubelement });
            ProcEle.CommandBindings.Add(cbApplyParameters);
            ProcEle.CommandBindings.Add(cbModelInit);
        }

        private void cbModelInit_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            listInformation.Items.Clear();

            //Selection selection = uidoc.Selection;
            //ElementSet collection = selection.Elements;

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"当前无选择。");
                return;
            }
            FilteredElementCollector collector = new FilteredElementCollector(doc, ids);
            LogicalAndFilter cwpanel_InstancesFilter =
                new LogicalAndFilter(
                    new ElementClassFilter(typeof(FamilyInstance)),
                    new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels));
            var eles = collector
                .WherePasses(cwpanel_InstancesFilter)
                .Where(x => (x as FamilyInstance).Symbol.Family.Name != "系统嵌板");

            if (eles.Count() == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"当前未选择幕墙嵌板。");
                return;
            }

            //FilterRule fr = ParameterFilterRuleFactory.CreateSharedParameterApplicableRule("分区区号");
            //ElementParameterFilter pf = new ElementParameterFilter(fr);
            Global.DocContent.CurtainPanelList.Clear();
            //Global.DocContent.ScheduleElementList.Clear();
            //Global.DocContent.GeneralElementFullList.Clear();
            //Global.DocContent.DeepElementList.Clear();
            int errorcount_zonecode = 0;
            string _zcode = "";
            foreach (var _ele in eles)
            {
                CurtainPanelInfo _gp = new CurtainPanelInfo(_ele as RvtDB.Panel);
                _gp.INF_Level = GetLevelByElevation(_gp.INF_OriginZ_Metric).INF_LevelIndex;
                SelectedCurtainPanelList.Add(_gp);
                Parameter _param = _ele.get_Parameter("分区区号");
                if (!_param.HasValue)
                    listInformation.SelectedIndex = listInformation.Items.Add($"{++errorcount_zonecode}/ 幕墙嵌板[{_ele.Id.IntegerValue}] 未设置分区区号");
                else
                {
                    _gp.INF_ZoneCode = _param.AsString();
                    if (_zcode == "")
                        _zcode = _gp.INF_ZoneCode;
                    else if (_zcode != _gp.INF_ZoneCode)
                        listInformation.SelectedIndex = listInformation.Items.Add($"{++errorcount_zonecode}/ 幕墙嵌板[{_ele.Id.IntegerValue}] 分区区号{_gp.INF_ZoneCode}差异({_zcode})");
                }

            }
            if (errorcount_zonecode == 0) listInformation.SelectedIndex = listInformation.Items.Add($"选择的{eles.Count()}幕墙嵌板均已设置相同的分区区号");
            griddataSelectedPanel.ItemsSource = SelectedCurtainPanelList;

            uidoc.Selection.Elements.Clear();
            foreach (var _ele in eles) uidoc.Selection.Elements.Add(_ele);
            using (Transaction trans = new Transaction(doc, "CreateGroup"))
            {
                trans.Start();
                var sfe = SelectionFilterElement.Create(doc, _zcode);
                sfe.AddSet(uidoc.Selection.GetElementIds());
                trans.Commit();
            }

        }

        public class PanelSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) { return elem.Category.Name == "Panel"; }
            public bool AllowReference(Reference reference, XYZ position) { return false; }
        }

        private void cbApplyParameters_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            e.Handled = true;
            ApplyParameters_CurtainPanels();
            ApplyParameters_SubElementsInPanels();

            //ContentSerialize();
        }

        private void cbLoadSubElement_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            e.Handled = true;
            Load_Data_Panels_And_SubElements(true);
        }

        private void cbLoadPanel_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            e.Handled = true;
            Load_Data_Panels_And_SubElements(false);
        }

        private void cbLoadWall_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            e.Handled = true;
            Load_Data_Walls();
            listInformation.SelectedIndex = listInformation.Items.Add($"提取当前模型中所有墻（幕墻）系统, 共{Global.DocContent.WallList.Count}...");
        }

        private void cbLoadCurtainSystem_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            e.Handled = true;
            Load_Data_CurtainSystems();
            listInformation.SelectedIndex = listInformation.Items.Add($"提取当前模型中所有幕墙系统, 共{Global.DocContent.CurtainSystemList.Count}...");
        }

        private void cbLoadlevel_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            e.Handled = true;
            Load_Data_Levels();
            listInformation.SelectedIndex = listInformation.Items.Add($"提取当前模型中所有标高, 共 {Global.DocContent.LevelList.Count} 個...");
        }

        #endregion
        private void InitProjectInfo()
        {
            if (Global.DocContent.CurtainPanelList.Count > 0) Global.DocContent.Lookup_CurtainPanels = Global.DocContent.CurtainPanelList.ToLookup(g => g.INF_ZoneCode);
            Global.DocContent.ParameterInfoList = ParameterHelper.RawGetProjectParametersInfo(doc);
            SetupProjectParameters();
            ProcCheckHost();
            Load_Tree_Levels();
            Load_Tree_CurtainSystems();
            Load_Tree_Panels();
            Load_Tree_SubElements();
        }

        private void ProcCheckHost()
        {
            #region 檢查幕墻系統/墙的容器分區數據
            FilteredElementCollector fecCS = new FilteredElementCollector(doc);
            ElementClassFilter filterCS = new ElementClassFilter(typeof(CurtainSystem));
            fecCS.WherePasses(filterCS);
            var _cslist = fecCS.ToElements();
            FilteredElementCollector fecWall = new FilteredElementCollector(doc);
            ElementCategoryFilter categoryfilterWall = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            ElementClassFilter classfilterWall = new ElementClassFilter(typeof(Wall));
            var _wallcalsslist = (new FilteredElementCollector(doc)).WherePasses(classfilterWall).ToElements();
            fecWall.WherePasses(categoryfilterWall);
            var _walllist = fecWall.ToElements();

            CurrentProcess = UC_Process.DataGrid_Check_Host;

            griddataZone.ItemsSource = null;

            if (Global.DocContent.CurtainSystemList.Count == 0)
            {
                Global.DocContent.CurtainSystemList.Clear();
                //Global.DocContent.IsCurtainsystemsLoaded = true;
                int i = 0;
                foreach (Element _e in _cslist)
                {
                    var _cs = new CurtainSystemInfo(_e as CurtainSystem);
                    if (_cs.INF_ErrorInfo.Contains("["))
                    {
                        Global.DocContent.IsCurtainsystemsLoaded = false;
                        listInformation.SelectedIndex = listInformation.Items.Add($"幕墻系統[{_cs.INF_ElementId}]錯誤：{_cs.INF_ErrorInfo}...");
                    }
                    Global.DocContent.CurtainSystemList.Add(_cs);

                    if (isRealTimeProgress)
                    {
                        progressDataCurtainSystem.Text = $"{++i}/{_cslist.Count()}";
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
            }
            if (Global.DocContent.CurtainSystemList.Count > 0) griddataZone.ItemsSource = new ObservableCollection<CurtainSystemInfo>(Global.DocContent.CurtainSystemList);

            if (Global.DocContent.WallList.Count == 0)
            {
                Global.DocContent.WallList.Clear();
                int i = 0;
                foreach (Element _e in _walllist)
                {
                    if (_e.GetType().Name == "WallType") continue;
                    var _wall = new WallInfo(_e as Wall);
                    if (_wall.INF_ErrorInfo.Contains("["))
                    {
                        Global.DocContent.IsCurtainsystemsLoaded = false;
                        listInformation.SelectedIndex = listInformation.Items.Add($"墙[{_wall.INF_ElementId}]錯誤：{_wall.INF_ErrorInfo}...");
                    }
                    Global.DocContent.WallList.Add(_wall);

                    if (isRealTimeProgress)
                    {
                        progressDataWall.Text = $"{++i}/{_walllist.Count()}";
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
            }
            if (Global.DocContent.WallList.Count > 0) griddataZone.ItemsSource = new ObservableCollection<WallInfo>(Global.DocContent.WallList);
            #endregion
        }

        /**
        private void DataGridHostCSItem_GotFocus(object sender, RoutedEventArgs e)
        {
            var hostrow = ((DataGridRow)sender);
            int _selectedcsid = int.Parse(((TextBlock)griddataCSZone.Columns[0].GetCellContent(hostrow)).Text);
        }

        private void DataGridHostWallItem_GotFocus(object sender, RoutedEventArgs e)
        {
            var hostrow = ((DataGridRow)sender);
            WallInfo _selectedcs = ((ContentPresenter)griddataWallZone.Columns[0].GetCellContent(hostrow)).Content as WallInfo;
        }
        **/

        private void ContentSerialize()
        {
            if (File.Exists(Global.DataFile)) File.Delete(Global.DataFile);
            using (FileStream fs = new FileStream(Global.DataFile, FileMode.Create))
            {
                Serializer.Serialize(fs, Global.DocContent);
            }
        }


        private LevelInfo GetLevelByElevation(double elev)
        {
            if (Global.DocContent.LevelList.Count == 0) return new LevelInfo { INF_LevelIndex = -99 };
            else if (elev < Global.DocContent.LevelList[0].INF_LevelElevation_Metric) return Global.DocContent.LevelList[0];
            else return Global.DocContent.LevelList.FindLast(x => x.INF_LevelElevation_Metric <= elev);

        }

        private void Load_Data_Levels()
        {
            #region 讀取標高數據
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            var _levelList = collector
                .OfClass(typeof(Level))
                .ToElements()
                .Where(x => !((Level)x).Name.Equals("地面"))
                .OrderBy(elev => ((Level)elev).Elevation);

            Global.DocContent.LevelList.Clear();
            int i = 0;
            listInformation.SelectedIndex = listInformation.Items.Add($"開始讀取項目標高數據...");
            foreach (RvtDB.Element _elvl in _levelList)
            {
                Level _lvl = _elvl as Level;
                double _velev = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _lvl.Elevation);
                var _li = new LevelInfo
                {
                    INF_ElementId = _lvl.Id.IntegerValue,
                    INF_LevelIndex = i + 1,
                    INF_LevelElevation_Metric = _velev,
                    INF_LevelElevation_US = _lvl.Elevation,
                    INF_LevelSign = _lvl.Name
                };
                Global.DocContent.LevelList.Add(_li);
                if (isRealTimeProgress)
                {
                    progressDataLevel.Text = $"{++i}/{_levelList.Count()}";
                    listInformation.SelectedIndex = listInformation.Items.Add($"[{_li.INF_LevelSign}]: {_li.INF_LevelElevation_Metric:N2}");
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            if (Global.DocContent.LevelList.Count > 0) Global.DocContent.IsLevelsLoaded = true;
            listInformation.SelectedIndex = listInformation.Items.Add($"讀取項目標高數據 [{Global.DocContent.LevelList.Count}] 完成");
            #endregion
            Load_Tree_Levels();
        }

        #region 左侧树展开控制
        private void ExpandNavTree(string expname)
        {
            switch (expname)
            {
                case "LEVEL":
                    expLevels.IsExpanded = true;
                    expCurtainSystem.IsExpanded = false;
                    expWall.IsExpanded = false;
                    expPanel.IsExpanded = false;
                    expSubElement.IsExpanded = false;
                    break;
                case "CS":
                    expLevels.IsExpanded = false;
                    expCurtainSystem.IsExpanded = true;
                    expWall.IsExpanded = false;
                    expPanel.IsExpanded = false;
                    expSubElement.IsExpanded = false;
                    break;
                case "WALL":
                    expLevels.IsExpanded = false;
                    expCurtainSystem.IsExpanded = false;
                    expWall.IsExpanded = true;
                    expPanel.IsExpanded = false;
                    expSubElement.IsExpanded = false;
                    break;
                case "PANEL":
                    expLevels.IsExpanded = false;
                    expCurtainSystem.IsExpanded = false;
                    expWall.IsExpanded = false;
                    expPanel.IsExpanded = true;
                    expSubElement.IsExpanded = false;
                    break;
                case "SE":
                    expLevels.IsExpanded = false;
                    expCurtainSystem.IsExpanded = false;
                    expWall.IsExpanded = false;
                    expPanel.IsExpanded = false;
                    expSubElement.IsExpanded = true;
                    break;
                default:
                    break;
            }
        }
        #endregion

        private void Load_Tree_Levels()
        {
            #region 生成標高數據分區
            if (Global.DocContent.IsLevelsLoaded)
                navLevels.Items.Clear();
            foreach (var lv in Global.DocContent.LevelList)
            {
                ListBoxItem _navitem_lvl = new ListBoxItem();
                //_nbi_cs.ImageSource = new BitmapImage(new Uri("pack://application:,,,/Images/private.png"));
                _navitem_lvl.Content = $"[Level {lv.INF_LevelSign}]";
                _navitem_lvl.Name = $"LEVEL_{lv.INF_LevelIndex:00}";
                _navitem_lvl.Tag = lv;

                MouseBinding mbind = new MouseBinding(cmdNavLevel, new MouseGesture(MouseAction.LeftDoubleClick));
                mbind.CommandParameter = _navitem_lvl;
                mbind.CommandTarget = griddataLevel;
                _navitem_lvl.InputBindings.Add(mbind);
                navLevels.Items.Add(_navitem_lvl);
            }
            #endregion
            ExpandNavTree("LEVEL");
            listInformation.SelectedIndex = listInformation.Items.Add($"生成項目標高分區列表 [{Global.DocContent.LevelList.Count}]");
        }

        private void Load_Data_CurtainSystems()
        {
            #region 讀取幕墻系統數據
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ElementClassFilter familyInstanceFilter = new ElementClassFilter(typeof(CurtainSystem));
            collector.WherePasses(familyInstanceFilter);

            var _cscoll = collector.ToElements();
            Global.DocContent.CurtainSystemList.Clear();
            Global.DocContent.IsCurtainsystemsLoaded = true;
            int i = 0;
            foreach (Element _e in _cscoll)
            {
                var _cs = new CurtainSystemInfo(_e as CurtainSystem);
                if (_cs.INF_ErrorInfo.Contains("["))
                {
                    Global.DocContent.IsCurtainsystemsLoaded = false;
                    listInformation.SelectedIndex = listInformation.Items.Add($"幕墻系統[{_cs.INF_ElementId}]錯誤：{_cs.INF_ErrorInfo}...");
                }
                Global.DocContent.CurtainSystemList.Add(_cs);

                if (isRealTimeProgress)
                {
                    progressDataCurtainSystem.Text = $"{++i}/{_cscoll.Count()}";
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            if (Global.DocContent.CurtainSystemList.Count <= 0) Global.DocContent.IsCurtainsystemsLoaded = false;
            listInformation.SelectedIndex = listInformation.Items.Add($"讀取幕墻系統數據 [{Global.DocContent.CurtainSystemList.Count}]");
            #endregion
            Load_Tree_CurtainSystems();
        }

        private void Load_Tree_CurtainSystems()
        {
            #region 生成幕墻系統分區
            if (Global.DocContent.IsCurtainsystemsLoaded)
            {
                navCurtainSystems.Items.Clear();
                foreach (var groupz in Global.DocContent.CurtainSystemList.OrderBy(o => o.INF_ZoneID).GroupBy(x => x.INF_ZoneID))
                {
                    ListBoxItem _navitem_cs = new ListBoxItem();
                    _navitem_cs.Content = $"分区 [{groupz.Key:00}]";
                    _navitem_cs.Name = $"ZONE_{groupz.Key:00}";
                    _navitem_cs.Tag = groupz;

                    MouseBinding mbind = new MouseBinding(cmdNavCurtainSystem, new MouseGesture(MouseAction.LeftDoubleClick));
                    mbind.CommandParameter = groupz;
                    mbind.CommandTarget = griddataCurtainSystem;
                    _navitem_cs.InputBindings.Add(mbind);
                    navCurtainSystems.Items.Add(_navitem_cs);
                }
            }
            #endregion
            ExpandNavTree("CS");
            listInformation.SelectedIndex = listInformation.Items.Add($"生成幕墻系統分區列表 [{navCurtainSystems.Items.Count}]");
        }

        private void Load_Data_Walls()
        {
            #region 讀取墻數據
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ElementClassFilter familyInstanceFilter = new ElementClassFilter(typeof(Wall));
            collector.WherePasses(familyInstanceFilter);

            var _wallcollection = collector.ToElements();
            Global.DocContent.WallList.Clear();
            Global.DocContent.IsWallsLoaded = true;
            int i = 0;
            foreach (RvtDB.Element _e in _wallcollection)
            {
                var _wall = new WallInfo(_e as Wall);
                if (_wall.INF_ErrorInfo.Contains("["))
                {
                    Global.DocContent.IsWallsLoaded = false;
                    listInformation.SelectedIndex = listInformation.Items.Add($"墻（幕墻）[{_wall.INF_ElementId}]錯誤：{_wall.INF_ErrorInfo}...");
                }
                Global.DocContent.WallList.Add(_wall);

                if (isRealTimeProgress)
                {
                    progressDataWall.Text = $"{++i}/{_wallcollection.Count()}";
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            if (Global.DocContent.WallList.Count <= 0) Global.DocContent.IsWallsLoaded = false;
            listInformation.SelectedIndex = listInformation.Items.Add($"讀取墻（幕墻）數據 [{Global.DocContent.WallList.Count}]");
            #endregion
            Load_Tree_Wall();
        }

        private void Load_Tree_Wall()
        {
            #region 生成墻分區
            if (Global.DocContent.IsWallsLoaded)
            {
                navWalls.Items.Clear();
                foreach (var groupz in Global.DocContent.WallList.OrderBy(o => o.INF_ZoneID).GroupBy(x => x.INF_ZoneID))
                {
                    ListBoxItem _navitem_wall = new ListBoxItem();
                    _navitem_wall.Content = $"分区 [{groupz.Key:00}]";
                    _navitem_wall.Name = $"ZONE_{groupz.Key:00}";
                    _navitem_wall.Tag = groupz;

                    MouseBinding mbind = new MouseBinding(cmdNavWall, new MouseGesture(MouseAction.LeftDoubleClick));
                    mbind.CommandParameter = groupz;
                    mbind.CommandTarget = griddataWall;
                    _navitem_wall.InputBindings.Add(mbind);
                    navWalls.Items.Add(_navitem_wall);
                }
            }
            #endregion
            ExpandNavTree("WALL");
            listInformation.SelectedIndex = listInformation.Items.Add($"生成墻（幕墻）分區列表 [{navWalls.Items.Count}]");
        }

        private void Load_Data_Panels_And_SubElements(bool deepcheck)
        {
            try
            {
                #region 讀取嵌板數據
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                LogicalAndFilter cwpanel_InstancesFilter =
                    new LogicalAndFilter(
                        new ElementClassFilter(typeof(FamilyInstance)),
                        new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels));
                var eles = collector
                    .WherePasses(cwpanel_InstancesFilter)
                    .ToElements()
                    .Where(x => (x as FamilyInstance).Symbol.Family.Name != "系统嵌板");

                Global.DocContent.IsCurtainPanelsLoaded = true;
                Global.DocContent.CurtainPanelList.Clear();
                Global.DocContent.ScheduleElementList.Clear();
                Global.DocContent.GeneralElementFullList.Clear();
                Global.DocContent.DeepElementList.Clear();
                int i = 0;
                foreach (var _ele in eles)
                {
                    CurtainPanelInfo _gp = new CurtainPanelInfo(_ele as RvtDB.Panel);
                    //_gp.INF_Level = ElevAtLevel(_gp.INF_OriginZ_Metric).Key + 1;
                    _gp.INF_Level = GetLevelByElevation(_gp.INF_OriginZ_Metric).INF_LevelIndex;
                    //_gp.INF_ZoneCode = $"CW-{_gp.INF_Level:00}-{_gp.INF_Type:00}-{_gp.INF_Direction}{_gp.INF_System}-Z{_gp.INF_ZoneID:00}";
                    _gp.INF_ZoneCode = $"CW-{_gp.INF_Level:00}-00-{_gp.INF_Direction}{_gp.INF_System}-Z{_gp.INF_ZoneID:00}";
                    Global.DocContent.CurtainPanelList.Add(_gp);

                    if (isRealTimeProgress)
                    {
                        progressDataPanel.Text = $"[{_gp.INF_ZoneCode}], {++i}/{eles.Count()}";
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                if (Global.DocContent.CurtainPanelList.Count > 0) Global.DocContent.IsCurtainPanelsLoaded = true;
                listInformation.SelectedIndex = listInformation.Items.Add($"讀取当前所有幕墙嵌板 [{Global.DocContent.CurtainPanelList.Count}]");

                Load_Tree_Panels();

                int __indexpanel = 0;
                foreach (var item in Global.DocContent.Lookup_CurtainPanels)
                {
                    foreach (var p in item
                        .OrderBy(p1 => Math.Round(p1.INF_OriginZ_Metric / Constants.RVTPrecision))
                        .ThenBy(p2 => Math.Round(p2.INF_OriginX_Metric / Constants.RVTPrecision)))
                    {
                        __indexpanel++;
                        p.INF_Code = $"CW-{p.INF_Level:00}-{p.INF_Type:00}-{p.INF_Direction}{p.INF_System}-{__indexpanel:0000}";//嵌板编码
                        if (deepcheck)
                        {
                            //查询嵌板中子构件
                            ICollection<ElementId> p_subs = ((RvtDB.Panel)doc.GetElement(new ElementId(p.INF_ElementId))).GetSubComponentIds();
                            p_subs.ToList<ElementId>().ForEach(id =>
                            {
                                ScheduleElementInfo _schedule_element = new ScheduleElementInfo();
                                RvtDB.Element _element = (doc.GetElement(id));
                                _schedule_element.INF_ElementId = id.IntegerValue;
                                _schedule_element.INF_Name = _element.Name;
                                _schedule_element.INF_HostCurtainPanel = p;

                                #region 判断组合构件
                                if (_schedule_element.INF_Name.Contains("挂件组")) _schedule_element.INF_Type = 71;
                                else if (_schedule_element.INF_Name.Contains("连接件组")) _schedule_element.INF_Type = 72;
                                else
                                {
                                    //非組合構件
                                    Parameter _param;
                                    if ((_param = _element.get_Parameter("分项")).HasValue)
                                    {
                                        if (int.TryParse(_param.AsString(), out int _type)) _schedule_element.INF_Type = _type;
                                        else
                                        {
                                            _schedule_element.INF_Type = -10;
                                            _schedule_element.INF_ErrorInfo = "構件[分項]參數錯誤(INF_Type)";
                                            listInformation.SelectedIndex = listInformation.Items.Add($"[{_schedule_element.INF_HostCurtainPanel.INF_ElementId}][{_schedule_element.INF_ElementId}]:[分项]参数错误...");
                                        }
                                    }
                                    else
                                    {
                                        _schedule_element.INF_Type = -10;
                                        _schedule_element.INF_ErrorInfo = "構件[分項]參數未設置(INF_Type)";
                                        listInformation.SelectedIndex = listInformation.Items.Add($"[{_schedule_element.INF_HostCurtainPanel.INF_ElementId}][{_schedule_element.INF_ElementId}]:[分项]参数错误...");
                                    }
                                }

                                if (_schedule_element.INF_Type == 71 || _schedule_element.INF_Type == 72) //处理嵌套族实例
                                {
                                    _schedule_element.INF_HasDeepElements = true;
                                    _schedule_element.INF_DeepElements = new List<DeepElementInfo>();
                                    var _deepids = ((FamilyInstance)_element).GetSubComponentIds();
                                    foreach (ElementId sid in _deepids)
                                    {
                                        DeepElementInfo _deep_element = new DeepElementInfo();
                                        RvtDB.Element _de = (doc.GetElement(sid));
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
                                                listInformation.SelectedIndex = listInformation.Items.Add($"[分項]參數錯誤");
                                                listInformation.SelectedIndex = listInformation.Items.Add($"[{_deep_element.INF_HostCurtainPanel.INF_ElementId}][{_deep_element.INF_ElementId}]:[分项]参数错误...");
                                            }
                                        }
                                        else
                                        {
                                            _deep_element.INF_Type = -10;
                                            listInformation.SelectedIndex = listInformation.Items.Add($"[分項]參數未設置");
                                            listInformation.SelectedIndex = listInformation.Items.Add($"[{_deep_element.INF_HostCurtainPanel.INF_ElementId}][{_deep_element.INF_ElementId}]:[分项]参数错误...");
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
                                #endregion
                                _schedule_element.INF_Level = p.INF_Level;
                                _schedule_element.INF_System = p.INF_System;
                                _schedule_element.INF_Direction = p.INF_Direction;

                                XYZ _xyzOrigin = ((FamilyInstance)_element).GetTotalTransform().Origin;
                                _schedule_element.INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.X);
                                _schedule_element.INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Y);
                                _schedule_element.INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Z);

                                _schedule_element.INF_ZoneID = p.INF_ZoneID;
                                _schedule_element.INF_ZoneCode = p.INF_ZoneCode;
                                _schedule_element.INF_Type_ZoneCode = $"[{_schedule_element.INF_ZoneCode}][{_schedule_element.INF_Type:00}]";
                                Global.DocContent.ScheduleElementList.Add(_schedule_element);
                            });
                        }

                        if (isRealTimeProgress)
                        {
                            progressDataPanel.Text = $"{__indexpanel}/{Global.DocContent.CurtainPanelList.Count}";
                            System.Windows.Forms.Application.DoEvents();
                        }
                    }
                }
                listInformation.SelectedIndex = listInformation.Items.Add($"讀取当前所有嵌板構件 [{Global.DocContent.ScheduleElementList.Count}]");
                if (Global.DocContent.CurtainPanelList.Count > 0) Global.DocContent.IsCurtainPanelsLoaded = true;
                if (Global.DocContent.GeneralElementFullList.Count > 0) Global.DocContent.IsDeepElementsInPanelsLoaded = true;

                if (deepcheck)
                {
                    Global.DocContent.IsDeepElementsInPanelsLoaded = true;

                    //深層构件完全清单
                    Global.DocContent.ScheduleElementList.ForEach(s =>
                    {
                        Global.DocContent.GeneralElementFullList.Add(s);
                        if (s.INF_HasDeepElements)
                        {
                            Global.DocContent.GeneralElementFullList.AddRange(s.INF_DeepElements);
                            Global.DocContent.DeepElementList.AddRange(s.INF_DeepElements);
                        }

                    });
                    if (Global.DocContent.ScheduleElementList.Count > 0) Global.DocContent.IsDeepElementsInPanelsLoaded = true;

                    Load_Tree_SubElements();
                    int __indexsubelement = 0;
                    foreach (var g in Global.DocContent.Lookup_GeneralElementsFull)
                    {
                        foreach (var ele in g
                            .OrderBy(e1 => Math.Round(e1.INF_HostCurtainPanel.INF_OriginZ_Metric / Constants.RVTPrecision))
                            .ThenBy(e2 => Math.Round(e2.INF_HostCurtainPanel.INF_OriginX_Metric / Constants.RVTPrecision))
                            .ThenBy(e3 => Math.Round(doc.GetElement(new ElementId(e3.INF_ElementId)).get_BoundingBox(doc.ActiveView).Min.Z / Constants.RVTPrecision))
                            .ThenBy(e4 => Math.Round(doc.GetElement(new ElementId(e4.INF_ElementId)).get_BoundingBox(doc.ActiveView).Min.X / Constants.RVTPrecision)))
                        {
                            __indexsubelement++;
                            ele.INF_Code = $"CW-{ele.INF_Level:00}-{ele.INF_Type:00}-{ele.INF_Direction}{ele.INF_System}-{__indexsubelement:0000}";//构件编码
                            ////ele.INF_Element.get_Parameter("分区编码").Set(ele.INF_AutoID);
                            if (isRealTimeProgress)
                            {
                                progressDataSubElement.Text = $"{__indexsubelement} / {Global.DocContent.GeneralElementFullList.Count}";
                                System.Windows.Forms.Application.DoEvents();
                            }
                        }
                    }
                    listInformation.SelectedIndex = listInformation.Items.Add($"讀取当前所有嵌板分拆構件 [{Global.DocContent.GeneralElementFullList.Count}]");
                }
                #endregion
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Source:{ex.Source}\nMessage:{ex.Message}\nTargetSite:{ex.TargetSite}\nStackTrace:{ex.StackTrace}");
            }
        }

        private void Load_Tree_Panels()
        {
            #region 生成嵌板分區
            if (Global.DocContent.IsCurtainPanelsLoaded)
            {
                navPanels.Items.Clear();
                Global.DocContent.Lookup_CurtainPanels = Global.DocContent.CurtainPanelList.OrderBy(x => x.INF_ZoneCode).ToLookup(g => g.INF_ZoneCode);
                foreach (var item in Global.DocContent.Lookup_CurtainPanels)
                {
                    ListBoxItem _navitem_zc = new ListBoxItem();
                    _navitem_zc.Content = $"分區 [{item.Key}]";
                    _navitem_zc.Name = $"ZONE_{item.Key.Replace('-', '_')}";
                    _navitem_zc.Tag = item;

                    MouseBinding mbind = new MouseBinding(cmdNavPanel, new MouseGesture(MouseAction.LeftDoubleClick));
                    mbind.CommandParameter = item;
                    mbind.CommandTarget = griddataPanel;
                    _navitem_zc.InputBindings.Add(mbind);
                    navPanels.Items.Add(_navitem_zc);
                }
            }
            ExpandNavTree("PANEL");
            listInformation.SelectedIndex = listInformation.Items.Add($"生成嵌板分區列表 [{navPanels.Items.Count}]");
            #endregion
        }


        private void Load_Tree_SubElements()
        {
            #region 生成子構件分區
            if (Global.DocContent.IsDeepElementsInPanelsLoaded)
            {
                Global.DocContent.Lookup_ScheduleElements = Global.DocContent.ScheduleElementList.OrderBy(x => x.INF_Type_ZoneCode).ToLookup(g => g.INF_Type_ZoneCode);
                Global.DocContent.Lookup_GeneralElementsFull = Global.DocContent.GeneralElementFullList.OrderBy(x => x.INF_Type_ZoneCode).ToLookup(g => g.INF_Type_ZoneCode);
                navSubElements.Items.Clear();
                foreach (var g in Global.DocContent.Lookup_GeneralElementsFull)
                {
                    ListBoxItem _navitem_se = new ListBoxItem();
                    _navitem_se.Content = $"分區 [{g.Key.ToString()}]";
                    _navitem_se.Name = $"ZONE_{System.Text.RegularExpressions.Regex.Replace(g.Key.ToString(), @"[-|\[|\]]", "_")}";
                    _navitem_se.Tag = g;

                    MouseBinding mbind = new MouseBinding(cmdNavSubElement, new MouseGesture(MouseAction.LeftDoubleClick));
                    mbind.CommandParameter = g;
                    mbind.CommandTarget = griddataSubelement;
                    _navitem_se.InputBindings.Add(mbind);
                    navSubElements.Items.Add(_navitem_se);
                }
            }
            ExpandNavTree("SE");
            listInformation.SelectedIndex = listInformation.Items.Add($"生成嵌板構件分區列表 [{navSubElements.Items.Count}]");
            #endregion
        }

        private void bnLoadElements_Click(object sender, RoutedEventArgs e)
        {
            Load_Data_Levels();
            listInformation.SelectedIndex = listInformation.Items.Add($"提取当前模型中所有标高, 共{Global.DocContent.LevelDictionary.Count}...");
            Load_Data_CurtainSystems();
            listInformation.SelectedIndex = listInformation.Items.Add($"提取当前模型中所有幕墙系统, 共{Global.DocContent.CurtainSystemList.Count}...");
            Load_Data_Walls();
            listInformation.SelectedIndex = listInformation.Items.Add($"提取当前模型中所有墻（幕墻）系统, 共{Global.DocContent.WallList.Count}...");
            Load_Data_Panels_And_SubElements(false);
            Load_Data_Panels_And_SubElements(true);

            ContentSerialize();
        }

        private void SetupProjectParameters()
        {
            #region 设置项目参数

            using (Transaction trans = new Transaction(doc, "CreateProjectParameters"))
            {
                trans.Start();
                #region 设置项目参数：立面朝向
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "立面朝向"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Walls));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtaSystem));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "立面朝向", ParameterType.Text, true, _catset, BuiltInParameterGroup.INVALID, true);
                }
                #endregion
                #region 设置项目参数：立面系统
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "立面系统"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Walls));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtaSystem));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "立面系统", ParameterType.Text, true, _catset, BuiltInParameterGroup.INVALID, true);
                }
                #endregion
                #region 设置项目参数：立面楼层
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "立面楼层"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "立面楼层", ParameterType.Integer, true, _catset, BuiltInParameterGroup.INVALID, true);
                }
                #endregion
                #region 设置项目参数：构件分项
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "构件分项"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "构件分项", ParameterType.Integer, true, _catset, BuiltInParameterGroup.INVALID, true);
                }
                #endregion
                #region 设置项目参数：构件子项
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "构件子项"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "构件子项", ParameterType.Text, true, _catset, BuiltInParameterGroup.INVALID, true);
                }
                #endregion
                #region 设置项目参数：加工图号
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "加工图号"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "加工图号", ParameterType.Text, true, _catset, BuiltInParameterGroup.INVALID, true);
                }
                #endregion

                #region 设置项目参数：分区
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "分区"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Walls));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtaSystem));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "分区", ParameterType.Integer, true, _catset, BuiltInParameterGroup.INVALID, true);
                }
                #endregion

                #region 设置项目参数：分区区号
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "分区区号"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "分区区号", ParameterType.Text, true, _catset, BuiltInParameterGroup.INVALID, true);
                }
                #endregion

                #region 设置项目参数：分区编码
                if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "分区编码"))
                {
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "分区编码", ParameterType.Text, true, _catset, BuiltInParameterGroup.INVALID, true);
                }
                #endregion
                #region 输出项目参数表用于设置校核
                //更新提取所有project parameter information
                Global.DocContent.ParameterInfoList = ParameterHelper.RawGetProjectParametersInfo(doc);
                /**
                using (StreamWriter sw = new StreamWriter(Path.Combine(Path.GetDirectoryName(doc.PathName), $"ProjectParametersInfo-{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv"), false, Encoding.UTF8))
                {
                    string title = string.Empty;
                    string rows = ParameterHelper.RawParametersInfoToCSVString(paramsInfo, ref title);
                    sw.WriteLine(title);
                    sw.Write(rows);
                }
                **/
                #endregion

                trans.Commit();

            }
            using (Transaction trans = new Transaction(doc, "InitCSWALLProjectParameters"))
            {
                FilteredElementCollector collectorcs = new FilteredElementCollector(doc);
                ElementClassFilter csFilter = new ElementClassFilter(typeof(CurtainSystem));
                collectorcs.WherePasses(csFilter);
                var _cscoll = collectorcs.ToElements();
                FilteredElementCollector collectorwall = new FilteredElementCollector(doc);
                ElementClassFilter wallFilter = new ElementClassFilter(typeof(Wall));
                collectorwall.WherePasses(wallFilter);
                var _wallcoll = collectorwall.ToElements();

                trans.Start();
                Parameter _param;

                void _proc_transform_params(Element element)
                {
                    _param = element.get_Parameter("系统");
                    if (_param != null) if (_param.HasValue) element.get_Parameter("立面系统").Set(_param.AsString());
                    _param = element.get_Parameter("方位");
                    if (_param != null) if (_param.HasValue) element.get_Parameter("立面朝向").Set(_param.AsString());
                    _param = element.get_Parameter("分区");
                    if (_param != null) if (_param.HasValue) element.get_Parameter("分区").Set(1);
                }

                foreach (var ele in _cscoll) _proc_transform_params(ele);
                foreach (var ele in _wallcoll) _proc_transform_params(ele);

                trans.Commit();

            }
            #endregion
        }

        private void ApplyParameters_CurtainPanels()
        {
            //try
            {
                using (Transaction trans = new Transaction(doc, "Apply_Parameters_CurtainPanels"))
                {
                    trans.Start();
                    int i = 0;
                    Global.DocContent.CurtainPanelList.ForEach(p =>
                    {
                        RvtDB.Element _element = doc.GetElement(new ElementId(p.INF_ElementId));
                        _element.get_Parameter("立面朝向").Set(p.INF_Direction);
                        _element.get_Parameter("立面系统").Set(p.INF_System);
                        _element.get_Parameter("立面楼层").Set(p.INF_Level);
                        _element.get_Parameter("构件分项").Set(p.INF_Type);
                        _element.get_Parameter("分区").Set(p.INF_ZoneID);
                        _element.get_Parameter("分区区号").Set(p.INF_ZoneCode);

                        if (isRealTimeProgress)
                        {
                            progressApplyPanelParameters.Text = $"{i++}/{Global.DocContent.CurtainPanelList.Count}, {i * 1.0 / Global.DocContent.CurtainPanelList.Count:P0}";
                            System.Windows.Forms.Application.DoEvents();
                        }
                    });
                    trans.Commit();
                }
                listInformation.SelectedIndex = listInformation.Items.Add($"寫入嵌板 [{Global.DocContent.CurtainPanelList.Count}] 參數 [{Global.DocContent.CurtainPanelList.Count * 6}]");
            }
            //catch (Exception ex)
            {
                //MessageBox.Show($"Source:{ex.Source}\nMessage:{ex.Message}\nTargetSite:{ex.TargetSite}\nStackTrace:{ex.StackTrace}");
            }
            //listboxOutput.SelectedIndex = listboxOutput.Items.Add($"写入[幕墙嵌板]:{Global.DocContent.CurtainPanelList.Count}，参数:{Global.DocContent.CurtainPanelList.Count * 6}...");
        }

        private void ApplyParameters_SubElementsInPanels()
        {
            using (Transaction trans = new Transaction(doc, "Apply_Parameters_SubElementsInPanels"))
            {
                trans.Start();
                int i = 0;
                foreach (var lks in Global.DocContent.Lookup_GeneralElementsFull)
                    foreach (var s in lks)
                    {
                        RvtDB.Element _selement = doc.GetElement(new ElementId(s.INF_ElementId));
                        _selement.get_Parameter("立面朝向").Set(s.INF_Direction);
                        _selement.get_Parameter("立面系统").Set(s.INF_System);
                        _selement.get_Parameter("立面楼层").Set(s.INF_Level);
                        _selement.get_Parameter("构件分项").Set(s.INF_Type);
                        //p.INF_Element.get_Parameter("构件子项").Set(p.INF_Type);
                        _selement.get_Parameter("分区").Set(s.INF_ZoneID);
                        _selement.get_Parameter("分区区号").Set(s.INF_ZoneCode);
                        _selement.get_Parameter("分区编码").Set(s.INF_Code);

                        if (isRealTimeProgress)
                        {
                            progressApplySubElementParameters.Text = $"{i++}/{Global.DocContent.GeneralElementFullList.Count}, {i * 1.0 / Global.DocContent.GeneralElementFullList.Count:P0}";
                            System.Windows.Forms.Application.DoEvents();
                        }
                    }
                trans.Commit();
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"写入嵌板构件 [{Global.DocContent.ScheduleElementList.Count}] 参数 [{Global.DocContent.ScheduleElementList.Count * 7}]");
        }

        private void checkRealTimeProgress_Checked(object sender, RoutedEventArgs e)
        {
            isRealTimeProgress = true;
        }

        private void checkRealTimeProgress_Unchecked(object sender, RoutedEventArgs e)
        {
            isRealTimeProgress = false;
        }

        private void griddataZone_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ElementInfoBase selectedhost = ((DataGrid)sender).SelectedItem as ElementInfoBase;
            Element _e = doc.GetElement(new ElementId(selectedhost.INF_ElementId));

            uidoc.Selection.Elements.Clear();
            uidoc.Selection.Elements.Add(_e);

            //var _b = _e.get_BoundingBox(doc.ActiveView);
            UIView activeuiview = uidoc.GetOpenUIViews().Single(v => v.ViewId.Equals(doc.ActiveView.Id));
            //activeuiview.ZoomAndCenterRectangle(_b.Min, _b.Max);
            //activeuiview.Zoom(0.8);
            activeuiview.ZoomToFit();

        }
    }
}
