using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.IO;
using AqlaSerializer;
using System.Globalization;
using System.ComponentModel;
using System.Collections;
using System.Collections.ObjectModel;

namespace FacadeHelper
{
    /// <summary>
    /// Interaction logic for Zone.xaml
    /// </summary>
    public partial class Zone : UserControl, INotifyPropertyChanged
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Document doc;
        private ExternalCommandData cdata;
        private List<CurtainPanelInfo> SelectedCurtainPanelList = new List<CurtainPanelInfo>();
        private ZoneInfoBase CurrentZoneInfo;

        private List<ZoneInfoBase> ResultZoneInfo = new List<ZoneInfoBase>();
        private List<CurtainPanelInfo> ResultPanelInfo = new List<CurtainPanelInfo>();
        private List<ScheduleElementInfo> ResultElementInfo = new List<ScheduleElementInfo>();

        public Window ParentWin { get; set; }

        private string _currentProjectID;
        public string CurrentProjectID { get { return _currentProjectID; } set { _currentProjectID = value; OnPropertyChanged(nameof(CurrentProjectID)); } }

        private int _currentZonePanelType = 51;
        public int CurrentZonePanelType { get { return _currentZonePanelType; } set { _currentZonePanelType = value; OnPropertyChanged(nameof(CurrentZonePanelType)); } }
        private int _currentZoneLevel = 1;
        private string _currentZoneDirection = "S";
        private string _currentZoneSystem = "A";
        private int _currentZoneIndex = 1;
        private string _currentZoneCode = "Z-00-01-SA-01";
        public int CurrentZoneLevel { get { return _currentZoneLevel; } set { _currentZoneLevel = value; OnPropertyChanged(nameof(CurrentZoneLevel)); } }
        public string CurrentZoneDirection { get { return _currentZoneDirection; } set { _currentZoneDirection = value; OnPropertyChanged(nameof(CurrentZoneDirection)); } }
        public string CurrentZoneSystem { get { return _currentZoneSystem; } set { _currentZoneSystem = value; OnPropertyChanged(nameof(CurrentZoneSystem)); } }
        public int CurrentZoneIndex { get { return _currentZoneIndex; } set { _currentZoneIndex = value; OnPropertyChanged(nameof(CurrentZoneIndex)); } }
        public string CurrentZoneCode { get { return _currentZoneCode; } set { _currentZoneCode = value; OnPropertyChanged(nameof(CurrentZoneCode)); } }

        private bool _isSearchRangeZone = true;
        private bool _isSearchRangePanel = true;
        private bool _isSearchRangeElement = true;
        private bool? _isSearchRangeAll = true;
        public bool IsSearchRangeZone { get { return _isSearchRangeZone; } set { _isSearchRangeZone = value; OnPropertyChanged(nameof(IsSearchRangeZone)); } }
        public bool IsSearchRangePanel { get { return _isSearchRangePanel; } set { _isSearchRangePanel = value; OnPropertyChanged(nameof(IsSearchRangePanel)); } }
        public bool IsSearchRangeElement { get { return _isSearchRangeElement; } set { _isSearchRangeElement = value; OnPropertyChanged(nameof(IsSearchRangeElement)); } }
        public bool? IsSearchRangeAll { get { return _isSearchRangeAll; } set { _isSearchRangeAll = value; OnPropertyChanged(nameof(IsSearchRangeAll)); } }

        private bool _isRealTimeProgress = true;
        public bool IsRealTimeProgress { get { return _isRealTimeProgress; } set { _isRealTimeProgress = value; OnPropertyChanged(nameof(IsRealTimeProgress)); } }
        private bool _isPIDInitialized = false;
        public bool IsPIDInitialized { get { return _isPIDInitialized; } set { _isPIDInitialized = value; OnPropertyChanged(nameof(IsPIDInitialized)); } }
        private bool _isElementIndexRangeInitialized = false;
        public bool IsElementIndexRangeInitialized { get { return _isElementIndexRangeInitialized; } set { _isElementIndexRangeInitialized = value; OnPropertyChanged(nameof(IsElementIndexRangeInitialized)); } }
        private bool _isZoneLayerInitialized = false;
        public bool IsZoneLayerInitialized { get { return _isZoneLayerInitialized; } set { _isZoneLayerInitialized = value; OnPropertyChanged(nameof(IsZoneLayerInitialized)); } }
        private bool _isElementClassInitialized = false;
        public bool IsElementClassInitialized { get { return _isElementClassInitialized; } set { _isElementClassInitialized = value; OnPropertyChanged(nameof(IsElementClassInitialized)); } }
        private bool _isExteriorDataLinked = false;
        public bool IsExteriorDataLinked { get { return _isExteriorDataLinked; } set { _isExteriorDataLinked = value; OnPropertyChanged(nameof(IsExteriorDataLinked)); } }
        private int _countExteriorDataLinked = 0;
        public int CountExteriorDataLinked { get { return _countExteriorDataLinked; } set { _countExteriorDataLinked = value; OnPropertyChanged(nameof(CountExteriorDataLinked)); } }

        private int _currentCultureIndex = System.Threading.Thread.CurrentThread.CurrentCulture.LCID;
        public int CurrentCultureIndex { get { return _currentCultureIndex; } set { _currentCultureIndex = value; OnPropertyChanged(nameof(CurrentCultureIndex)); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public Zone(ExternalCommandData commandData)
        {
            InitializeComponent();
            DataContext = this;
            InitializeCommand();

            cdata = commandData;
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;

            string cid = Global.GetAppConfig("CurrentProjectID");
            if (cid is null)
            {
                CurrentProjectID = "NOT INITIALIZED.";
                IsPIDInitialized = false;
            }
            else
            {
                CurrentProjectID = cid;
                IsPIDInitialized = true;
            }

            #region 检测 ElementIndexRange 是否加载
            var eirlist = ZoneHelper.FnElementIndexRangeDeserialize();
            if (eirlist is null)
            {
                Global.ElementIndexRangeList = new List<ElementIndexRange>();
                IsElementIndexRangeInitialized = false;
            }
            else
            {
                Global.ElementIndexRangeList = eirlist;
                IsElementIndexRangeInitialized = true;
            }
            #endregion

            #region 检测 ElementClass 是否加载
            var eclist = ZoneHelper.FnFilterClassDeserialize();
            if (eclist is null) IsElementClassInitialized = false;
            else
            {
                Global.ElementClassList = eclist;
                IsElementClassInitialized = true;
            }
            #endregion

            #region 检测 ZoneLayer 是否加载
            var zllist = ZoneHelper.FnZoneDataDeserialize();
            if (zllist is null) IsZoneLayerInitialized = false;
            else
            {
                Global.ZoneLayerList = zllist;
                IsZoneLayerInitialized = true;
            }
            #endregion

            //确认PID/分区表/类型表都已加载才可允许工具条
            if (!(IsPIDInitialized && IsZoneLayerInitialized && IsElementClassInitialized)) panelCommandBar.IsEnabled = false;

            Global.DataFile = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(doc.PathName), $"{System.IO.Path.GetFileNameWithoutExtension(doc.PathName)}.data");
            if (Global.DocContent is null)
            {
                Global.DocContent = new DocumentContent();
                if (File.Exists(Global.DataFile))
                {
                    using (Stream stream = new FileStream(Global.DataFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        Global.DocContent = Serializer.Deserialize<DocumentContent>(stream);
                    }
                }
            }

            Global.DocContent.ParameterInfoList = ParameterHelper.RawGetProjectParametersInfo(doc);

            navZone.ItemsSource = Global.DocContent.ZoneList;
            datagridZones.ItemsSource = Global.DocContent.ZoneList;



        }


        #region 初始化 Command

        private RoutedCommand cmdModelInit = new RoutedCommand();
        private RoutedCommand cmdT24Fix = new RoutedCommand();
        private RoutedCommand cmdOnElementClassify = new RoutedCommand();
        private RoutedCommand cmdElementClassify = new RoutedCommand();
        private RoutedCommand cmdElementLink = new RoutedCommand();
        private RoutedCommand cmdElementUnLink = new RoutedCommand();
        private RoutedCommand cmdElementResolve = new RoutedCommand();
        private RoutedCommand cmdElementExplode = new RoutedCommand();
        private RoutedCommand cmdElementDeepResolve = new RoutedCommand();

        private RoutedCommand cmdNavZone = new RoutedCommand();

        private RoutedCommand cmdLoadData = new RoutedCommand();
        private RoutedCommand cmdSaveData = new RoutedCommand();
        private RoutedCommand cmdResetData = new RoutedCommand();
        private RoutedCommand cmdApplyParameters = new RoutedCommand();
        private RoutedCommand cmdExportElementSchedule = new RoutedCommand();

        private RoutedCommand cmdSearch = new RoutedCommand();

        private RoutedCommand cmdPopupClose = new RoutedCommand();


        private void InitializeCommand()
        {
            CommandBinding cbModelInit = new CommandBinding(cmdModelInit, cbModelInit_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbT24Fix = new CommandBinding(cmdT24Fix, (sender, e) =>
            {
                using (StreamWriter file = new StreamWriter(@"d:\11.csv", true))
                    Global.DocContent.ScheduleElementList.ForEach(se => file.WriteLine($"{se.INF_ElementId},{se.INF_Name},{se.INF_Type},{se.INF_TaskLayer},{se.INF_TaskSubLayer},{se.INF_Code}"));
                using (StreamWriter file = new StreamWriter(@"d:\12.csv", true))
                    Global.DocContent.FullScheduleElementList.ForEach(se => file.WriteLine($"{se.INF_ElementId},{se.INF_Name},{se.INF_Type},{se.INF_TaskLayer},{se.INF_TaskSubLayer},{se.INF_Code}"));


                Global.DocContent.CurtainPanelList.ForEach(p =>
                {
                    Autodesk.Revit.DB.Panel _cp = (doc.GetElement(new ElementId(p.INF_ElementId))) as Autodesk.Revit.DB.Panel;
                    var p_subs = _cp.GetSubComponentIds();
                    foreach (ElementId eid in p_subs)
                    {
                        ScheduleElementInfo _sei = new ScheduleElementInfo();
                        Element __element = (doc.GetElement(eid));
                        _sei.INF_ElementId = eid.IntegerValue;
                        _sei.INF_Name = __element.Name;
                        _sei.INF_ZoneCode = p.INF_ZoneCode;
                        _sei.INF_ZoneIndex = p.INF_ZoneIndex;
                        _sei.INF_HostZoneInfo = p.INF_HostZoneInfo;
                        _sei.INF_Level = p.INF_Level;
                        _sei.INF_Direction = p.INF_Direction;
                        _sei.INF_System = p.INF_System;
                        _sei.INF_HostCurtainPanel = p;

                        XYZ _xyzOrigin = ((FamilyInstance)__element).GetTotalTransform().Origin;
                        _sei.INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.X);
                        _sei.INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Y);
                        _sei.INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Z);

                        #region 确定分项参数 + 工序层级
                        Parameter _parameter;
                        _parameter = __element.get_Parameter("分项");
                        if (int.Parse(_parameter.AsString()) == 24)
                        {
                            _sei.INF_Type = 24;
                            _sei.INF_TaskLayer = 1;
                            _sei.INF_TaskSubLayer = 26;
                            _sei.INF_IsScheduled = true;
                            Global.DocContent.ScheduleElementList.Add(_sei);
                            listInformation.SelectedIndex = listInformation.Items.Add($"{_sei.INF_ElementId}, {_sei.INF_Name}, {_sei.INF_Type}, {_sei.INF_TaskLayer}");
                        }
                        #endregion
                    }
                });
                listInformation.SelectedIndex = listInformation.Items.Add($"{Global.DocContent.ScheduleElementList.Where(se => se.INF_Type == 24).Count()}");
                Global.DocContent.FullScheduleElementList.Clear();
                Global.DocContent.FullScheduleElementList.AddRange(Global.DocContent.ScheduleElementList);
                using (StreamWriter file = new StreamWriter(@"d:\21.csv", true))
                    Global.DocContent.ScheduleElementList.ForEach(se => file.WriteLine($"{se.INF_ElementId},{se.INF_Name},{se.INF_Type},{se.INF_TaskLayer},{se.INF_TaskSubLayer},{se.INF_Code}"));
                using (StreamWriter file = new StreamWriter(@"d:\22.csv", true))
                    Global.DocContent.FullScheduleElementList.ForEach(se => file.WriteLine($"{se.INF_ElementId},{se.INF_Name},{se.INF_Type},{se.INF_TaskLayer},{se.INF_TaskSubLayer},{se.INF_Code}"));

            },
            (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbOnElementClassify = new CommandBinding(cmdOnElementClassify, (sender, e) =>
            {
                if (bnOnElementClassify.IsChecked == true) return;
                if (Global.DocContent.CurrentZoneInfo.ZoneIndex == 0)
                {
                    CurrentZoneLevel = Global.DocContent.CurrentZoneInfo.ZoneLevel = 1;
                    CurrentZoneDirection = Global.DocContent.CurrentZoneInfo.ZoneDirection = "S";
                    CurrentZoneSystem = Global.DocContent.CurrentZoneInfo.ZoneSystem = "A";
                    CurrentZoneIndex = Global.DocContent.CurrentZoneInfo.ZoneIndex = 1;
                    CurrentZoneCode = Global.DocContent.CurrentZoneInfo.ZoneCode = "Z-00-01-SA-01";
                }
                else
                {
                    CurrentZoneLevel = Global.DocContent.CurrentZoneInfo.ZoneLevel;
                    CurrentZoneDirection = Global.DocContent.CurrentZoneInfo.ZoneDirection;
                    CurrentZoneSystem = Global.DocContent.CurrentZoneInfo.ZoneSystem;
                    CurrentZoneIndex = Global.DocContent.CurrentZoneInfo.ZoneIndex;
                    CurrentZoneCode = Global.DocContent.CurrentZoneInfo.ZoneCode;
                }
            },
            (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbElementClassify = new CommandBinding(cmdElementClassify, cbElementClassify_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbElementLink = new CommandBinding(cmdElementLink,
                (sender, e) =>
                {
                    Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog()
                    {
                        Multiselect = true,
                        InitialDirectory = System.IO.Path.GetDirectoryName(Global.DataFile),
                        DefaultExt = "*.elist",
                        Filter = "Element Collection Files(*.elist)|*.elist|All(*.*)|*.*"
                    };
                    if (ofd.ShowDialog() == true)
                    {
                        ZoneHelper.FnLinkedElementsDeserialize(ofd.FileNames);
                        bnElementResolve.SetResourceReference(ContentControl.TagProperty, "IconMassSort");
                        bnElementLink.Foreground = new SolidColorBrush(Colors.DarkGreen);
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - LINK: FILE, {string.Join(",", ofd.SafeFileNames)}.");
                        IsExteriorDataLinked = true;
                    }
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbElementUnLink = new CommandBinding(cmdElementUnLink,
                (sender, e) =>
                {
                    Global.DocContent.ExternalElementDataList.Clear();
                    Global.DocContent.FullCurtainPanelList.Clear();
                    Global.DocContent.FullScheduleElementList.Clear();
                    Global.DocContent.FullZoneList.Clear();
                    IsExteriorDataLinked = false;
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbElementResolve = new CommandBinding(cmdElementResolve,
                (sender, e) =>
                {
                    //Global.DocContent.ScheduleElementList.Clear();
                    if (Global.DocContent.FullCurtainPanelList.Count == 0) Global.DocContent.FullCurtainPanelList.AddRange(Global.DocContent.CurtainPanelList);
                    if (Global.DocContent.FullScheduleElementList.Count == 0) Global.DocContent.FullScheduleElementList.AddRange(Global.DocContent.ScheduleElementList);
                    if (Global.DocContent.FullZoneList.Count == 0) foreach (var z in Global.DocContent.ZoneList) Global.DocContent.FullZoneList.Add(z);

                    progbarGlobalProcess.Maximum = Global.DocContent.FullZoneList.Count;
                    progbarGlobalProcess.Value = 0;

                    foreach (var zn in Global.DocContent.FullZoneList.GroupBy(z => z.ZoneCode))
                    {
                        ZoneHelper.FnResolveZone(uidoc, zn.ElementAt(0),
                            ref txtCurrentProcessElement,
                            ref txtCurrentProcessOperation,
                            ref progbarCurrentProcess,
                            ref listInformation);
                        progbarGlobalProcess.Value++;
                    }
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbElementExplode = new CommandBinding(cmdElementExplode,
                (sender, e) =>
                {
                    if (MessageBox.Show(
                        "组合构件分解后的深层子构件将进入总构件清单参与排序和编码，该过程不可逆，是否继续？",
                        "组合构件分解",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question,
                        MessageBoxResult.No) == MessageBoxResult.No) return;
                    int[] assemlabels = new int[] { 21, 22, 24 };
                    var _ele_asm = Global.DocContent.ScheduleElementList.Where(ele => assemlabels.Contains(ele.INF_Type)).ToList();
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - LIST: E/{_ele_asm.Count}/DEEP.");
                    bool _haserror = false;

                    _ele_asm.ForEach(ae =>
                    {
                        var _ae = doc.GetElement(new ElementId(ae.INF_ElementId));
                        var _deepids = ((FamilyInstance)_ae).GetSubComponentIds();
                        if (_deepids != null && _deepids.Count > 0)
                        {
                            ae.INF_HasDeepElements = true;
                            foreach (ElementId sid in _deepids)
                            {
                                txtCurrentProcessElement.Content = $"E/{ae.INF_ElementId}";
                                txtCurrentProcessOperation.Content = "RESOLVE-DEEP/E";
                                System.Windows.Forms.Application.DoEvents();

                                Element __element = (doc.GetElement(sid));
                                XYZ _xyzOrigin = ((FamilyInstance)__element).GetTotalTransform().Origin;
                                ScheduleElementInfo _sei = new ScheduleElementInfo()
                                {
                                    INF_ElementId = sid.IntegerValue,
                                    INF_Name = __element.Name,
                                    INF_ZoneCode = ae.INF_ZoneCode,
                                    INF_ZoneIndex = ae.INF_ZoneIndex,
                                    INF_HostZoneInfo = ae.INF_HostZoneInfo,
                                    INF_Level = ae.INF_Level,
                                    INF_Direction = ae.INF_Direction,
                                    INF_System = ae.INF_System,
                                    INF_HostCurtainPanel = ae.INF_HostCurtainPanel,
                                    INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.X),
                                    INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Y),
                                    INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Z)
                                };

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
                                                _sei.INF_TaskLayer = __sec.ETaskLayer;
                                                _sei.INF_TaskSubLayer = __sec.ETaskSubLayer;
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
                                            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: VALUE TYPE, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, E/{_sei.INF_ElementId}, {_sei.INF_Name}.");
                                            uidoc.Selection.Elements.Add(__element);
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        _sei.INF_Type = -2;
                                        _haserror = true;
                                        _sei.INF_ErrorInfo = "构件[分项]参数无参数值(INF_Type)";
                                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: NO VALUE TYPE, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, E/{_sei.INF_ElementId}, {_sei.INF_Name}.");
                                        uidoc.Selection.Elements.Add(__element);
                                        continue;
                                    }
                                }
                                else
                                {
                                    _sei.INF_Type = -3;
                                    _haserror = true;
                                    _sei.INF_ErrorInfo = "构件[分项]参数项未设置(INF_Type)";
                                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: PARAM NOTSET, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, {_sei.INF_ElementId}.");
                                    uidoc.Selection.Elements.Add(__element);
                                    continue;
                                }
                                #endregion

                                DeepElementInfo _dei = new DeepElementInfo()
                                {
                                    INF_ElementId = sid.IntegerValue,
                                    INF_Name = __element.Name,
                                    INF_ZoneCode = ae.INF_ZoneCode,
                                    INF_ZoneIndex = ae.INF_ZoneIndex,
                                    INF_HostZoneInfo = ae.INF_HostZoneInfo,
                                    INF_Level = ae.INF_Level,
                                    INF_Direction = ae.INF_Direction,
                                    INF_System = ae.INF_System,
                                    INF_HostCurtainPanel = ae.INF_HostCurtainPanel,
                                    INF_Type = _sei.INF_Type,
                                    INF_TaskLayer = _sei.INF_TaskLayer,
                                    INF_TaskSubLayer = _sei.INF_TaskSubLayer,
                                    INF_IsScheduled = _sei.INF_IsScheduled,
                                    INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.X),
                                    INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Y),
                                    INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Z),
                                    INF_HostScheduleElement = ae
                                };

                                Global.DocContent.ScheduleElementList.Add(_sei);
                                ae.INF_DeepElements.Add(_dei);
                            }
                        }
                    });
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbElementDeepResolve = new CommandBinding(cmdElementDeepResolve,
                (sender, e) =>
                {
                    Global.DocContent.DeepElementList.Clear();
                    int[] assemlabels = new int[] { 21, 22, 24 };
                    var _ele_asm = Global.DocContent.ScheduleElementList.Where(ele => assemlabels.Contains(ele.INF_Type)).ToList();
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - LIST: E/{_ele_asm.Count}/DEEP.");
                    bool _haserror = false;

                    _ele_asm.ForEach(ae =>
                    {
                        var _ae = doc.GetElement(new ElementId(ae.INF_ElementId));
                        var _deepids = ((FamilyInstance)_ae).GetSubComponentIds();
                        if (_deepids != null && _deepids.Count > 0)
                        {
                            ae.INF_HasDeepElements = true;
                            foreach (ElementId sid in _deepids)
                            {
                                txtCurrentProcessElement.Content = $"E/{ae.INF_ElementId}";
                                txtCurrentProcessOperation.Content = "RESOLVE-DEEP/E";
                                System.Windows.Forms.Application.DoEvents();

                                Element __element = (doc.GetElement(sid));
                                XYZ _xyzOrigin = ((FamilyInstance)__element).GetTotalTransform().Origin;
                                DeepElementInfo _sei = new DeepElementInfo()
                                {
                                    INF_ElementId = sid.IntegerValue,
                                    INF_Name = __element.Name,
                                    INF_ZoneCode = ae.INF_ZoneCode,
                                    INF_ZoneIndex = ae.INF_ZoneIndex,
                                    INF_HostZoneInfo = ae.INF_HostZoneInfo,
                                    INF_Level = ae.INF_Level,
                                    INF_Direction = ae.INF_Direction,
                                    INF_System = ae.INF_System,
                                    INF_HostCurtainPanel = ae.INF_HostCurtainPanel,
                                    INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.X),
                                    INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Y),
                                    INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Z),
                                    INF_HostScheduleElement = ae
                                };

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
                                                _sei.INF_TaskLayer = __sec.ETaskLayer;
                                                _sei.INF_TaskSubLayer = __sec.ETaskSubLayer;
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
                                            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: VALUE TYPE, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, E/{_sei.INF_ElementId}, {_sei.INF_Name}.");
                                            uidoc.Selection.Elements.Add(__element);
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        _sei.INF_Type = -2;
                                        _haserror = true;
                                        _sei.INF_ErrorInfo = "构件[分项]参数无参数值(INF_Type)";
                                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: NO VALUE TYPE, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, E/{_sei.INF_ElementId}, {_sei.INF_Name}.");
                                        uidoc.Selection.Elements.Add(__element);
                                        continue;
                                    }
                                }
                                else
                                {
                                    _sei.INF_Type = -3;
                                    _haserror = true;
                                    _sei.INF_ErrorInfo = "构件[分项]参数项未设置(INF_Type)";
                                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: PARAM NOTSET, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, {_sei.INF_ElementId}.");
                                    uidoc.Selection.Elements.Add(__element);
                                    continue;
                                }
                                #endregion

                                Global.DocContent.DeepElementList.Add(_sei);
                                ae.INF_DeepElements.Add(_sei);
                            }
                        }
                    });

                    #region 確定嵌板內明細構件數據
                    var groupsDeepElements = Global.DocContent.DeepElementList.ToLookup(se => se.INF_HostScheduleElement.INF_Code);
                    int eidx = 0;
                    foreach (var grp in groupsDeepElements)
                        foreach (var ele in grp)
                        {
                            ele.INF_Index = ++eidx;
                            ele.INF_Code = $"CW-{ele.INF_Type:00}-{ele.INF_Level:00}-{ele.INF_Direction}{ele.INF_System}{ele.INF_ZoneIndex:0}-X{eidx:000}";//X构件编码
                            txtCurrentProcessElement.Content = $"E/{ele.INF_HostScheduleElement.INF_Code}.DEEP-E/{ele.INF_ElementId}, {ele.INF_Code}";
                            txtCurrentProcessOperation.Content = "W/TASK";
                            System.Windows.Forms.Application.DoEvents();

                        }
                    #endregion

                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });


            CommandBinding cbNavZone = new CommandBinding(cmdNavZone,
                (sender, e) =>
                {
                    datagridPanels.ItemsSource = null;
                    datagridPanels.ItemsSource = e.Parameter as List<CurtainPanelInfo>;
                    tabPanel.Content = $"嵌板 / {(e.Parameter as List<CurtainPanelInfo>).Count}, Z/{((navZone.SelectedItem as ListBoxItem).Tag as ZoneInfoBase).ZoneCode}";
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbLoadData = new CommandBinding(cmdLoadData,
                (sender, e) =>
                {
                    Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog()
                    {
                        InitialDirectory = System.IO.Path.GetDirectoryName(Global.DataFile),
                        DefaultExt = "*.data",
                        Filter = "Data Files(*.data)|*.data|Data Backup Files(*.bak)|*.bak|All(*.*)|*.*"
                    };
                    ZoneHelper.FnContentDeserialize(ofd.FileName);
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - LOAD: FILE, {ofd.FileName}.");
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbSaveData = new CommandBinding(cmdSaveData,
                (sender, e) =>
                {
                    ZoneHelper.FnContentSerializeWithBackup();
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - WRITE: FILE, {Global.DataFile}.bak.");
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - UPDATE: FILE, {Global.DataFile}.");
                    ZoneHelper.FnLinkedElementsSerialize(ref listInformation);
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbResetData = new CommandBinding(cmdResetData,
                (sender, e) =>
                {
                    if (MessageBox.Show($"重置当前数据后将不能恢复，确定继续操作？",
                        "重置数据...",
                        MessageBoxButton.OKCancel,
                        MessageBoxImage.Exclamation,
                        MessageBoxResult.OK) == MessageBoxResult.OK)
                        Global.DocContent = new DocumentContent();
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - RESET: ALL DATA.");
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbSearch = new CommandBinding(cmdSearch,
                (sender, e) =>
                {
                    #region 數據檢索
                    if (!IsSearchRangeZone && !IsSearchRangePanel && !IsSearchRangeElement) return;
                    ResultZoneInfo.Clear();
                    ResultPanelInfo.Clear();
                    ResultElementInfo.Clear();

                    ZoneHelper.FnSearch(uidoc,
                        txtSearchKeyword.Text.Trim(),
                        ref ResultZoneInfo, ref ResultPanelInfo, ref ResultElementInfo,
                        IsSearchRangeZone, IsSearchRangePanel, IsSearchRangeElement,
                        ref listInformation);
                    if (IsSearchRangeZone)
                    {
                        txtResultZone.Content = $"Z/{ResultZoneInfo.Count}; ";
                        datagridZones.ItemsSource = null;
                        datagridZones.ItemsSource = ResultZoneInfo;
                        tabZone.Content = $"分区 / {ResultZoneInfo.Count}";
                    }
                    else
                    {
                        txtResultZone.Content = $"Z/SKIP; ";
                        datagridZones.ItemsSource = null;
                        tabZone.Content = $"分区 / SKIP";
                    }
                    if (IsSearchRangePanel)
                    {
                        txtResultPanel.Content = $"P/{ResultPanelInfo.Count}; ";
                        datagridPanels.ItemsSource = null;
                        datagridPanels.ItemsSource = ResultPanelInfo;
                        tabPanel.Content = $"嵌板 / {ResultPanelInfo.Count}";
                    }
                    else
                    {
                        txtResultPanel.Content = $"P/SKIP; ";
                        datagridPanels.ItemsSource = null;
                        tabPanel.Content = $"嵌板 / SKIP";
                    }
                    if (IsSearchRangeElement)
                    {
                        txtResultElement.Content = $"E/{ResultElementInfo.Count}; ";
                        datagridScheduleElements.ItemsSource = null;
                        datagridScheduleElements.ItemsSource = ResultElementInfo;
                        tabElement.Content = $"构件 / {ResultElementInfo.Count}";
                    }
                    else
                    {
                        txtResultElement.Content = $"E/SKIP; ";
                        datagridScheduleElements.ItemsSource = null;
                        tabElement.Content = $"构件 / SKIP";
                    }

                    #endregion
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbApplyParameters = new CommandBinding(cmdApplyParameters,
                (sender, e) =>
                {
                    if (MessageBox.Show($"是否采用当前数据写入模型参数？\n如不采用将读取外部数据文件(ExternalElementData序列化数据)覆盖当前数据。",
                        "参数数据源...",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question,
                        MessageBoxResult.Yes) == MessageBoxResult.No)
                    {
                        Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog()
                        {
                            InitialDirectory = System.IO.Path.GetDirectoryName(Global.DataFile),
                            DefaultExt = "*.sort",
                            Filter = "Sorted Elements Collection Files(*.sort)|*.sort|Element Collection Files(*.elist)|*.elist|All(*.*)|*.*"
                        };
                        if (ofd.ShowDialog() == true)
                        {
                            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - READ: DATA, {ofd.FileName}");
                            using (Stream stream = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.Read))
                            {
                                var extdata = Serializer.Deserialize<ExternalElementData>(stream);
                                extdata.ExternalId = 0;
                                extdata.ExternalFileName = ofd.FileName;
                                Global.DocContent.CurtainPanelList.Clear();
                                Global.DocContent.ScheduleElementList.Clear();
                                Global.DocContent.CurtainPanelList.AddRange(extdata.CurtainPanelList);
                                Global.DocContent.ScheduleElementList.AddRange(extdata.ScheduleElementList);
                                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - UPDATE: LIST/P");
                                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - UPDATE: LIST/E");
                            }
                        }
                    }

                    #region 參數寫入
                    //try
                    {
                        using (Transaction trans = new Transaction(doc, "Apply_Parameters_CurtainPanels"))
                        {
                            trans.Start();
                            progbarGlobalProcess.Maximum = 100;
                            progbarGlobalProcess.Value = 0;
                            progbarCurrentProcess.Maximum = Global.DocContent.CurtainPanelList.Count;
                            progbarCurrentProcess.Value = 0;

                            Global.DocContent.CurtainPanelList.ForEach(p =>
                            {
                                Element _element = doc.GetElement(new ElementId(p.INF_ElementId));
                                if (_element != null)
                                {
                                    _element.get_Parameter("立面朝向").Set(p.INF_Direction);
                                    _element.get_Parameter("立面系统").Set(p.INF_System);
                                    _element.get_Parameter("立面楼层").Set(p.INF_Level);
                                    _element.get_Parameter("构件分项").Set(p.INF_Type);
                                    _element.get_Parameter("分区序号").Set(p.INF_ZoneIndex);
                                    _element.get_Parameter("分区区号").Set(p.INF_ZoneCode);
                                    _element.get_Parameter("分区编码").Set(p.INF_Code);
                                }

                                if (IsRealTimeProgress)
                                {
                                    txtCurrentProcessElement.Content = $"Z/{p.INF_ZoneCode}.P/{p.INF_ElementId}, {p.INF_Code}";
                                    txtCurrentProcessOperation.Content = "W/PARAM";
                                    progbarCurrentProcess.Value++;
                                    progbarGlobalProcess.Value += 35d / Global.DocContent.CurtainPanelList.Count;
                                    txtGlobalProcessElement.Content = $"Z/{p.INF_ZoneCode}";
                                    txtGlobalProcessOperation.Content = $"W/PARAM/P";
                                    System.Windows.Forms.Application.DoEvents();
                                }
                            });
                            trans.Commit();
                        }
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - WRITE: PARAM, P/{Global.DocContent.CurtainPanelList.Count}, Param#{Global.DocContent.CurtainPanelList.Count * 7}");

                        using (Transaction trans = new Transaction(doc, "Apply_Parameters_ScheduleElements"))
                        {
                            trans.Start();
                            //progbarElement.Maximum = Global.DocContent.ScheduleElementList.Count;
                            //progbarElement.Value = 0;
                            progbarGlobalProcess.Value = 35;
                            progbarCurrentProcess.Maximum = Global.DocContent.ScheduleElementList.Count;
                            progbarCurrentProcess.Value = 0;

                            Global.DocContent.ScheduleElementList.ForEach(ele =>
                            {
                                Element _element = doc.GetElement(new ElementId(ele.INF_ElementId));
                                if (_element != null)
                                {
                                    _element.get_Parameter("立面朝向").Set(ele.INF_Direction);
                                    _element.get_Parameter("立面系统").Set(ele.INF_System);
                                    _element.get_Parameter("立面楼层").Set(ele.INF_Level);
                                    _element.get_Parameter("构件分项").Set(ele.INF_Type);
                                    _element.get_Parameter("分区序号").Set(ele.INF_ZoneIndex);
                                    _element.get_Parameter("分区区号").Set(ele.INF_ZoneCode);
                                    _element.get_Parameter("分区编码").Set(ele.INF_Code);
                                }

                                if (IsRealTimeProgress)
                                {
                                    txtCurrentProcessElement.Content = $"P/{ele.INF_HostCurtainPanel.INF_Code}.E/{ele.INF_ElementId}, {ele.INF_Code}";
                                    txtCurrentProcessOperation.Content = "W/PARAM";
                                    progbarCurrentProcess.Value++;
                                    progbarGlobalProcess.Value += 35d / Global.DocContent.ScheduleElementList.Count;
                                    txtGlobalProcessElement.Content = $"P/{ele.INF_ZoneCode}]";
                                    txtGlobalProcessOperation.Content = $"W/PARAM/E";
                                    System.Windows.Forms.Application.DoEvents();
                                }
                            });
                            trans.Commit();
                        }

                        using (Transaction trans = new Transaction(doc, "Apply_Parameters_DeepElements"))
                        {
                            trans.Start();
                            //progbarElement.Maximum = Global.DocContent.ScheduleElementList.Count;
                            //progbarElement.Value = 0;
                            progbarGlobalProcess.Value = 70;
                            progbarCurrentProcess.Maximum = Global.DocContent.DeepElementList.Count;
                            progbarCurrentProcess.Value = 0;

                            #region 设置项目参数：组件编码
                            if (!Global.DocContent.ParameterInfoList.Exists(x => x.Name == "组件编码"))
                            {
                                CategorySet _catset = new CategorySet();
                                _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                                ParameterHelper.RawCreateProjectParameter(doc.Application, "组件编码", ParameterType.Text, true, _catset, BuiltInParameterGroup.PG_DATA, true);
                            }
                            #endregion

                            Global.DocContent.DeepElementList.ForEach(ele =>
                            {
                                Element _element = doc.GetElement(new ElementId(ele.INF_ElementId));
                                _element.get_Parameter("立面朝向").Set(ele.INF_Direction);
                                _element.get_Parameter("立面系统").Set(ele.INF_System);
                                _element.get_Parameter("立面楼层").Set(ele.INF_Level);
                                _element.get_Parameter("构件分项").Set(ele.INF_Type);
                                _element.get_Parameter("分区序号").Set(ele.INF_ZoneIndex);
                                _element.get_Parameter("分区区号").Set(ele.INF_ZoneCode);
                                _element.get_Parameter("分区编码").Set(ele.INF_Code);
                                _element.get_Parameter("组件编码").Set(ele.INF_HostScheduleElement.INF_Code);

                                if (IsRealTimeProgress)
                                {
                                    txtCurrentProcessElement.Content = $"E/{ele.INF_HostScheduleElement.INF_Code}.DEEP-E/{ele.INF_ElementId}, {ele.INF_Code}";
                                    txtCurrentProcessOperation.Content = "W/PARAM";
                                    progbarCurrentProcess.Value++;
                                    progbarGlobalProcess.Value += 30d / Global.DocContent.DeepElementList.Count;
                                    txtGlobalProcessElement.Content = $"P/{ele.INF_ZoneCode}]";
                                    txtGlobalProcessOperation.Content = $"W/PARAM/E";
                                    System.Windows.Forms.Application.DoEvents();
                                }
                            });
                            trans.Commit();
                        }

                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - WRITE: PARAM, E/{Global.DocContent.ScheduleElementList.Count}, Param#{Global.DocContent.ScheduleElementList.Count * 7}");
                    }
                    //catch (Exception ex)
                    {
                        //MessageBox.Show($"Source:{ex.Source}\nMessage:{ex.Message}\nTargetSite:{ex.TargetSite}\nStackTrace:{ex.StackTrace}");
                    }
                    //listboxOutput.SelectedIndex = listboxOutput.Items.Add($"写入[幕墙嵌板]:{Global.DocContent.CurtainPanelList.Count}，参数:{Global.DocContent.CurtainPanelList.Count * 6}...");

                    #endregion
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbExportElementSchedule = new CommandBinding(cmdExportElementSchedule,
                (sender, e) =>
                {
                    using (StreamWriter writer = new StreamWriter(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(doc.PathName), $"{System.IO.Path.GetFileNameWithoutExtension(doc.PathName)}.4dzone.csv"), false))
                    {
                        int idtask = 0;
                        writer.WriteLine($"Id,Name,Start,Finish,Outline Level");
                        foreach (var zone in Global.DocContent.ZoneList)
                        {
                            //var _plist = Global.DocContent.CurtainPanelList.Where(p => p.INF_ZoneCode == zone.ZoneCode);
                            var _elist = Global.DocContent.MullionList.Where(m => m.INF_ZoneCode == zone.ZoneCode);
                            writer.WriteLine($"{++idtask},{zone.ZoneCode},,,1");

                            //foreach (var p in _plist) writer.WriteLine($"{++idtask},{p.INF_Code},{p.INF_TaskStart},{p.INF_TaskFinish},2");
                            foreach (var m in _elist) writer.WriteLine($"{++idtask},{m.INF_Code},{m.INF_TaskStart},{m.INF_TaskFinish},2");
                        }
                    }
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbPopupClose = new CommandBinding(cmdPopupClose, (sender, e) => { bnQuickStart.IsChecked = false; }, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            bnModelInit.Command = cmdModelInit;
            bnT24Fix.Command = cmdT24Fix;
            bnOnElementClassify.Command = cmdOnElementClassify;
            bnElementClassify.Command = cmdElementClassify;
            bnElementExplode.Command = cmdElementExplode;
            bnElementLink.Command = cmdElementLink;
            bnElementUnLink.Command = cmdElementUnLink;
            bnElementResolve.Command = cmdElementResolve;
            bnElementDeepResolve.Command = cmdElementDeepResolve;
            bnLoadData.Command = cmdLoadData;
            bnSaveData.Command = cmdSaveData;
            bnResetData.Command = cmdResetData;
            bnApplyParameters.Command = cmdApplyParameters;
            bnExportElementSchedule.Command = cmdExportElementSchedule;
            bnSearch.Command = cmdSearch;
            bnPopupClose.Command = cmdPopupClose;

            ProcZone.CommandBindings.AddRange(new CommandBinding[]
            {
                cbModelInit,
                cbT24Fix,
                cbOnElementClassify,
                cbElementClassify,
                cbElementExplode,
                cbElementLink,
                cbElementUnLink,
                cbElementResolve,
                cbElementDeepResolve,
                cbNavZone,
                cbLoadData,
                cbSaveData,
                cbResetData,
                cbApplyParameters,
                cbSearch,
                cbPopupClose
            });
        }

        private void navZone_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var item = ItemsControl.ContainerFromElement(sender as ListBox, e.OriginalSource as DependencyObject) as ListBoxItem;
            if (item != null)
            {
                var zi = item.Content as ZoneInfoBase;
                datagridPanels.ItemsSource = null;
                var _plist = Global.DocContent.FullCurtainPanelList.Where(p => p.INF_ZoneCode == zi.ZoneCode);
                datagridPanels.ItemsSource = _plist;
                tabPanel.Content = $"嵌板 / {_plist.Count()}, Z/{zi.ZoneCode}]";
                tabPanel.IsChecked = true;
                datagridScheduleElements.ItemsSource = null;
                var _elist = Global.DocContent.FullScheduleElementList.Where(ele => ele.INF_ZoneCode == zi.ZoneCode);
                datagridScheduleElements.ItemsSource = _elist;
                tabElement.Content = $"构件 / {_elist.Count()}, Z/{zi.ZoneCode}";
            }

        }

        #region Command -- bnElementClassify : 嵌板和构件归类
        private void cbElementClassify_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            bnOnElementClassify.IsChecked = false;
            popZC.IsOpen = false;

            CurrentZoneInfo = new ZoneInfoBase()
            {
                ZoneLevel = CurrentZoneLevel,
                ZoneDirection = CurrentZoneDirection,
                ZoneSystem = CurrentZoneSystem,
                ZoneIndex = CurrentZoneIndex,
                ZoneCode = CurrentZoneCode = (txtZoneCode.Content as TextBlock).Text
            };
            Global.DocContent.CurrentZoneInfo = CurrentZoneInfo;

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: NONE SELECTED, E/ALL.");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - SELECT: ELE/{ids.Count}.");
            //FilteredElementCollector panelcollector = new FilteredElementCollector(doc, ids);
            LogicalAndFilter cwpanel_InstancesFilter =
                new LogicalAndFilter(
                    new ElementClassFilter(typeof(FamilyInstance)),
                    new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels));
            LogicalAndFilter win_InstancesFilter =
                new LogicalAndFilter(
                    new ElementClassFilter(typeof(FamilyInstance)),
                    new ElementCategoryFilter(BuiltInCategory.OST_Windows));
            var panels = new FilteredElementCollector(doc, ids).WherePasses(cwpanel_InstancesFilter)
                .Where(x =>
                (x as FamilyInstance).Symbol.Family.Name != "系统嵌板" &&
                (x as FamilyInstance).Symbol.Family.Name != "空嵌板" &&
                (x as FamilyInstance).Symbol.Family.Name != "空系统嵌板" &&
                (x as FamilyInstance).Symbol.Name != "NULL")
                .ToList();
            var wins = new FilteredElementCollector(doc, ids).WherePasses(win_InstancesFilter).ToList();

            if (panels.Count == 0 && wins.Count == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: NONE SELECTED, P&W/VALID.");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - SELECT: P/{panels.Count()}, W/{wins.Count()}.");
            SelectedCurtainPanelList.Clear();
            var pw = panels.Union(wins).ToList();

            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - LIST: P/{panels.Count()}/E, W/{wins.Count()}/E.");
            bool _haserror = false;
            pw.ForEach(_ele =>
            {
                bool ispanel = ((BuiltInCategory)_ele.Category.Id.IntegerValue == BuiltInCategory.OST_Windows) ? false : true;
                CurtainPanelInfo _gp = new CurtainPanelInfo(_ele, Global.DocContent.CurrentZoneInfo.ZoneCode, ispanel)
                {
                    INF_Type = CurrentZonePanelType,
                    INF_HostZoneInfo = Global.DocContent.CurrentZoneInfo
                };
                SelectedCurtainPanelList.Add(_gp);

                //確定嵌板內明細構件數據
                progbarCurrentProcess.Maximum = panels.Count();
                progbarCurrentProcess.Value = 0;

                var p_subs = ((FamilyInstance)_ele).GetSubComponentIds();
                foreach (ElementId eid in p_subs)
                {
                    txtCurrentProcessElement.Content = $"P/{_gp.INF_ElementId}/E{eid}";
                    txtCurrentProcessOperation.Content = "RESOLVE/P.E";
                    progbarCurrentProcess.Value++;
                    System.Windows.Forms.Application.DoEvents();

                    ScheduleElementInfo _sei = new ScheduleElementInfo();
                    Element __element = (doc.GetElement(eid));
                    _sei.INF_ElementId = eid.IntegerValue;
                    _sei.INF_Name = __element.Name;
                    _sei.INF_ZoneCode = Global.DocContent.CurrentZoneInfo.ZoneCode;
                    _sei.INF_ZoneIndex = Global.DocContent.CurrentZoneInfo.ZoneIndex;
                    _sei.INF_HostZoneInfo = Global.DocContent.CurrentZoneInfo;
                    _sei.INF_Level = Global.DocContent.CurrentZoneInfo.ZoneLevel;
                    _sei.INF_Direction = Global.DocContent.CurrentZoneInfo.ZoneDirection;
                    _sei.INF_System = Global.DocContent.CurrentZoneInfo.ZoneSystem;
                    _sei.INF_HostCurtainPanel = _gp;

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
                                    _sei.INF_TaskLayer = __sec.ETaskLayer;
                                    _sei.INF_TaskSubLayer = __sec.ETaskSubLayer;
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
                                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: VALUE TYPE, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, E/{_sei.INF_ElementId}, {_sei.INF_Name}.");
                                continue;
                            }
                        }
                        else
                        {
                            _sei.INF_Type = -2;
                            _haserror = true;
                            _sei.INF_ErrorInfo = "构件[分项]参数无参数值(INF_Type)";
                            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: NO VALUE TYPE, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, E/{_sei.INF_ElementId}, {_sei.INF_Name}.");
                            continue;
                        }
                    }
                    else
                    {
                        _sei.INF_Type = -3;
                        _haserror = true;
                        _sei.INF_ErrorInfo = "构件[分项]参数项未设置(INF_Type)";
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: PARAM NOTSET, PARAM/TYPE, P/{_sei.INF_HostCurtainPanel.INF_ElementId}, {_sei.INF_ElementId}.");
                        continue;
                    }
                    #endregion

                    XYZ _xyzOrigin = ((FamilyInstance)__element).GetTotalTransform().Origin;
                    _sei.INF_OriginX_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.X);
                    _sei.INF_OriginY_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Y);
                    _sei.INF_OriginZ_Metric = Unit.CovertFromAPI(DisplayUnitType.DUT_MILLIMETERS, _xyzOrigin.Z);

                    Global.DocContent.ScheduleElementList.Add(_sei);
                }

            });

            #region 构件参数错误中断操作确认
            if (_haserror)
            {
                if (MessageBox.Show($"有部分构件存在参数错误，是否继续处理数据？", "构件参数错误...", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.No) return;
            }
            #endregion

            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - WRITE: PARAM, ZONECODE, Z/{Global.DocContent.CurrentZoneInfo.ZoneCode}, P/{panels.Count()}.");

            datagridPanels.ItemsSource = null;
            datagridPanels.ItemsSource = SelectedCurtainPanelList;
            tabPanel.Content = $"嵌板 / {SelectedCurtainPanelList.Count}";

            using (Transaction trans = new Transaction(doc, "CreatePanelGroup"))
            {
                trans.Start();
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                ICollection<Element> typecollection = collector.OfClass(typeof(SelectionFilterElement)).ToElements();
                SelectionFilterElement selectset = typecollection.Cast<SelectionFilterElement>().FirstOrDefault(ele => ele.Name == CurrentZoneInfo.ZoneCode);
                if (selectset != null) selectset.Clear();
                else selectset = SelectionFilterElement.Create(doc, CurrentZoneInfo.ZoneCode);
                panels.ForEach(p => selectset.AddSingle(p.Id));
                doc.ActiveView.HideElementsTemporary(selectset.GetElementIds());
                trans.Commit();
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - WRITE: SECECTSET, SSET/{CurrentZoneInfo.FilterName}, P/{selectset.GetElementIds().Count}");
            }

            Global.DocContent.ZoneList.Add(CurrentZoneInfo);
            Global.DocContent.CurtainPanelList.AddRange(SelectedCurtainPanelList);

            ZoneHelper.FnContentSerialize();
        }

        #endregion

        #region Command -- bnModelInit : 模型初始化
        private void cbModelInit_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            InitModelData();
        }
        #endregion

        #region Function 初始化模型数据
        private void InitModelData()
        {
            ParameterHelper.InitProjectParameters(ref doc);

            //统计 
            LogicalAndFilter cwpanel_InstancesFilter =
                new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels));
            LogicalAndFilter win_InstancesFilter =
                new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_Windows));
            var panels = new FilteredElementCollector(doc).WherePasses(cwpanel_InstancesFilter)
                .Where(x => (x as FamilyInstance).Symbol.Family.Name != "系统嵌板" && (x as FamilyInstance).Symbol.Name != "空嵌板" && (x as FamilyInstance).Symbol.Name != "空系统嵌板");
            var wins = new FilteredElementCollector(doc).WherePasses(win_InstancesFilter).ToElements();
            var syspanels = new FilteredElementCollector(doc).WherePasses(cwpanel_InstancesFilter).Where(x => (x as FamilyInstance).Symbol.Family.Name == "系统嵌板");

            int nele = 0;
            int nwele = 0;
            int nsysele = 0;
            foreach (var _p in panels) nele += (_p as FamilyInstance).GetSubComponentIds().Count;
            foreach (var _w in wins) nwele += (_w as FamilyInstance).GetSubComponentIds().Count;
            foreach (var _p in syspanels) nsysele += (_p as FamilyInstance).GetSubComponentIds().Count;
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - CALC: TOTAL, P/{panels.Count():N0}, E/{nele:N0}, SYS-P/{syspanels.Count():N0}, SYS-E/{nsysele:N0}");
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - CALC: TOTAL, W/{wins.Count():N0}, E/{nwele:N0}");
        }

        #endregion        
        #endregion

        private void listInformation_SelectionChanged(object sender, SelectionChangedEventArgs e) { var lb = sender as ListBox; lb.ScrollIntoView(lb.Items[lb.Items.Count - 1]); }

        private void chkbox_Checked(object sender, RoutedEventArgs e)
        {
            if (((CheckBox)sender).Name == "chkSearchRangeAll")
            {
                switch (chkSearchRangeAll.IsChecked)
                {
                    case true:
                        IsSearchRangeZone = true;
                        IsSearchRangePanel = true;
                        IsSearchRangeElement = true;
                        break;
                    case false:
                        IsSearchRangeZone = false;
                        IsSearchRangePanel = false;
                        IsSearchRangeElement = false;
                        break;
                    default:
                        break;
                }
            }
            else
            {
                if (IsSearchRangeZone && IsSearchRangePanel && IsSearchRangeElement) IsSearchRangeAll = true;
                else if (!IsSearchRangeZone && !IsSearchRangePanel && !IsSearchRangeElement) IsSearchRangeAll = false;
                else IsSearchRangeAll = null;
            }
            ((CheckBox)sender).GetBindingExpression(CheckBox.IsCheckedProperty).UpdateTarget();
        }

        private void datagridElements_GotFocus(object sender, RoutedEventArgs e)
        {
            ScheduleElementInfo currentsei = new ScheduleElementInfo();
            if (datagridScheduleElements.CurrentItem != null) currentsei = (ScheduleElementInfo)(datagridScheduleElements.CurrentItem);

        }

    }

    public class IntToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (int)value == int.Parse(parameter.ToString());
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var data = (bool)value;
            if (data)
            {
                return System.Convert.ToInt32(parameter);
            }
            return -1;
        }
    }

    public class VarToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value.Equals(parameter);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value.Equals(true) ? parameter : System.Windows.Data.Binding.DoNothing;
        }
    }


    public class PrefixConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) return false;
            return (value as string).StartsWith(parameter as string);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value.Equals(true) ? parameter : System.Windows.Data.Binding.DoNothing;
        }
    }

    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool bValue = (bool)value;
            if (bValue) return System.Windows.Visibility.Visible;
            else return System.Windows.Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            System.Windows.Visibility visibility = (System.Windows.Visibility)value;

            if (visibility == System.Windows.Visibility.Visible) return true;
            else return false;
        }
    }

    public class NegateBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return !(bool)value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return !(bool)value;
        }
    }

    public class MultiValueElementClassConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool bsch = (bool)values[0]; //IsScheduled, bool
            int ilayer = (int)values[1]; //ETaskLayer, int
            int ilc = int.Parse(parameter as string);
            if (bsch && ilayer == ilc)
            {
                return System.Windows.Visibility.Visible;
            }
            return System.Windows.Visibility.Collapsed;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class NotNullValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            if (string.IsNullOrEmpty(value as string) || string.IsNullOrWhiteSpace(value as string))
            {
                return new ValidationResult(false, "不能为空！");
            }
            return new ValidationResult(true, null);
        }
    }
    public class ProjectIDRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            string _id = value as string;

            if (!string.IsNullOrWhiteSpace(_id))
            {
                string pIDFormartRegex = @"^BIM\d{10}[C|W]$";

                // 检查输入的字符串是否符合IP地址格式
                if (!System.Text.RegularExpressions.Regex.IsMatch(_id, pIDFormartRegex))
                {
                    return new ValidationResult(false, "项目编号格式不正确");
                }
            }
            return new ValidationResult(true, null);
        }
    }
}
