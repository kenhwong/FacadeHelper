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

        public Window parentWin { get; set; }

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
        private RoutedCommand cmdOnElementClassify = new RoutedCommand();
        private RoutedCommand cmdElementClassify = new RoutedCommand();
        private RoutedCommand cmdElementResolve = new RoutedCommand();

        private RoutedCommand cmdNavZone = new RoutedCommand();

        private RoutedCommand cmdLoadData = new RoutedCommand();
        private RoutedCommand cmdSaveData = new RoutedCommand();
        private RoutedCommand cmdApplyParameters = new RoutedCommand();
        private RoutedCommand cmdExportElementSchedule = new RoutedCommand();

        private RoutedCommand cmdSearch = new RoutedCommand();

        private RoutedCommand cmdPopupClose = new RoutedCommand();

        private void InitializeCommand()
        {
            CommandBinding cbModelInit = new CommandBinding(cmdModelInit, cbModelInit_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbOnElementClassify = new CommandBinding(cmdOnElementClassify, (sender, e) =>
            {
                expZoneCodeInput.IsExpanded = true;
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
            CommandBinding cbElementResolve = new CommandBinding(cmdElementResolve,
                (sender, e) =>
                {
                    Global.DocContent.ScheduleElementList.Clear();
                    foreach (var zn in Global.DocContent.ZoneList) ZoneHelper.FnResolveZone(uidoc, zn, 
                        ref txtCurrentProcessElement,
                        ref txtCurrentProcessOperation,
                        ref progbarCurrentProcess,
                        ref txtGlobalProcessElement,
                        ref txtGlobalProcessOperation,
                        ref progbarGlobalProcess,
                        ref listInformation, ref txtProcessInfo);
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbNavZone = new CommandBinding(cmdNavZone,
                (sender, e) =>
                {
                    datagridPanels.ItemsSource = null;
                    datagridPanels.ItemsSource = e.Parameter as List<CurtainPanelInfo>;
                    expDataGridPanels.Header = $"分区[{((navZone.SelectedItem as ListBoxItem).Tag as ZoneInfoBase).ZoneCode}]幕墙嵌板数据列表";
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbLoadData = new CommandBinding(cmdLoadData,
                (sender, e) =>
                {
                    Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
                    ofd.InitialDirectory = System.IO.Path.GetDirectoryName(Global.DataFile);
                    ofd.DefaultExt = "*.data";
                    ofd.Filter = "Data Files(*.data)|*.data|Data Backup Files(*.bak)|*.bak|All(*.*)|*.*";
                    if (ofd.ShowDialog() == true)
                        if (MessageBox.Show($"确认加载新的数据文件 {ofd.FileName}？\n\n现有数据将被新的数据覆盖，且不可恢复，但不会影响模型文件。选择确认继续，取消则不会有任何操作。", "加载新的数据文件...",
                            MessageBoxButton.OKCancel,
                            MessageBoxImage.Question,
                            MessageBoxResult.OK) == MessageBoxResult.OK)
                        {
                            ZoneHelper.FnContentDeserialize(ofd.FileName);
                            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - LOAD: FILE, {ofd.FileName}.");
                        }
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbSaveData = new CommandBinding(cmdSaveData,
                (sender, e) =>
                {
                    if (MessageBox.Show($"确认更新数据的更改？\n\n更新的数据将保存至 {Global.DataFile}，现有数据将创建备份，不会影响模型文件。选择确认继续，取消则不会有任何操作。", "更新数据修改...",
                        MessageBoxButton.OKCancel,
                        MessageBoxImage.Question,
                        MessageBoxResult.OK) == MessageBoxResult.OK)
                    {
                        ZoneHelper.FnContentSerializeWithBackup();
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - WRITE: FILE, {Global.DataFile}.bak.");
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - UPDATE: FILE, {Global.DataFile}.");
                    }
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
                    txtProcessInfo.Content = $"檢索結果：";
                    if (IsSearchRangeZone)
                    {
                        txtResultZone.Content = $"[分區： {ResultZoneInfo.Count}]";
                        datagridZones.ItemsSource = null;
                        datagridZones.ItemsSource = ResultZoneInfo;
                        expDataGridZones.Header = $"分區檢索結果： {ResultZoneInfo.Count}";
                    }
                    else
                    {
                        txtResultZone.Content = $"[分區： 不檢索]";
                        datagridZones.ItemsSource = null;
                        expDataGridZones.Header = $"分區檢索結果： 不檢索";
                    }
                    if (IsSearchRangePanel)
                    {
                        txtResultPanel.Content = $"[幕墻嵌板： {ResultPanelInfo.Count}]";
                        datagridPanels.ItemsSource = null;
                        datagridPanels.ItemsSource = ResultPanelInfo;
                        expDataGridPanels.Header = $"幕墻嵌板檢索結果： {ResultPanelInfo.Count}";
                    }
                    else
                    {
                        txtResultPanel.Content = $"[幕墻嵌板： 不檢索]";
                        datagridPanels.ItemsSource = null;
                        expDataGridPanels.Header = $"幕墻嵌板檢索結果： 不檢索";
                    }
                    if (IsSearchRangeElement)
                    {
                        txtResultElement.Content = $"[明細構件： {ResultElementInfo.Count}]";
                        datagridScheduleElements.ItemsSource = null;
                        datagridScheduleElements.ItemsSource = ResultElementInfo;
                        expDataGridScheduleElements.Header = $"分區檢索結果： {ResultElementInfo.Count}";
                    }
                    else
                    {
                        txtResultElement.Content = $"[明細構件： 不檢索]";
                        datagridScheduleElements.ItemsSource = null;
                        expDataGridScheduleElements.Header = $"分區檢索結果： 不檢索";
                    }

                    #endregion
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbApplyParameters = new CommandBinding(cmdApplyParameters,
                (sender, e) =>
                {
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
                                _element.get_Parameter("立面朝向").Set(p.INF_Direction);
                                _element.get_Parameter("立面系统").Set(p.INF_System);
                                _element.get_Parameter("立面楼层").Set(p.INF_Level);
                                _element.get_Parameter("构件分项").Set(p.INF_Type);
                                _element.get_Parameter("分区序号").Set(p.INF_ZoneIndex);
                                _element.get_Parameter("分区区号").Set(p.INF_ZoneCode);
                                _element.get_Parameter("分区编码").Set(p.INF_Code);

                                if (IsRealTimeProgress)
                                {
                                    txtCurrentProcessElement.Content = $"[分区-{p.INF_ZoneCode}].[嵌板-{p.INF_ElementId}, {p.INF_Code}]";
                                    txtCurrentProcessOperation.Content = "参数写入";
                                    progbarCurrentProcess.Value++;
                                    progbarGlobalProcess.Value += 50d / Global.DocContent.CurtainPanelList.Count;
                                    txtGlobalProcessElement.Content = $"[分区-{p.INF_ZoneCode}]";
                                    txtGlobalProcessOperation.Content = $"嵌板";
                                    System.Windows.Forms.Application.DoEvents();
                                }
                            });
                            trans.Commit();
                        }
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - WRITE: PARAM, P/{Global.DocContent.CurtainPanelList.Count}, Param#{Global.DocContent.CurtainPanelList.Count * 7}");

                        using (Transaction trans = new Transaction(doc, "Apply_Parameters_ScheduleElements"))
                        {
                            trans.Start();
                            //progbarElement.Maximum = Global.DocContent.ScheduleElementList.Count;
                            //progbarElement.Value = 0;
                            progbarGlobalProcess.Value = 50;
                            progbarCurrentProcess.Maximum = Global.DocContent.ScheduleElementList.Count;
                            progbarCurrentProcess.Value = 0;

                            Global.DocContent.ScheduleElementList.ForEach(ele =>
                            {
                                Element _element = doc.GetElement(new ElementId(ele.INF_ElementId));
                                _element.get_Parameter("立面朝向").Set(ele.INF_Direction);
                                _element.get_Parameter("立面系统").Set(ele.INF_System);
                                _element.get_Parameter("立面楼层").Set(ele.INF_Level);
                                _element.get_Parameter("构件分项").Set(ele.INF_Type);
                                _element.get_Parameter("分区序号").Set(ele.INF_ZoneIndex);
                                _element.get_Parameter("分区区号").Set(ele.INF_ZoneCode);
                                _element.get_Parameter("分区编码").Set(ele.INF_Code);

                                if (IsRealTimeProgress)
                                {
                                    progbarElement.Value++;

                                    txtCurrentProcessElement.Content = $"[嵌板-{ele.INF_HostCurtainPanel.INF_Code}].[构件-{ele.INF_ElementId}, {ele.INF_Code}]";
                                    txtCurrentProcessOperation.Content = "参数写入";
                                    progbarCurrentProcess.Value++;
                                    progbarGlobalProcess.Value += 50d / Global.DocContent.CurtainPanelList.Count;
                                    txtGlobalProcessElement.Content = $"[分区-{ele.INF_ZoneCode}]";
                                    txtGlobalProcessOperation.Content = $"构件";
                                    System.Windows.Forms.Application.DoEvents();
                                }
                            });
                            trans.Commit();
                        }
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - WRITE: PARAM, E/{Global.DocContent.ScheduleElementList.Count}, Param#{Global.DocContent.ScheduleElementList.Count * 7}");
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
            bnOnElementClassify.Command = cmdOnElementClassify;
            bnElementClassify.Command = cmdElementClassify;
            bnElementResolve.Command = cmdElementResolve;
            bnLoadData.Command = cmdLoadData;
            bnSaveData.Command = cmdSaveData;
            bnApplyParameters.Command = cmdApplyParameters;
            bnExportElementSchedule.Command = cmdExportElementSchedule;
            bnSearch.Command = cmdSearch;
            bnPopupClose.Command = cmdPopupClose;

            ProcZone.CommandBindings.AddRange(new CommandBinding[]
            {
                cbModelInit,
                cbOnElementClassify,
                cbElementClassify,
                cbElementResolve,
                cbNavZone,
                cbLoadData,
                cbSaveData,
                cbApplyParameters,
                cbSearch,
                cbPopupClose
            });
        }

        private void navZone_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var _addeditems = e.AddedItems;
            if (_addeditems.Count > 0)
            {
                var zi = _addeditems[0] as ZoneInfoBase;
                datagridPanels.ItemsSource = null;
                var _plist = Global.DocContent.CurtainPanelList.Where(p => p.INF_ZoneCode == zi.ZoneCode);
                datagridPanels.ItemsSource = _plist;
                expDataGridPanels.Header = $"分区[{zi.ZoneCode}]幕墙嵌板数据列表，数量 {_plist.Count()}";
                expDataGridPanels.IsExpanded = true;

                navPanels.ItemsSource = _plist;
                expPanel.Header = $"幕墙嵌板：{_plist.Count()}，分区[{zi.ZoneCode}]";
                expPanel.IsExpanded = true;
                datagridScheduleElements.ItemsSource = null;
                var _elist = Global.DocContent.ScheduleElementList.Where(ele => ele.INF_ZoneCode == zi.ZoneCode);
                datagridScheduleElements.ItemsSource = _elist;
                expDataGridScheduleElements.Header = $"分区[{zi.ZoneCode}]明细构件数据列表，数量 {_elist.Count()}";
                navZone.SelectedItem = null;
            }
        }

        private void navPanels_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var _addeditem = e.AddedItems;
            if (_addeditem.Count > 0)
            {
                var pi = _addeditem[0] as CurtainPanelInfo;
                datagridScheduleElements.ItemsSource = null;
                var _elist = Global.DocContent.ScheduleElementList.Where(ele => ele.INF_HostCurtainPanel.INF_ElementId == pi.INF_ElementId);
                datagridScheduleElements.ItemsSource = _elist;
                expDataGridScheduleElements.Header = $"分区[{pi.INF_ZoneCode}]，嵌板[{pi.INF_Code}]明细构件数据列表，数量 {_elist.Count()}";
                expDataGridScheduleElements.IsExpanded = true;
                navPanels.SelectedItem = null;
            }
        }
        

        #region Command -- bnElementClassify : 嵌板和构件归类
        private void cbElementClassify_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            //parentWin.Hide();
            //var _p_sel_list = uidoc.Selection.PickObjects(ObjectType.Element, "選擇同一分區的所有嵌板");
            //parentWin.Visibility = System.Windows.Visibility.Visible;
            Global.DocContent.CurrentZoneInfo.ZoneLevel = CurrentZoneLevel;
            Global.DocContent.CurrentZoneInfo.ZoneDirection = CurrentZoneDirection;
            Global.DocContent.CurrentZoneInfo.ZoneSystem = CurrentZoneSystem;
            Global.DocContent.CurrentZoneInfo.ZoneIndex = CurrentZoneIndex;
            Global.DocContent.CurrentZoneInfo.ZoneCode = CurrentZoneCode = (txtZoneCode.Content as TextBlock).Text;
            CurrentZoneInfo = Global.DocContent.CurrentZoneInfo;

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - ERR: NONE SELECTED, E/ALL.");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - SELECT: ELE/{ids.Count}.");
            FilteredElementCollector panelcollector = new FilteredElementCollector(doc, ids);
            LogicalAndFilter cwpanel_InstancesFilter =
                new LogicalAndFilter(
                    new ElementClassFilter(typeof(FamilyInstance)),
                    new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels));
            var panels = panelcollector
                .WherePasses(cwpanel_InstancesFilter)
                .Where(x => (x as FamilyInstance).Symbol.Family.Name != "系統嵌板" && (x as FamilyInstance).Symbol.Name != "空嵌板");

            if (panels.Count() == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - ERR: NONE SELECTED, P/VALID.");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - SELECT: P/{panels.Count()}.");
            SelectedCurtainPanelList.Clear();

            using (Transaction trans = new Transaction(doc, "Apply_Panels_ZoneCode"))
            {
                trans.Start();
                foreach (var _ele in panels)
                {
                    CurtainPanelInfo _gp = new CurtainPanelInfo(_ele as Autodesk.Revit.DB.Panel);
                    _gp.INF_Type = CurrentZonePanelType;
                    SelectedCurtainPanelList.Add(_gp);
                    Parameter _param = _ele.get_Parameter("分区区号");
                    if (_param == null)
                    {
                        MessageBox.Show($"当前选择的幕墙嵌板[{_gp.INF_ElementId}]嵌板未设置分区区号参数，请修改后继续。", "错误 - 幕墙嵌板 - 分区区号", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        return;
                    }
                    _param.Set(Global.DocContent.CurrentZoneInfo.ZoneCode);
                }
                trans.Commit();
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - WRITE: PARAM, ZONECODE, Z/{Global.DocContent.CurrentZoneInfo.ZoneCode}, P/{panels.Count()}.");

            datagridPanels.ItemsSource = null;
            datagridPanels.ItemsSource = SelectedCurtainPanelList;
            expDataGridPanels.Header = "选择区域幕墙嵌板数据列表";

            uidoc.Selection.Elements.Clear();
            foreach (var _ele in panels) uidoc.Selection.Elements.Add(_ele);

            using (Transaction trans = new Transaction(doc, "CreatePanelGroup"))
            {
                trans.Start();
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                ICollection<Element> typecollection = collector.OfClass(typeof(SelectionFilterElement)).ToElements();
                SelectionFilterElement selectset = typecollection.Cast<SelectionFilterElement>().FirstOrDefault(ele => ele.Name == CurrentZoneInfo.ZoneCode);
                if (selectset != null) selectset.Clear();
                else selectset = SelectionFilterElement.Create(doc, CurrentZoneInfo.ZoneCode);
                selectset.AddSet(uidoc.Selection.GetElementIds());
                doc.ActiveView.HideElementsTemporary(selectset.GetElementIds());
                trans.Commit();
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - WRITE: SECECTSET, SSET/{CurrentZoneInfo.FilterName}, P/{selectset.GetElementIds().Count}");
            }

            MouseBinding mbind = new MouseBinding(cmdNavZone, new MouseGesture(MouseAction.LeftDoubleClick));
            mbind.CommandParameter = SelectedCurtainPanelList;
            mbind.CommandTarget = datagridPanels;

            Global.DocContent.ZoneList.Add(CurrentZoneInfo);
            Global.DocContent.CurtainPanelList.AddRange(SelectedCurtainPanelList);

            ZoneHelper.FnContentSerialize();
            expZone.IsExpanded = true;
            subbuttongroup_SelectPanels.Visibility = System.Windows.Visibility.Collapsed;
        }

        #endregion
        private void expDataGrid_Expanded(object sender, RoutedEventArgs e)
        {
            if (((Expander)sender).Name == "expDataGridZones") { expDataGridPanels.IsExpanded = false; expDataGridScheduleElements.IsExpanded = false; }
            if (((Expander)sender).Name == "expDataGridPanels") { expDataGridZones.IsExpanded = false; expDataGridScheduleElements.IsExpanded = false; }
            if (((Expander)sender).Name == "expDataGridScheduleElements") { expDataGridZones.IsExpanded = false; expDataGridPanels.IsExpanded = false; }
        }

        private void cbModelInit_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            InitModelData();
        }

        private void InitModelData()
        {
            ParameterHelper.InitProjectParameters(ref doc);

            //加載分區進度數據
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
            ofd.InitialDirectory = System.IO.Path.GetDirectoryName(Global.DataFile);
            ofd.DefaultExt = "*.txt";
            ofd.Filter = "Schedule Text Files(*.txt)|*.txt|All(*.*)|*.*";
            if (ofd.ShowDialog() == true)
                if (MessageBox.Show($"加载新的分区进度定义文件 {ofd.FileName}？\n\n现有的进度数据将被覆盖。选择确认继续。", "加载新的进度数据文件...",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Question,
                    MessageBoxResult.OK) == MessageBoxResult.OK)
                {
                    ZoneHelper.FnLoadZoneScheduleData(ofd.FileName);
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - L: {ofd.FileName}.");
                }

        }
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

}
