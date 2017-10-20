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
    public partial class SimpleZone : UserControl, INotifyPropertyChanged
    {
        private UIApplication uiapp;
        private UIDocument uidoc;
        private Document doc;
        private ExternalCommandData cdata;
        private List<CurtainPanelInfo> SelectedCurtainPanelList = new List<CurtainPanelInfo>();
        private List<MullionInfo> SelectedMullionList = new List<MullionInfo>();
        private ZoneInfoBase CurrentZoneInfo;

        private List<ZoneInfoBase> ResultZoneInfo = new List<ZoneInfoBase>();
        private List<CurtainPanelInfo> ResultPanelInfo = new List<CurtainPanelInfo>();
        private List<MullionInfo> ResultMullionInfo = new List<MullionInfo>();

        public Window parentWin { get; set; }

        private int _currentZonePanelType = 52;
        public int CurrentZonePanelType { get { return _currentZonePanelType; } set { _currentZonePanelType = value; OnPropertyChanged(nameof(CurrentZonePanelType)); } }

        private bool _isSearchRangeZone = true;
        private bool _isSearchRangePanel = true;
        private bool _isSearchRangeMullion = true;
        private bool? _isSearchRangeAll = true;
        public bool IsSearchRangeZone { get { return _isSearchRangeZone; } set { _isSearchRangeZone = value; OnPropertyChanged(nameof(IsSearchRangeZone)); } }
        public bool IsSearchRangePanel { get { return _isSearchRangePanel; } set { _isSearchRangePanel = value; OnPropertyChanged(nameof(IsSearchRangePanel)); } }
        public bool IsSearchRangeMullion { get { return _isSearchRangeMullion; } set { _isSearchRangeMullion = value; OnPropertyChanged(nameof(IsSearchRangeMullion)); } }
        public bool? IsSearchRangeAll { get { return _isSearchRangeAll; } set { _isSearchRangeAll = value; OnPropertyChanged(nameof(IsSearchRangeAll)); } }

        private bool _isRealTimeProgress = true;
        public bool IsRealTimeProgress { get { return _isRealTimeProgress; } set { _isRealTimeProgress = value; OnPropertyChanged(nameof(IsRealTimeProgress)); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public SimpleZone(ExternalCommandData commandData)
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

            if (Global.DocContent.CurtainPanelList.Count > 0) Global.DocContent.Lookup_CurtainPanels = Global.DocContent.CurtainPanelList.ToLookup(g => g.INF_ZoneCode);
            Global.DocContent.ParameterInfoList = ParameterHelper.RawGetProjectParametersInfo(doc);

            navZone.ItemsSource = Global.DocContent.ZoneList;
            datagridZones.ItemsSource = Global.DocContent.ZoneList;
        }


        #region 初始化 Command

        private RoutedCommand cmdModelInit = new RoutedCommand();
        private RoutedCommand cmdSelectPanels = new RoutedCommand();
        private RoutedCommand cmdSelectMullions = new RoutedCommand();
        private RoutedCommand cmdZonedPnM = new RoutedCommand();
        private RoutedCommand cmdPanelResolve = new RoutedCommand();

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
            CommandBinding cbSelectPanels = new CommandBinding(cmdSelectPanels, cbSelectPanels_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbSelectMullions = new CommandBinding(cmdSelectMullions, cbSelectMullions_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbZonedPnM = new CommandBinding(cmdZonedPnM, cbZonedPnM_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbPanelResolve = new CommandBinding(cmdPanelResolve,
                (sender, e) =>
                {
                    foreach (var zn in Global.DocContent.ZoneList) ZoneHelper.FnResolveSimpleZone(uidoc, zn, ref listInformation, ref txtProcessInfo);
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });


            CommandBinding cbNavZone = new CommandBinding(cmdNavZone,
                (sender, e) =>
                {
                    datagridPanels.ItemsSource = null;
                    datagridPanels.ItemsSource = e.Parameter as List<CurtainPanelInfo>;
                    expDataGridPanels.Header = $"分區[{((navZone.SelectedItem as ListBoxItem).Tag as ZoneInfoBase).ZoneCode}]幕墻嵌板數據列表";
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
                        if (MessageBox.Show($"確認加載新數據文件 {ofd.FileName}？\n\n現有数据将被新數據覆蓋，且不可恢復，但不会影响模型文件。單擊確認繼續，取消則不會有任何操作。", "加載新數據...",
                            MessageBoxButton.OKCancel,
                            MessageBoxImage.Question,
                            MessageBoxResult.OK) == MessageBoxResult.OK)
                        {
                            ZoneHelper.FnContentDeserialize(ofd.FileName);
                            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 加載新數據文件{ofd.FileName}.");
                        }
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbSaveData = new CommandBinding(cmdSaveData,
                (sender, e) =>
                {
                    if (MessageBox.Show($"確認更新改動的數據？\n\n更新的數據將保存至 {Global.DataFile}，現有數據將創建備份，不会影响模型文件。單擊確認繼續，取消則不會有任何操作。", "更新數據修改...",
                        MessageBoxButton.OKCancel,
                        MessageBoxImage.Question,
                        MessageBoxResult.OK) == MessageBoxResult.OK)
                    {
                        ZoneHelper.FnContentSerializeWithBackup();
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 備份數據文件 {Global.DataFile}.bak.");
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 更新數據文件 {Global.DataFile}.");
                    }
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbSearch = new CommandBinding(cmdSearch,
                (sender, e) =>
                {
                    #region 數據檢索
                    if (!IsSearchRangeZone && !IsSearchRangePanel && !IsSearchRangeMullion) return;
                    ResultZoneInfo.Clear();
                    ResultPanelInfo.Clear();
                    ResultMullionInfo.Clear();
                    ZoneHelper.FnSearch(uidoc,
                        txtSearchKeyword.Text.Trim(),
                        ref ResultZoneInfo, ref ResultPanelInfo, ref ResultMullionInfo,
                        IsSearchRangeZone, IsSearchRangePanel, IsSearchRangeMullion,
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
                    if (IsSearchRangeMullion)
                    {
                        txtResultElement.Content = $"[幕墻竪梃： {ResultMullionInfo.Count}]";
                        datagridScheduleElements.ItemsSource = null;
                        datagridScheduleElements.ItemsSource = ResultMullionInfo;
                        expDataGridScheduleElements.Header = $"分區檢索結果： {ResultMullionInfo.Count}";
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
                            progbarPanel.Maximum = Global.DocContent.CurtainPanelList.Count;
                            progbarPanel.Value = 0;
                            Global.DocContent.CurtainPanelList.ForEach(p =>
                            {
                                Element _element = doc.GetElement(new ElementId(p.INF_ElementId));
                                _element.get_Parameter("立面朝向").Set(p.INF_Direction);
                                _element.get_Parameter("立面系统").Set(p.INF_System);
                                _element.get_Parameter("立面楼层").Set(p.INF_Level);
                                _element.get_Parameter("构件分项").Set(p.INF_Type);
                                _element.get_Parameter("分区序号").Set(p.INF_ZoneID);
                                _element.get_Parameter("分区区号").Set(p.INF_ZoneCode);
                                _element.get_Parameter("分区编码").Set(p.INF_Code);
                                _element.get_Parameter("安装开始").Set(p.INF_TaskStart.ToString("MM-dd-yy HH:mm:ss"));
                                _element.get_Parameter("安装结束").Set(p.INF_TaskFinish.ToString("MM-dd-yy HH:mm:ss"));

                                if (IsRealTimeProgress)
                                {
                                    txtProcessInfo.Content = $"當前處理進度：[分區：{p.INF_ZoneCode}] - [幕墻嵌板：{p.INF_ElementId}, {p.INF_Code})]";
                                    progbarPanel.Value++;
                                    System.Windows.Forms.Application.DoEvents();
                                }
                            });
                            trans.Commit();
                        }
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 寫入：嵌板 [{Global.DocContent.CurtainPanelList.Count}] 參數 [{Global.DocContent.CurtainPanelList.Count * 9}]");

                        using (Transaction trans = new Transaction(doc, "Apply_Parameters_Mullions"))
                        {
                            trans.Start();
                            progbarMullion.Maximum = Global.DocContent.MullionList.Count;
                            progbarMullion.Value = 0;
                            Global.DocContent.MullionList.ForEach(mi =>
                            {
                                Element _element = doc.GetElement(new ElementId(mi.INF_ElementId));
                                _element.get_Parameter("立面朝向").Set(mi.INF_Direction);
                                _element.get_Parameter("立面系统").Set(mi.INF_System);
                                _element.get_Parameter("立面楼层").Set(mi.INF_Level);
                                _element.get_Parameter("构件分项").Set(mi.INF_Type);
                                _element.get_Parameter("分区序号").Set(mi.INF_ZoneID);
                                _element.get_Parameter("分区区号").Set(mi.INF_ZoneCode);
                                _element.get_Parameter("分区编码").Set(mi.INF_Code);
                                _element.get_Parameter("安装开始").Set(mi.INF_TaskStart.ToString("MM-dd-yy HH:mm:ss"));
                                _element.get_Parameter("安装结束").Set(mi.INF_TaskFinish.ToString("MM-dd-yy HH:mm:ss"));

                                if (IsRealTimeProgress)
                                {
                                    txtProcessInfo.Content = $"當前處理進度：[分區：{mi.INF_ZoneCode}] - [幕墻竪梃：{mi.INF_ElementId}, {mi.INF_Code})]";
                                    progbarMullion.Value++;
                                    System.Windows.Forms.Application.DoEvents();
                                }
                            });
                            trans.Commit();
                        }
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 寫入：竪梃 [{Global.DocContent.MullionList.Count}] 參數 [{Global.DocContent.MullionList.Count * 9}]");
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
                            var _plist = Global.DocContent.CurtainPanelList.Where(p => p.INF_ZoneCode == zone.ZoneCode);
                            var _mlist = Global.DocContent.MullionList.Where(m => m.INF_ZoneCode == zone.ZoneCode);
                            writer.WriteLine($"{++idtask},{zone.ZoneCode},,,1");

                            foreach (var p in _plist) writer.WriteLine($"{++idtask},{p.INF_Code},{p.INF_TaskStart},{p.INF_TaskFinish},2");
                            foreach (var m in _mlist) writer.WriteLine($"{++idtask},{m.INF_Code},{m.INF_TaskStart},{m.INF_TaskFinish},2");
                        }
                    }
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbPopupClose = new CommandBinding(cmdPopupClose,
                (sender, e) =>
                {
                    bnQuickStart.IsChecked = false;
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            bnModelInit.Command = cmdModelInit;
            bnSelectPanels.Command = cmdSelectPanels;
            bnSelectMullions.Command = cmdSelectMullions;
            bnZonedPnM.Command = cmdZonedPnM;
            bnPanelResolve.Command = cmdPanelResolve;
            bnLoadData.Command = cmdLoadData;
            bnSaveData.Command = cmdSaveData;
            bnApplyParameters.Command = cmdApplyParameters;
            bnExportElementSchedule.Command = cmdExportElementSchedule;
            bnSearch.Command = cmdSearch;
            bnPopupClose.Command = cmdPopupClose;

            ProcZone.CommandBindings.Add(cbModelInit);
            ProcZone.CommandBindings.AddRange(new CommandBinding[] { cbSelectPanels, cbSelectMullions, cbZonedPnM });
            ProcZone.CommandBindings.Add(cbPanelResolve);
            ProcZone.CommandBindings.Add(cbNavZone);
            ProcZone.CommandBindings.AddRange(new CommandBinding[] { cbLoadData, cbSaveData });
            ProcZone.CommandBindings.Add(cbApplyParameters);
            ProcZone.CommandBindings.Add(cbExportElementSchedule);
            ProcZone.CommandBindings.Add(cbSearch);
            ProcZone.CommandBindings.Add(cbPopupClose);
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
                expDataGridPanels.Header = $"分區[{zi.ZoneCode}]幕墻嵌板數據列表，數量 {_plist.Count()}";
                if (_plist.Count() > 0) expDataGridPanels.IsExpanded = true;

                datagridMullions.ItemsSource = null;
                var _mlist = Global.DocContent.MullionList.Where(ele => ele.INF_ZoneCode == zi.ZoneCode);
                datagridMullions.ItemsSource = _mlist;
                expDataGridMullions.Header = $"分區[{zi.ZoneCode}]幕墻竪梃數據列表，數量 {_mlist.Count()}";
                if (_mlist.Count() > 0) expDataGridMullions.IsExpanded = true;
                navZone.SelectedItem = null;
            }
        }

        #region Command -- bnZonedPnM : 嵌板和竪梃歸類
        private void cbZonedPnM_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前無選擇。");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前選擇 {ids.Count} 構件");
            FilteredElementCollector collectorpanels = new FilteredElementCollector(doc, ids);
            FilteredElementCollector collectormullions = new FilteredElementCollector(doc, ids);
            LogicalAndFilter panel_InstancesFilter = new LogicalAndFilter(
                    new ElementClassFilter(typeof(FamilyInstance)),
                    new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels));
            LogicalAndFilter mullion_InstancesFilter = new LogicalAndFilter(
                new ElementClassFilter(typeof(FamilyInstance)),
                new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallMullions));

            var panels = collectorpanels.WherePasses(panel_InstancesFilter).Where(x => (x as FamilyInstance).Symbol.Name != "空嵌板"); ;
            var mullions = collectormullions.WherePasses(mullion_InstancesFilter);

            //處理嵌板
            if (panels.Count() == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前未選擇幕墻嵌板。");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前篩選 {panels.Count()} 幕墻嵌板");
            SelectedCurtainPanelList.Clear();
            int errorcount_zonecode = 0;
            CurrentZoneInfo = new ZoneInfoBase("Z-00-00-ZZ-00");
            foreach (var _ele in panels)
            {
                CurtainPanelInfo _p = new CurtainPanelInfo(_ele as Autodesk.Revit.DB.Panel);
                _p.INF_Type = CurrentZonePanelType;
                SelectedCurtainPanelList.Add(_p);
                Parameter _param = _ele.get_Parameter("分区区号");
                if (!_param.HasValue)
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - {++errorcount_zonecode}/ 幕墻嵌板[{_ele.Id.IntegerValue}] 未設置分區區號");
                else
                {
                    var _zc = _param.AsString();
                    if (CurrentZoneInfo.ZoneIndex == 0)
                        CurrentZoneInfo = new ZoneInfoBase(_zc);
                    else if (CurrentZoneInfo.ZoneCode != _zc)
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - {++errorcount_zonecode}/ 幕墻嵌板[{_ele.Id.IntegerValue}] 分區區號{_zc}差異({CurrentZoneInfo.ZoneCode})");
                    _p.INF_ZoneInfo = CurrentZoneInfo;
                }

            }
            datagridPanels.ItemsSource = null;
            datagridPanels.ItemsSource = SelectedCurtainPanelList;
            expDataGridPanels.Header = "選擇區域幕墻嵌板數據列表";
            expDataGridPanels.IsExpanded = true;
            if (errorcount_zonecode == 0)
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 選擇的{panels.Count()}幕墻嵌板均已設置相同的分區區號");
            else if (MessageBox.Show(
                $"選擇的{panels.Count()}幕墻嵌板設置的分區區號不相同或有未設置等錯誤，是否繼續？",
                "錯誤 - 幕墻嵌板 - 分區區號",
                MessageBoxButton.YesNo,
                MessageBoxImage.Exclamation,
                MessageBoxResult.No) == MessageBoxResult.No) return;

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
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 選擇的{selectset.GetElementIds().Count}幕墻嵌板已保存至選擇集[{CurrentZoneInfo.FilterName}]");
            }

            Global.DocContent.ZoneList.AddEx(CurrentZoneInfo);
            Global.DocContent.CurtainPanelList.AddRangeEx(SelectedCurtainPanelList);

            //處理竪梃
            if (mullions.Count() == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前未選擇幕墻竪梃。");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前篩選 {mullions.Count()} 幕墻竪梃");
            SelectedMullionList.Clear();
            errorcount_zonecode = 0;
            CurrentZoneInfo = new ZoneInfoBase("Z-00-00-ZZ-00");
            foreach (var _ele in mullions)
            {
                MullionInfo _m = new MullionInfo(_ele as Mullion);
                SelectedMullionList.Add(_m);
                Parameter _param = _ele.get_Parameter("分区区号");
                if (!_param.HasValue)
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - {++errorcount_zonecode}/ 幕墻竪梃[{_ele.Id.IntegerValue}] 未設置分區區號");
                else
                {
                    var _zc = _param.AsString();
                    if (CurrentZoneInfo.ZoneIndex == 0)
                        CurrentZoneInfo = new ZoneInfoBase(_zc);
                    else if (CurrentZoneInfo.ZoneCode != _zc)
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - {++errorcount_zonecode}/ 幕墻竪梃[{_ele.Id.IntegerValue}] 分區區號{_zc}差異({CurrentZoneInfo.ZoneCode})");
                    _m.INF_ZoneInfo = CurrentZoneInfo;
                }

            }
            datagridMullions.ItemsSource = null;
            datagridMullions.ItemsSource = SelectedMullionList;
            expDataGridPanels.Header = "選擇區域幕墻竪梃數據列表";
            if (errorcount_zonecode == 0)
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 選擇的{mullions.Count()}幕墻竪梃均已設置相同的分區區號");
            else if (MessageBox.Show(
                $"選擇的{mullions.Count()}幕墻竪梃設置的分區區號不相同或有未設置等錯誤，是否繼續？",
                "錯誤 - 幕墻竪梃 - 分區區號",
                MessageBoxButton.YesNo,
                MessageBoxImage.Exclamation,
                MessageBoxResult.No) == MessageBoxResult.No) return;

            uidoc.Selection.Elements.Clear();
            foreach (var _ele in mullions) uidoc.Selection.Elements.Add(_ele);

            using (Transaction trans = new Transaction(doc, "CreateMullionGroup"))
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
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 選擇的{selectset.GetElementIds().Count}幕墻竪梃已保存至選擇集[{CurrentZoneInfo.FilterName}]");
            }

            Global.DocContent.ZoneList.AddEx(CurrentZoneInfo);
            Global.DocContent.MullionList.AddRangeEx(SelectedMullionList);

            ZoneHelper.FnContentSerialize();
            expZone.IsExpanded = true;
            subbuttongroup_SelectPanels.Visibility = System.Windows.Visibility.Collapsed;
        }
        #endregion


        #region Command -- bnSelectPanels : 嵌板歸類
        private void cbSelectPanels_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前無選擇。");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前選擇 {ids.Count} 構件");
            FilteredElementCollector collectorpanels = new FilteredElementCollector(doc, ids);
            LogicalAndFilter panel_InstancesFilter =
                new LogicalAndFilter(
                    new ElementClassFilter(typeof(FamilyInstance)),
                    new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels));
            var panels = collectorpanels.WherePasses(panel_InstancesFilter);

            if (panels.Count() == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前未選擇幕墻嵌板。");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前篩選 {panels.Count()} 幕墻嵌板");
            SelectedCurtainPanelList.Clear();
            int errorcount_zonecode = 0;
            CurrentZoneInfo = new ZoneInfoBase("Z-00-00-ZZ-00");
            foreach (var _ele in panels)
            {
                CurtainPanelInfo _p = new CurtainPanelInfo(_ele as Autodesk.Revit.DB.Panel);
                _p.INF_Type = CurrentZonePanelType;
                SelectedCurtainPanelList.Add(_p);
                Parameter _param = _ele.get_Parameter("分区区号");
                if (!_param.HasValue)
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - {++errorcount_zonecode}/ 幕墻嵌板[{_ele.Id.IntegerValue}] 未設置分區區號");
                else
                {
                    var _zc = _param.AsString();
                    if (CurrentZoneInfo.ZoneIndex == 0)
                        CurrentZoneInfo = new ZoneInfoBase(_zc);
                    else if (CurrentZoneInfo.ZoneCode != _zc)
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - {++errorcount_zonecode}/ 幕墻嵌板[{_ele.Id.IntegerValue}] 分區區號{_zc}差異({CurrentZoneInfo.ZoneCode})");
                    _p.INF_ZoneInfo = CurrentZoneInfo;
                }

            }
            datagridPanels.ItemsSource = null;
            datagridPanels.ItemsSource = SelectedCurtainPanelList;
            expDataGridPanels.Header = "選擇區域幕墻嵌板數據列表";
            expDataGridPanels.IsExpanded = true;
            if (errorcount_zonecode == 0)
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 選擇的{panels.Count()}幕墻嵌板均已設置相同的分區區號");
            else if (MessageBox.Show(
                $"選擇的{panels.Count()}幕墻嵌板設置的分區區號不相同或有未設置等錯誤，是否繼續？",
                "錯誤 - 幕墻嵌板 - 分區區號",
                MessageBoxButton.YesNo,
                MessageBoxImage.Exclamation,
                MessageBoxResult.No) == MessageBoxResult.No) return;

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
                trans.Commit();
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 選擇的{selectset.GetElementIds().Count}幕墻嵌板已保存至選擇集[{CurrentZoneInfo.FilterName}]");
            }

            MouseBinding mbind = new MouseBinding(cmdNavZone, new MouseGesture(MouseAction.LeftDoubleClick));
            mbind.CommandParameter = SelectedCurtainPanelList;
            mbind.CommandTarget = datagridPanels;

            Global.DocContent.ZoneList.AddEx(CurrentZoneInfo);
            Global.DocContent.CurtainPanelList.AddRangeEx(SelectedCurtainPanelList);

            ZoneHelper.FnContentSerialize();
            expZone.IsExpanded = true;
            subbuttongroup_SelectPanels.Visibility = System.Windows.Visibility.Collapsed;
        }
        #endregion

        #region Command -- bnSelectMullions : 竪梃歸類
        private void cbSelectMullions_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前無選擇。");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前選擇 {ids.Count} 構件");
            FilteredElementCollector collectormullions = new FilteredElementCollector(doc, ids);
            LogicalAndFilter mullion_InstancesFilter =
                new LogicalAndFilter(
                    new ElementClassFilter(typeof(FamilyInstance)),
                    new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallMullions));
            //.Where(x => (x as FamilyInstance).Symbol.Family.Name != "系統嵌板");
            var mullions = collectormullions.WherePasses(mullion_InstancesFilter);

            if (mullions.Count() == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前未選擇幕墻竪梃。");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前篩選 {mullions.Count()} 幕墻竪梃");
            SelectedMullionList.Clear();
            int errorcount_zonecode = 0;
            CurrentZoneInfo = new ZoneInfoBase("Z-00-00-ZZ-00");
            foreach (var _ele in mullions)
            {
                MullionInfo _m = new MullionInfo(_ele as Mullion);
                SelectedMullionList.Add(_m);
                Parameter _param = _ele.get_Parameter("分区区号");
                if (!_param.HasValue)
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - {++errorcount_zonecode}/ 幕墻竪梃[{_ele.Id.IntegerValue}] 未設置分區區號");
                else
                {
                    var _zc = _param.AsString();
                    if (CurrentZoneInfo.ZoneIndex == 0)
                        CurrentZoneInfo = new ZoneInfoBase(_zc);
                    else if (CurrentZoneInfo.ZoneCode != _zc)
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - {++errorcount_zonecode}/ 幕墻竪梃[{_ele.Id.IntegerValue}] 分區區號{_zc}差異({CurrentZoneInfo.ZoneCode})");
                    _m.INF_ZoneInfo = CurrentZoneInfo;
                }

            }
            datagridMullions.ItemsSource = null;
            datagridMullions.ItemsSource = SelectedMullionList;
            expDataGridPanels.Header = "選擇區域幕墻竪梃數據列表";
            if (errorcount_zonecode == 0)
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 選擇的{mullions.Count()}幕墻竪梃均已設置相同的分區區號");
            else if (MessageBox.Show(
                $"選擇的{mullions.Count()}幕墻竪梃設置的分區區號不相同或有未設置等錯誤，是否繼續？",
                "錯誤 - 幕墻竪梃 - 分區區號",
                MessageBoxButton.YesNo,
                MessageBoxImage.Exclamation,
                MessageBoxResult.No) == MessageBoxResult.No) return;

            uidoc.Selection.Elements.Clear();
            foreach (var _ele in mullions) uidoc.Selection.Elements.Add(_ele);

            using (Transaction trans = new Transaction(doc, "CreateMullionGroup"))
            {
                trans.Start();
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                ICollection<Element> typecollection = collector.OfClass(typeof(SelectionFilterElement)).ToElements();
                SelectionFilterElement selectset = typecollection.Cast<SelectionFilterElement>().FirstOrDefault(ele => ele.Name == CurrentZoneInfo.ZoneCode);
                if (selectset != null) selectset.Clear();
                else selectset = SelectionFilterElement.Create(doc, CurrentZoneInfo.ZoneCode);
                selectset.AddSet(uidoc.Selection.GetElementIds());
                trans.Commit();
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 選擇的{selectset.GetElementIds().Count}幕墻竪梃已保存至選擇集[{CurrentZoneInfo.FilterName}]");
            }

            MouseBinding mbind = new MouseBinding(cmdNavZone, new MouseGesture(MouseAction.LeftDoubleClick));
            mbind.CommandParameter = SelectedMullionList;
            mbind.CommandTarget = datagridMullions;

            Global.DocContent.ZoneList.AddEx(CurrentZoneInfo);
            Global.DocContent.MullionList.AddRangeEx(SelectedMullionList);

            ZoneHelper.FnContentSerialize();
            expZone.IsExpanded = true;
            subbuttongroup_SelectPanels.Visibility = System.Windows.Visibility.Collapsed;
        }
        #endregion

        private void expDataGrid_Expanded(object sender, RoutedEventArgs e)
        {
            if (((Expander)sender).Name == "expDataGridZones") { expDataGridPanels.IsExpanded = false; expDataGridMullions.IsExpanded = false; }
            if (((Expander)sender).Name == "expDataGridPanels") { expDataGridZones.IsExpanded = false; expDataGridMullions.IsExpanded = false; }
            if (((Expander)sender).Name == "expDataGridMullions") { expDataGridZones.IsExpanded = false; expDataGridPanels.IsExpanded = false; }
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
                if (MessageBox.Show($"加載新分區進度文件 {ofd.FileName}？\n\n現有数据将被新數據覆蓋，且不可恢復。單擊確認繼續，取消則不會有任何操作。", "加載新進度數據...",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Question,
                    MessageBoxResult.OK) == MessageBoxResult.OK)
                {
                    ZoneHelper.FnLoadSimpleZoneScheduleData(ofd.FileName);
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 加載新分區進度文件{ofd.FileName}.");
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
                        IsSearchRangeMullion = true;
                        break;
                    case false:
                        IsSearchRangeZone = false;
                        IsSearchRangePanel = false;
                        IsSearchRangeMullion = false;
                        break;
                    default:
                        break;
                }
            }
            else
            {
                if (IsSearchRangeZone && IsSearchRangePanel && IsSearchRangeMullion) IsSearchRangeAll = true;
                else if (!IsSearchRangeZone && !IsSearchRangePanel && !IsSearchRangeMullion) IsSearchRangeAll = false;
                else IsSearchRangeAll = null;
            }
            ((CheckBox)sender).GetBindingExpression(CheckBox.IsCheckedProperty).UpdateTarget();
        }

    }


}
