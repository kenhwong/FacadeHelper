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

            if (Global.DocContent.CurtainPanelList.Count > 0) Global.DocContent.Lookup_CurtainPanels = Global.DocContent.CurtainPanelList.ToLookup(g => g.INF_ZoneCode);
            Global.DocContent.ParameterInfoList = ParameterHelper.RawGetProjectParametersInfo(doc);

            navZone.ItemsSource = Global.DocContent.ZoneList;
            datagridZones.ItemsSource = Global.DocContent.ZoneList;
        }


        #region 初始化 Command

        private RoutedCommand cmdModelInit = new RoutedCommand();
        private RoutedCommand cmdSelectPanels = new RoutedCommand();
        private RoutedCommand cmdSelectPanelsCheckData = new RoutedCommand();
        private RoutedCommand cmdSelectPanelsCreateSelectionSet = new RoutedCommand();
        private RoutedCommand cmdPanelResolve = new RoutedCommand();

        private RoutedCommand cmdNavZone = new RoutedCommand();

        private RoutedCommand cmdLoadData = new RoutedCommand();
        private RoutedCommand cmdSaveData = new RoutedCommand();
        private RoutedCommand cmdApplyParameters = new RoutedCommand();

        private RoutedCommand cmdSearch = new RoutedCommand();

        private void InitializeCommand()
        {
            CommandBinding cbModelInit = new CommandBinding(cmdModelInit, cbModelInit_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbSelectPanels = new CommandBinding(cmdSelectPanels,
                (sender, e) =>
                {
                    if (subbuttongroup_SelectPanels.Visibility == System.Windows.Visibility.Visible)
                        subbuttongroup_SelectPanels.Visibility = System.Windows.Visibility.Collapsed;
                    else
                        subbuttongroup_SelectPanels.Visibility = System.Windows.Visibility.Visible;
                    listInformation.Items.Clear();
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbSelectPanelsCheckData = new CommandBinding(cmdSelectPanelsCheckData, cbSelectPanelsCheckData_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbSelectPanelsCreateSelectionSet = new CommandBinding(cmdSelectPanelsCreateSelectionSet, cbSelectPanelsCreateSelectionSet_Executed, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbPanelResolve = new CommandBinding(cmdPanelResolve,
                (sender, e) =>
                {
                    Global.DocContent.ScheduleElementList.Clear();
                    Global.DocContent.DeepElementList.Clear();
                    foreach (var zn in Global.DocContent.ZoneList) ZoneHelper.FnResolveZone(uidoc, zn, ref listInformation, ref txtProcessInfo);
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

                                if (IsRealTimeProgress)
                                {
                                    txtProcessInfo.Content = $"當前處理進度：[分區：{p.INF_ZoneCode}] - [幕墻嵌板：{p.INF_ElementId}, {p.INF_Code})]";
                                    progbarPanel.Value++;
                                    System.Windows.Forms.Application.DoEvents();
                                }
                            });
                            trans.Commit();
                        }
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 寫入：嵌板 [{Global.DocContent.CurtainPanelList.Count}] 參數 [{Global.DocContent.CurtainPanelList.Count * 7}]");

                        using (Transaction trans = new Transaction(doc, "Apply_Parameters_ScheduleElements"))
                        {
                            trans.Start();
                            progbarElement.Maximum = Global.DocContent.ScheduleElementList.Count;
                            progbarElement.Value = 0;
                            Global.DocContent.ScheduleElementList.ForEach(ele =>
                            {
                                Element _element = doc.GetElement(new ElementId(ele.INF_ElementId));
                                _element.get_Parameter("立面朝向").Set(ele.INF_Direction);
                                _element.get_Parameter("立面系统").Set(ele.INF_System);
                                _element.get_Parameter("立面楼层").Set(ele.INF_Level);
                                _element.get_Parameter("构件分项").Set(ele.INF_Type);
                                _element.get_Parameter("分区序号").Set(ele.INF_ZoneID);
                                _element.get_Parameter("分区区号").Set(ele.INF_ZoneCode);
                                _element.get_Parameter("分区编码").Set(ele.INF_Code);

                                if (IsRealTimeProgress)
                                {
                                    txtProcessInfo.Content = $"當前處理進度：[嵌板：{ele.INF_HostCurtainPanel.INF_Code}] - [明細構件：{ele.INF_ElementId}, {ele.INF_Code})]";
                                    progbarElement.Value++;
                                    System.Windows.Forms.Application.DoEvents();
                                }
                            });
                            trans.Commit();
                        }
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 寫入：明細構件 [{Global.DocContent.ScheduleElementList.Count}] 參數 [{Global.DocContent.ScheduleElementList.Count * 7}]");
                    }
                    //catch (Exception ex)
                    {
                        //MessageBox.Show($"Source:{ex.Source}\nMessage:{ex.Message}\nTargetSite:{ex.TargetSite}\nStackTrace:{ex.StackTrace}");
                    }
                    //listboxOutput.SelectedIndex = listboxOutput.Items.Add($"写入[幕墙嵌板]:{Global.DocContent.CurtainPanelList.Count}，参数:{Global.DocContent.CurtainPanelList.Count * 6}...");

                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            bnModelInit.Command = cmdModelInit;
            bnSelectPanels.Command = cmdSelectPanels;
            bnSelectPanels_CheckData.Command = cmdSelectPanelsCheckData;
            bnSelectPanels_CreateSelectionSet.Command = cmdSelectPanelsCreateSelectionSet;
            bnPanelResolve.Command = cmdPanelResolve;
            bnLoadData.Command = cmdLoadData;
            bnSaveData.Command = cmdSaveData;
            bnApplyParameters.Command = cmdApplyParameters;
            bnSearch.Command = cmdSearch;

            ProcZone.CommandBindings.Add(cbModelInit);
            ProcZone.CommandBindings.AddRange(new CommandBinding[] { cbSelectPanels, cbSelectPanelsCheckData, cbSelectPanelsCreateSelectionSet });
            ProcZone.CommandBindings.Add(cbPanelResolve);
            ProcZone.CommandBindings.Add(cbNavZone);
            ProcZone.CommandBindings.AddRange(new CommandBinding[] { cbLoadData, cbSaveData });
            ProcZone.CommandBindings.Add(cbApplyParameters);
            ProcZone.CommandBindings.Add(cbSearch);
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
                expDataGridPanels.IsExpanded = true;
                navPanels.ItemsSource = _plist;
                expPanel.Header = $"幕墻嵌板：{_plist.Count()}，分區[{zi.ZoneCode}]";
                expPanel.IsExpanded = true;
                datagridScheduleElements.ItemsSource = null;
                var _elist = Global.DocContent.ScheduleElementList.Where(ele => ele.INF_ZoneCode == zi.ZoneCode);
                datagridScheduleElements.ItemsSource = _elist;
                expDataGridScheduleElements.Header = $"分區[{zi.ZoneCode}]明細構件數據列表，數量 {_elist.Count()}";
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
                expDataGridScheduleElements.Header = $"分區[{pi.INF_ZoneCode}]，嵌板[{pi.INF_Code}]明細構件數據列表，數量 {_elist.Count()}";
                expDataGridScheduleElements.IsExpanded = true;
                navPanels.SelectedItem = null;
            }
        }


        private void cbSelectPanelsCheckData_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            //parentWin.Hide();
            //var _p_sel_list = uidoc.Selection.PickObjects(ObjectType.Element, "選擇同一分區的所有嵌板");
            //parentWin.Visibility = System.Windows.Visibility.Visible;

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            if (ids.Count == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前無選擇。");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前選擇 {ids.Count} 構件");
            FilteredElementCollector collector = new FilteredElementCollector(doc, ids);
            LogicalAndFilter cwpanel_InstancesFilter =
                new LogicalAndFilter(
                    new ElementClassFilter(typeof(FamilyInstance)),
                    new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels));
            var eles = collector
                .WherePasses(cwpanel_InstancesFilter)
                .Where(x => (x as FamilyInstance).Symbol.Family.Name != "系統嵌板");

            if (eles.Count() == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前未選擇幕墻嵌板。");
                return;
            }
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 當前篩選 {eles.Count()} 幕墻嵌板");
            SelectedCurtainPanelList.Clear();
            int errorcount_zonecode = 0;
            CurrentZoneInfo = new ZoneInfoBase("Z-00-00-ZZ-00");
            foreach (var _ele in eles)
            {
                CurtainPanelInfo _gp = new CurtainPanelInfo(_ele as Autodesk.Revit.DB.Panel);
                _gp.INF_Type = CurrentZonePanelType;
                SelectedCurtainPanelList.Add(_gp);
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
                    _gp.INF_ZoneInfo = CurrentZoneInfo;
                }

            }
            if (errorcount_zonecode == 0) listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 選擇的{eles.Count()}幕墻嵌板均已設置相同的分區區號");
            datagridPanels.ItemsSource = SelectedCurtainPanelList;
            expDataGridPanels.Header = "選擇區域幕墻嵌板數據列表";

            uidoc.Selection.Elements.Clear();
            foreach (var _ele in eles) uidoc.Selection.Elements.Add(_ele);
        }

        private void cbSelectPanelsCreateSelectionSet_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            using (Transaction trans = new Transaction(doc, "CreateGroup"))
            {
                trans.Start();
                var sfe = SelectionFilterElement.Create(doc, CurrentZoneInfo.ZoneCode);
                sfe.AddSet(uidoc.Selection.GetElementIds());
                trans.Commit();
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:hh:MM:ss} - 選擇的{sfe.GetElementIds().Count}幕墻嵌板已保存至選擇集[{CurrentZoneInfo.FilterName}]");
            }

            /**
            ListBoxItem _zoneitem = new ListBoxItem();
            _zoneitem.Content = CurrentZoneInfo;
            _zoneitem.Name = CurrentZoneInfo.ZoneCode.Replace("-", ""); 
            _zoneitem.Tag = CurrentZoneInfo;
            **/

            MouseBinding mbind = new MouseBinding(cmdNavZone, new MouseGesture(MouseAction.LeftDoubleClick));
            mbind.CommandParameter = SelectedCurtainPanelList;
            mbind.CommandTarget = datagridPanels;
            //_zoneitem.InputBindings.Add(mbind);
            //navZone.Items.Add(_zoneitem);

            Global.DocContent.ZoneList.Add(CurrentZoneInfo);
            Global.DocContent.CurtainPanelList.AddRange(SelectedCurtainPanelList);

            ZoneHelper.FnContentSerialize();
            expZone.IsExpanded = true;
            subbuttongroup_SelectPanels.Visibility = System.Windows.Visibility.Collapsed;
        }

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
        }
        #endregion

        private void bnTest_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show($"CurrentZonePanelType：{CurrentZonePanelType}");
            MessageBox.Show($"Range Zone：{IsSearchRangeZone}\nRange Panel：{IsSearchRangePanel}\nRange Element：{IsSearchRangeElement}\nRange All：{IsSearchRangeAll}");
        }

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
    }

    public class IntToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (int)value == int.Parse(parameter.ToString());
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var data = (bool)value;
            if (data)
            {
                return System.Convert.ToInt32(parameter);
            }
            return -1;
        }
    }

}
