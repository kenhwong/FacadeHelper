using AqlaSerializer;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
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

namespace FacadeHelper
{
    /// <summary>
    /// Interaction logic for SelectFilter.xaml
    /// </summary>
    public partial class FacadeCodeMapping : UserControl, INotifyPropertyChanged
    {
        public Window ParentWin { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private int _currentSelectedType = 1;
        public int CurrentSelectedType { get { return _currentSelectedType; } set { _currentSelectedType = value; OnPropertyChanged(nameof(CurrentSelectedType)); } }

        private UIApplication uiapp;
        private UIDocument uidoc;
        private Document doc;
        private ExternalCommandData cdata;
        private ICollection<ElementId> sids;

        private List<Element> _currentElementList = new List<Element>();
        private ObservableCollection<ScheduleElementInfo> _currentElementInfoList = new ObservableCollection<ScheduleElementInfo>();
        private ObservableCollection<ScheduleElementInfo> _appliedElementInfoList = new ObservableCollection<ScheduleElementInfo>();
        public List<Element> CurrentElementList { get { return _currentElementList; } set { _currentElementList = value; OnPropertyChanged(nameof(CurrentElementList)); } }
        public ObservableCollection<ScheduleElementInfo> CurrentElementInfoList { get { return _currentElementInfoList; } set { _currentElementInfoList = value; OnPropertyChanged(nameof(CurrentElementInfoList)); } }
        public ObservableCollection<ScheduleElementInfo> AppliedElementInfoList { get { return _appliedElementInfoList; } set { _appliedElementInfoList = value; OnPropertyChanged(nameof(AppliedElementInfoList)); } }

        private ObservableCollection<string> _paramListSource = new ObservableCollection<string>();
        public ObservableCollection<string> ParamListSource { get { return _paramListSource; } set { _paramListSource = value; OnPropertyChanged(nameof(ParamListSource)); } }
        private ObservableCollection<string> _currentZoneList = new ObservableCollection<string>();
        public ObservableCollection<string> CurrentZoneList { get { return _currentZoneList; } set { _currentZoneList = value; OnPropertyChanged(nameof(CurrentZoneList)); } }

        private List<string> _zoneListSource = new List<string>();
        public List<string> ZoneListSource { get { return _zoneListSource; } set { _zoneListSource = value; OnPropertyChanged(nameof(ZoneListSource)); } }

        private ObservableCollection<ElementFabricationInfo> _currentFabricationList = new ObservableCollection<ElementFabricationInfo>();
        public ObservableCollection<ElementFabricationInfo> CurrentFabricationList
        {
            get { return _currentFabricationList; }
            set
            {
                _currentFabricationList = value;
                OnPropertyChanged(nameof(CurrentFabricationList));
                OnPropertyChanged(nameof(CurrentFabricationList_ElementCode));
                OnPropertyChanged(nameof(CurrentFabricationList_OrderCode));
            }
        }
        public ICollectionView CurrentFabricationList_ViewElementCode { get { return CollectionViewSource.GetDefaultView(CurrentFabricationList); } }
        public ICollectionView CurrentFabricationList_ViewOrderCode { get { return CollectionViewSource.GetDefaultView(CurrentFabricationList); } }
        public ObservableCollection<string> CurrentFabricationList_ElementCode { get { return new ObservableCollection<string>(CurrentFabricationList.Select(f => f.ElementCode).Distinct()); } }
        public ObservableCollection<string> CurrentFabricationList_OrderCode { get { return new ObservableCollection<string>(CurrentFabricationList.Select(f => f.OrderCode).Distinct()); } }

        public FacadeCodeMapping(ExternalCommandData commandData)
        {
            InitializeComponent();
            this.DataContext = this;

            cdata = commandData;
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;
            sids = uidoc.Selection.Elements.Cast<Element>().Select(e => e.Id).ToList<ElementId>();

            /**
            if (Global.ZoneLayerList.Count == 0) Global.ZoneLayerList = ZoneHelper.FnZoneDataDeserialize();
            ZoneListSource = Global.ZoneLayerList.Select(zl => zl.ZoneCode).Distinct().ToList();
            **/

            //分区统计2
            var pwlist = new FilteredElementCollector(doc).WherePasses(new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels)))
                .Union(new FilteredElementCollector(doc).WherePasses(new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_Windows))));
            Parameter _parameter_zone;
            foreach (var pw in pwlist)
            {
                _parameter_zone = pw.get_Parameter("分区区号");
                if (_parameter_zone?.HasValue == true)
                {
                    string _zone = _parameter_zone?.AsString();
                    if (!ZoneListSource.Contains(_zone)) ZoneListSource.Add(_zone);
                }
            }


            if (CurrentZoneList.Count == 0) ZoneListSource.ForEach(z => CurrentZoneList.Add(z));

            InitializeCommand();


        }

        private void FuncProcessPreSelection()
        {
            #region Selection 预筛选
            LogicalAndFilter GM_InstancesFilter = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_GenericModel));
            LogicalAndFilter P_InstancesFilter = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels));
            LogicalAndFilter W_InstancesFilter = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_Windows));
            List<Element> listselected = new List<Element>();
            CurrentElementList.Clear();
            CurrentElementInfoList.Clear();

            if (sids.Count == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - SELECT: ELE/NONE(SET TO ALL).");
                var pwlist = new FilteredElementCollector(doc).WherePasses(P_InstancesFilter)
                    .Union(new FilteredElementCollector(doc).WherePasses(W_InstancesFilter));
                foreach (Element ele in pwlist)
                {
                    var subids = (ele as FamilyInstance).GetSubComponentIds();
                    foreach (var eid in subids)
                    {
                        var e = doc.GetElement(eid);
                        listselected.Add(e);
                    }
                }
                foreach (Element ele in pwlist) listselected.AddRange((ele as FamilyInstance).GetSubComponentIds().Select(pwid => doc.GetElement(pwid)));
            }
            else
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - SELECT: ELE/{sids.Count}.");
                var pwlist = new FilteredElementCollector(doc, sids).WherePasses(P_InstancesFilter).Union((new FilteredElementCollector(doc, sids)).WherePasses(W_InstancesFilter));
                foreach (Element ele in pwlist) listselected.AddRange((ele as FamilyInstance).GetSubComponentIds().Select(pwid => doc.GetElement(pwid)));
            }

            uidoc.Selection.Elements.Clear();
            Parameter _parameter_type;
            Parameter _parameter_zone;
            int _i = 0;
            listselected.ForEach(ele =>
            {
                double v = _i++ * 1.0 / listselected.Count;
                lblStatus1.Content = string.Format("{0:F2}%", v * 100);
                barStatusRefreshList1.Value = v * 100;
                System.Windows.Forms.Application.DoEvents();

                _parameter_type = ele.get_Parameter("分项");
                _parameter_zone = ele.get_Parameter("分区区号");
                if (_parameter_type?.HasValue == true && _parameter_zone?.HasValue == true && int.TryParse(_parameter_type?.AsString(), out int _type))
                {
                    string _zone = _parameter_zone?.AsString();
                    if (_type == CurrentSelectedType && CurrentZoneList.Contains(_zone.ToUpper()))
                    {
                        uidoc.Selection.Elements.Add(ele);
                        CurrentElementList.Add(ele);
                        ScheduleElementInfo ei = new ScheduleElementInfo()
                        {
                            INF_ElementId = ele.Id.IntegerValue,
                            INF_Name = ele.Name,
                            INF_Type = CurrentSelectedType,
                            INF_Index = ele.get_Parameter("分区序号")?.AsInteger() ?? 0,
                            INF_Direction = ele.get_Parameter("立面朝向")?.AsString() ?? "",
                            INF_System = ele.get_Parameter("立面系统")?.AsString() ?? "",
                            INF_Level = ele.get_Parameter("立面楼层")?.AsInteger() ?? 0,
                            INF_ZoneCode = ele.get_Parameter("分区区号")?.AsString() ?? "",
                            INF_Code = ele.get_Parameter("分区编码")?.AsString() ?? ele.Id.ToString()
                        };
                        CurrentElementInfoList.Add(ei);
                    }
                }
            });

            lblStatus1.Content = string.Format("{0:F2}%", 100);
            barStatusRefreshList1.Value = 100;
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - FILTERED: ELE/{uidoc.Selection.Elements.Size}/{CurrentElementList.Count}.");
            #endregion


            return;
        }

        private RoutedCommand cmdRefreshList1 = new RoutedCommand();
        private RoutedCommand cmdApplySelection = new RoutedCommand();
        private RoutedCommand cmdApplyParam = new RoutedCommand();
        private RoutedCommand cmdZPopApply = new RoutedCommand();
        private RoutedCommand cmdZPopClose = new RoutedCommand();
        private RoutedCommand cmdZPopClear = new RoutedCommand();
        private RoutedCommand cmdInitZone = new RoutedCommand();

        private RoutedCommand cmdLoadOrder = new RoutedCommand();

        private void InitializeCommand()
        {
            CommandBinding cbRefreshList1 = new CommandBinding(cmdRefreshList1, (sender, e) => FuncProcessPreSelection(), (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbApplySelection = new CommandBinding(cmdApplySelection, (sender, e) => FuncApplySelection(), (sender, e) =>
            {
                if (Lst1.SelectedItems.Count > 0)
                {
                    e.CanExecute = true;
                    e.Handled = true;
                }
            });

            CommandBinding cbApplyParam = new CommandBinding(cmdApplyParam, (sender, e) =>
            {
                Element ele = null;
                using (Transaction trans = new Transaction(doc, "Apply Params"))
                {
                    trans.Start();
                    int _iae = 0;
                    foreach (var ei in AppliedElementInfoList)
                    {
                        double v = _iae++ * 1.0 / AppliedElementInfoList.Count;
                        lblStatusApplyParams.Content = string.Format("{0:F2}%", v * 100);
                        barStatusApplyParams.Value = v * 100;
                        System.Windows.Forms.Application.DoEvents();

                        ele = doc.GetElement(new ElementId(ei.INF_ElementId));
                        Parameter pa = ele.get_Parameter((string)ParamNameList.SelectedValue);
                        if (pa == null)
                        {
                            MessageBox.Show($"构件[{ei.INF_ElementId}]/[{ei.INF_Code}]中参数【{(string)ParamNameList.SelectedValue}】不存在", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                            return;
                        }
                        switch (pa.StorageType)
                        {
                            case StorageType.String:
                                pa.Set((string)ParamValueList.Text);
                                break;
                            case StorageType.Integer:
                                if (!int.TryParse((string)ParamValueList.Text, out int i))
                                {
                                    MessageBox.Show($"构件[{ei.INF_ElementId}]/[{ei.INF_Code}]中整数型参数【{(string)ParamNameList.SelectedValue}】不匹配输入值", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                                    return;
                                }
                                pa.Set(i);
                                break;
                            case StorageType.Double:
                                if (!double.TryParse((string)ParamValueList.Text, out double d))
                                {
                                    MessageBox.Show($"构件[{ei.INF_ElementId}]/[{ei.INF_Code}]中实数型参数【{(string)ParamNameList.SelectedValue}】不匹配输入值", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                                    return;
                                }
                                pa.Set(d);
                                break;
                            default:
                                MessageBox.Show($"构件[{ei.INF_ElementId}]/[{ei.INF_Code}]中参数【{(string)ParamNameList.SelectedValue}】类型可能有误", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                                break;
                        }
                        listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - WRITE: ELE/[{ei.INF_ElementId}]/[{ei.INF_Code}], {(string)ParamNameList.SelectedValue}: {(string)ParamValueList.Text}.");
                    }
                    trans.Commit();
                    lblStatusApplyParams.Content = string.Format("{0:F0}%", 100);
                    barStatusApplyParams.Value = 100;
                }

            }, (sender, e) =>
            {
                if (ParamNameList.SelectedIndex >= 0)
                {
                    e.CanExecute = true;
                    e.Handled = true;
                }
            });

            CommandBinding cbZPopApply = new CommandBinding(cmdZPopApply, (sender, e) =>
            {
                CurrentZoneList.Clear();
                if (LstZone.SelectedIndex != -1) foreach (var z in LstZone.SelectedItems) CurrentZoneList.Add((string)z);

                if (LstZone.SelectedItems.Count == ZoneListSource.Count)
                    txtSelectedZoneList.Content = "所有分区";
                else if (LstZone.SelectedItems.Count == 0)
                {
                    LstZone.SelectAll();
                    ZoneListSource.ForEach(z => CurrentZoneList.Add(z));
                    txtSelectedZoneList.Content = "所有分区";
                }
                else
                    txtSelectedZoneList.Content = string.Join(", ", CurrentZoneList.ToArray());


                bnAddZone.IsChecked = false;
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbInitZone = new CommandBinding(cmdInitZone, (sender, e) =>
            {
                var pwlist = new FilteredElementCollector(doc)
                .WherePasses(new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels)))
                    .Union(new FilteredElementCollector(doc)
                    .WherePasses(new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_Windows))));

                #region 设置分区区号参数
                using (Transaction trans = new Transaction(doc, "CreateZoneParameter"))
                {
                    trans.Start();
                    CategorySet _catset = new CategorySet();
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_CurtainWallPanels));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel));
                    _catset.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_Windows));
                    ParameterHelper.RawCreateProjectParameter(doc.Application, "分区区号", ParameterType.Text, true, _catset, BuiltInParameterGroup.PG_DATA, true);
                    trans.Commit();
                }

                using (Transaction trans = new Transaction(doc, "SetZoneParameter"))
                {
                    trans.Start();
                    Parameter _parameter_zone;
                    foreach (var pw in pwlist)
                    {
                        _parameter_zone = pw.get_Parameter("分区区号");
                        if (_parameter_zone?.HasValue == false) _parameter_zone.Set("Z-00-00-00-00");
                        var eles = (pw as FamilyInstance).GetSubComponentIds().Select(eid => doc.GetElement(eid));
                        foreach (var ele in eles)
                        {
                            _parameter_zone = ele.get_Parameter("分区区号");
                            if (_parameter_zone?.HasValue == false) _parameter_zone.Set("Z-00-00-00-00");
                        }
                    }
                    trans.Commit();
                }

                #endregion
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbLoadOrder = new CommandBinding(cmdLoadOrder, (sender, e) =>
            {
                Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog()
                {
                    InitialDirectory = System.IO.Path.GetDirectoryName(Global.DataFile),
                    DefaultExt = "*.csv",
                    Filter = "订单交换文件(*.csv)|*.csv|All(*.*)|*.*"
                };
                if (ofd.ShowDialog() == true)
                {
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - READ: DATA, {ofd.FileName}");
                    using (StreamReader reader = new StreamReader(ofd.FileName))
                    {

                        string dataline;
                        reader.ReadLine();
                        string linezonecode = string.Empty;
                        while ((dataline = reader.ReadLine()) != null)
                        {
                            string[] rowdata = dataline.Split(',');
                            CurrentFabricationList.Add(new ElementFabricationInfo() { ElementCode = rowdata[0], FabrQuantity = int.Parse(rowdata[1]), OrderCode = rowdata[2] });
                        }
                    }
                }
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbZPopClear = new CommandBinding(cmdZPopClear, (sender, e) => { LstZone.SelectedIndex = -1; }, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbZPopClose = new CommandBinding(cmdZPopClose, (sender, e) => { bnAddZone.IsChecked = false; }, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            bnRefreshList1.Command = cmdRefreshList1;
            bnApplySelection.Command = cmdApplySelection;
            bnApplyParam.Command = cmdApplyParam;
            bnZPopApply.Command = cmdZPopApply;
            bnZPopClose.Command = cmdZPopClose;
            bnZPopClear.Command = cmdZPopClear;
            bnLoadOrder.Command = cmdLoadOrder;
            bnInitZone.Command = cmdInitZone;

            ProcFCM.CommandBindings.AddRange(new CommandBinding[]
            {
                cbRefreshList1,
                cbApplySelection,
                cbApplyParam,
                cbZPopApply,
                cbZPopClear,
                cbZPopClose,
                cbInitZone,
                cbLoadOrder
            });
        }

        private void FuncApplySelection()
        {
            #region Selection 确定选择
            AppliedElementInfoList.Clear();
            ParamListSource.Clear();
            List<Element> listappliedselection = new List<Element>();
            uidoc.Selection.Elements.Clear();

            //初始化列表选择项数据
            int _isel = 0;
            foreach (var ei in Lst1.SelectedItems.Cast<ScheduleElementInfo>())
            {
                double v = _isel++ * 1.0 / Lst1.SelectedItems.Count;
                lblStatusSelection.Content = string.Format("{0:F0}%", v * 100);
                barStatusSelection.Value = v * 100;
                System.Windows.Forms.Application.DoEvents();

                AppliedElementInfoList.Add(ei);
                Element ele = doc.GetElement(new ElementId(ei.INF_ElementId));
                ParameterSet parameters = ele.Parameters;
                foreach (Parameter parameter in parameters) if (!parameter.IsReadOnly) ParamListSource.Add(parameter.Definition.Name);
                listappliedselection.Add(ele);
                uidoc.Selection.Elements.Add(ele);
            }

            //清理重复参数
            HashSet<string> hs = new HashSet<string>(ParamListSource);
            ParamListSource.Clear();
            hs.ToList().ForEach(d => ParamListSource.Add(d));

            lblStatusSelection.Content = string.Format("{0:F0}%", 100);
            barStatusSelection.Value = 100;
            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - SELECTED: ELE/{uidoc.Selection.Elements.Size}/{CurrentElementList.Count}.");
            #endregion
        }

        private void listInformation_SelectionChanged(object sender, SelectionChangedEventArgs e) { var lb = sender as ListBox; lb.ScrollIntoView(lb.Items[lb.Items.Count - 1]); }

        private void chkbox_Checked(object sender, RoutedEventArgs e)
        {


            ((CheckBox)sender).GetBindingExpression(CheckBox.IsCheckedProperty).UpdateTarget();
        }

    }
}
