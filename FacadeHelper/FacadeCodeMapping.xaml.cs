using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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
        private List<ElementFabricationInfo> _currentFabricationList = new List<ElementFabricationInfo>();
        private ObservableCollection<ScheduleElementInfo> _currentElementInfoList = new ObservableCollection<ScheduleElementInfo>();
        public List<Element> CurrentElementList { get { return _currentElementList; } set { _currentElementList = value; OnPropertyChanged(nameof(CurrentElementList)); } }
        public List<ElementFabricationInfo> CurrentFabricationList { get { return _currentFabricationList; } set { _currentFabricationList = value; OnPropertyChanged(nameof(CurrentFabricationList)); } }
        public ObservableCollection<ScheduleElementInfo> CurrentElementInfoList { get { return _currentElementInfoList; } set { _currentElementInfoList = value; OnPropertyChanged(nameof(CurrentElementInfoList)); } }

        public FacadeCodeMapping(ExternalCommandData commandData)
        {
            InitializeComponent();
            this.DataContext = this;

            cdata = commandData;
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;
            sids = uidoc.Selection.Elements.Cast<Element>().Select(e => e.Id).ToList<ElementId>();

            InitializeCommand();


        }

        private void ProcessSelection()
        {
            #region Selection 筛选
            LogicalAndFilter GM_InstancesFilter = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_GenericModel));
            LogicalAndFilter P_InstancesFilter = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels));
            LogicalAndFilter W_InstancesFilter = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_Windows));
            List<Element> listselected = new List<Element>();

            if (sids.Count == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - SELECT: ELE/NONE(SET TO ALL).");
                listselected = new FilteredElementCollector(doc).WherePasses(GM_InstancesFilter).ToElements().ToList();
            }
            else
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - SELECT: ELE/{sids.Count}.");
                var pwlist = new FilteredElementCollector(doc, sids).WherePasses(P_InstancesFilter).Union((new FilteredElementCollector(doc, sids)).WherePasses(W_InstancesFilter));
                foreach (Element ele in pwlist) listselected.AddRange((ele as FamilyInstance).GetSubComponentIds().Select(pwid => doc.GetElement(pwid)));
            }


            uidoc.Selection.Elements.Clear();
            Parameter _parameter;
            listselected.ForEach(ele =>
            {
                _parameter = ele.get_Parameter("分项");
                if (_parameter?.HasValue == true && int.TryParse(_parameter?.AsString(), out int _type)) if (_type == CurrentSelectedType)
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
                            INF_Code = ele.get_Parameter("分区编码")?.AsString() ?? ""
                        };
                        CurrentElementInfoList.Add(ei);
                    }
            });

            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - FILTERED: ELE/{uidoc.Selection.Elements.Size}/{CurrentElementList.Count}.");
            #endregion


            return;
        }

        private RoutedCommand cmdRefreshList1 = new RoutedCommand();

        private RoutedCommand cmdApplySelection = new RoutedCommand();
        private RoutedCommand cmdParamsRefresh = new RoutedCommand();
        private RoutedCommand cmdSelectParam = new RoutedCommand();
        public RoutedCommand CmdSelectParam { get { return cmdSelectParam; } set { cmdSelectParam = value; OnPropertyChanged(nameof(CmdSelectParam)); } }
        private void InitializeCommand()
        {
            CommandBinding cbRefreshList1 = new CommandBinding(cmdRefreshList1, (sender, e) => ProcessSelection(), (sender, e) => { e.CanExecute = true; e.Handled = true; });


            CommandBinding cbParamsRefresh = new CommandBinding(cmdParamsRefresh, (sender, e) =>
            {
                ProcessSelection();
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbApplySelection = new CommandBinding(cmdApplySelection, (sender, e) =>
            {
                ProcessSelection();

                using (Transaction trans = new Transaction(doc, "CreateSelectionFilter"))
                {
                    trans.Start();

                    trans.Commit();
                }

            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            bnRefreshList1.Command = cmdRefreshList1;
            //bnApplySelection.Command = cmdApplySelection;
            //bnParamsRefresh.Command = cmdParamsRefresh;

            ProcFCM.CommandBindings.AddRange(new CommandBinding[]
            {
                cbRefreshList1,
                cbParamsRefresh,
                cbApplySelection
            });
        }

        private void listInformation_SelectionChanged(object sender, SelectionChangedEventArgs e) { var lb = sender as ListBox; lb.ScrollIntoView(lb.Items[lb.Items.Count - 1]); }

        private void chkbox_Checked(object sender, RoutedEventArgs e)
        {


            ((CheckBox)sender).GetBindingExpression(CheckBox.IsCheckedProperty).UpdateTarget();
        }

    }
}
