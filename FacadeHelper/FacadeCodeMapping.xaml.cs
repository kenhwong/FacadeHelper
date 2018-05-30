using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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

        public event PropertyChangedEventHandler CurrentTypeChanged;
        protected void OnCurrentTypeChanged()
        {
            //List2 change.
            ProcessSelection();
            CurrentTypeChanged?.Invoke(this, PropertyChangedEventArgs.Empty as PropertyChangedEventArgs);
        }

        private int _currentSelectedType = 1;
        public int CurrentSelectedType { get { return _currentSelectedType; } set { _currentSelectedType = value; OnCurrentTypeChanged(); } }

        private ObservableCollection<ScheduleElementInfo> _listElements = new ObservableCollection<ScheduleElementInfo>();
        private ObservableCollection<string> _paramListFiltered = new ObservableCollection<string>();
        public ObservableCollection<ScheduleElementInfo> ListElements { get { return _listElements; } set { _listElements = value; OnPropertyChanged(nameof(ListElements)); } }
        public ObservableCollection<string> ParamListFiltered { get { return _paramListFiltered; } set { _paramListFiltered = value; OnPropertyChanged(nameof(ParamListFiltered)); } }


        private UIApplication uiapp;
        private UIDocument uidoc;
        private Document doc;
        private ExternalCommandData cdata;
        private Autodesk.Revit.UI.Selection.Selection sel; 

        private List<Element> _currentElementList = new List<Element>();
        private List<ElementFabricationInfo> _currentFabricationList = new List<ElementFabricationInfo>();
        public List<Element> CurrentElementList { get { return _currentElementList; } set { _currentElementList = value; OnPropertyChanged(nameof(CurrentElementList)); } }
        public List<ElementFabricationInfo> CurrentFabricationList { get { return _currentFabricationList; } set { _currentFabricationList = value; OnPropertyChanged(nameof(CurrentFabricationList)); } }

        public FacadeCodeMapping(ExternalCommandData commandData)
        {
            InitializeComponent();
            this.DataContext = this;

            cdata = commandData;
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;
            sel = uidoc.Selection;

            InitializeCommand();


        }

        private void ProcessSelection()
        {
            FilteredElementCollector ecollector;
            ICollection<ElementId> ids = sel.GetElementIds();
            if (ids.Count == 0)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - SELECT: ELE/NONE(SET TO ALL).");
                ecollector = new FilteredElementCollector(doc);
            }
            else
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - SELECT: ELE/{ids.Count}.");
                ecollector = new FilteredElementCollector(doc, ids);
            }

            CurrentElementList = new FilteredElementCollector(doc, ids).WherePasses(new LogicalAndFilter(
                    new ElementClassFilter(typeof(FamilyInstance)),
                    new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels)).Where(x => (x as FamilyInstance).Symbol.Name != "NULL").ToList();

            uidoc.Selection.Elements.Clear();
            CurrentElementList.ForEach(ele =>
            {
                if (ele.get_Parameter("分项")?.AsInteger() == CurrentSelectedType) uidoc.Selection.Elements.Add(ele);
            });

            listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - FILTERED: ELE/{CurrentElementList.Count}.");
            return;
        }

        private string sfname = "SF";
        private RoutedCommand cmdApplySelection = new RoutedCommand();
        private RoutedCommand cmdParamsRefresh = new RoutedCommand();
        private RoutedCommand cmdSelectParam = new RoutedCommand();
        public RoutedCommand CmdSelectParam { get { return cmdSelectParam; } set { cmdSelectParam = value; OnPropertyChanged(nameof(CmdSelectParam)); } }
        private void InitializeCommand()
        {
            CommandBinding cbSelectParam = new CommandBinding(cmdSelectParam, (sender, e) =>
            {
                if((e.OriginalSource as CheckBox).IsChecked == true)
                    ParamListFiltered.Add(e.Parameter.ToString());
                else
                    ParamListFiltered.Remove(e.Parameter.ToString());

            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });


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

            //bnApplySelection.Command = cmdApplySelection;
            //bnParamsRefresh.Command = cmdParamsRefresh;

            ProcFCM.CommandBindings.AddRange(new CommandBinding[]
            {
                cbSelectParam,
                cbParamsRefresh,
                cbApplySelection
            });
        }

        private void listInformation_SelectionChanged(object sender, SelectionChangedEventArgs e) { var lb = sender as ListBox; lb.ScrollIntoView(lb.Items[lb.Items.Count - 1]); }

        private void chkbox_Checked(object sender, RoutedEventArgs e)
        {
            sfname = "SF-";


            ((CheckBox)sender).GetBindingExpression(CheckBox.IsCheckedProperty).UpdateTarget();
        }

    }
}
