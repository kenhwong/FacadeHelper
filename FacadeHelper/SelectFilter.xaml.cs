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
    public partial class SelectFilter : UserControl, INotifyPropertyChanged
    {
        public Window ParentWin { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private bool _isFilterWall = false;
        private bool _isFilterCurtainSystem = false;
        private bool _isFilterCurtainGrid = false;
        private bool _isFilterCurtainPanel = true;
        private bool _isFilterGenericModel = false;
        private bool _isFilterCurtainWallMullion = false;
        public bool IsFilterWall { get { return _isFilterWall; } set { _isFilterWall = value; OnPropertyChanged(nameof(IsFilterWall)); } }
        public bool IsFilterCurtainSystem { get { return _isFilterCurtainSystem; } set { _isFilterCurtainSystem = value; OnPropertyChanged(nameof(IsFilterCurtainSystem)); } }
        public bool IsFilterCurtainGrid { get { return _isFilterCurtainGrid; } set { _isFilterCurtainGrid = value; OnPropertyChanged(nameof(IsFilterCurtainGrid)); } }
        public bool IsFilterCurtainPanel { get { return _isFilterCurtainPanel; } set { _isFilterCurtainPanel = value; OnPropertyChanged(nameof(IsFilterCurtainPanel)); } }
        public bool IsFilterGenericModel { get { return _isFilterGenericModel; } set { _isFilterGenericModel = value; OnPropertyChanged(nameof(IsFilterGenericModel)); } }
        public bool IsFilterCurtainWallMullion { get { return _isFilterCurtainWallMullion; } set { _isFilterCurtainWallMullion = value; OnPropertyChanged(nameof(IsFilterCurtainWallMullion)); } }

        private ObservableCollection<string> _paramListSource = new ObservableCollection<string>();
        private ObservableCollection<string> _paramListFiltered = new ObservableCollection<string>();
        public ObservableCollection<string> ParamListSource { get { return _paramListSource; } set { _paramListSource = value; OnPropertyChanged(nameof(ParamListSource)); } }
        public ObservableCollection<string> ParamListFiltered { get { return _paramListFiltered; } set { _paramListFiltered = value; OnPropertyChanged(nameof(ParamListFiltered)); } }


        private UIApplication uiapp;
        private UIDocument uidoc;
        private Document doc;
        private ExternalCommandData cdata;

        private List<Element> CurrentElementList;

        public SelectFilter(ExternalCommandData commandData)
        {
            InitializeComponent();
            this.DataContext = this;

            cdata = commandData;
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;

            InitializeCommand();
        }

        private void ProcessSelection()
        {
            FilteredElementCollector ecollector;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
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
            LogicalAndFilter _InstancesFilterWA = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_Walls));
            LogicalAndFilter _InstancesFilterCS = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_Curtain_Systems));
            LogicalAndFilter _InstancesFilterCG = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_CurtainGrids));
            LogicalAndFilter _InstancesFilterCP = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallPanels));
            LogicalAndFilter _InstancesFilterGM = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_GenericModel));
            LogicalAndFilter _InstancesFilterCM = new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_CurtainWallMullions));

            FilteredElementCollector fec = null;
            if (IsFilterWall) fec = ecollector.WherePasses(_InstancesFilterWA);
            if (IsFilterCurtainSystem) fec = ecollector.WherePasses(_InstancesFilterCS);
            if (IsFilterCurtainGrid) fec = ecollector.WherePasses(_InstancesFilterCG);
            if (IsFilterCurtainPanel) fec = ecollector.WherePasses(_InstancesFilterCP);
            if (IsFilterGenericModel) fec = ecollector.WherePasses(_InstancesFilterGM);
            if (IsFilterCurtainWallMullion) fec = ecollector.WherePasses(_InstancesFilterCM);

            CurrentElementList = fec.Where(x => (x as FamilyInstance).Symbol.Name != "NULL").ToList();

            uidoc.Selection.Elements.Clear();
            CurrentElementList.ForEach(ele => uidoc.Selection.Elements.Add(ele));

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

                CurrentElementList.ForEach(ele =>
                {
                    ParameterSet parameters = ele.Parameters;
                    foreach (Parameter parameter in parameters) ParamListSource.Add(parameter.Definition.Name);
                });

                HashSet<string> hs = new HashSet<string>(ParamListSource);
                ParamListSource.Clear();
                hs.ToList().ForEach(d => ParamListSource.Add(d));

            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbApplySelection = new CommandBinding(cmdApplySelection, (sender, e) =>
            {
                ProcessSelection();

                using (Transaction trans = new Transaction(doc, "CreateSelectionFilter"))
                {
                    trans.Start();
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    ICollection<Element> typecollection = collector.OfClass(typeof(SelectionFilterElement)).ToElements();
                    SelectionFilterElement selectset = typecollection.Cast<SelectionFilterElement>().FirstOrDefault(ele => ele.Name == txtSelectFilterName.Text);
                    if (selectset != null) selectset.Clear();
                    else selectset = SelectionFilterElement.Create(doc, txtSelectFilterName.Text);
                    CurrentElementList.ForEach(ele => selectset.AddSingle(ele.Id));
                    //doc.ActiveView.IsolateElementsTemporary(selectset.GetElementIds());
                    trans.Commit();
                }

            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            bnApplySelection.Command = cmdApplySelection;
            bnParamsRefresh.Command = cmdParamsRefresh;

            ProcSF.CommandBindings.AddRange(new CommandBinding[]
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
            if (IsFilterWall) sfname += "-W";
            if (IsFilterCurtainSystem) sfname += "-CS";
            if (IsFilterCurtainGrid) sfname += "-CG";
            if (IsFilterCurtainPanel) sfname += "-CP";
            if (IsFilterGenericModel) sfname += "-GM";
            if (IsFilterCurtainWallMullion) sfname += "-CM";

            txtSelectFilterName.Text = sfname;

            ((CheckBox)sender).GetBindingExpression(CheckBox.IsCheckedProperty).UpdateTarget();
        }

    }
}
