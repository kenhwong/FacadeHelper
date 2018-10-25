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
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace FacadeHelper
{
    /// <summary>
    /// Interaction logic for CodeFix.xaml
    /// </summary>
    public partial class CodeFix : UserControl
    {
        public Window ParentWin { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private readonly UIApplication _uiapp;
        private readonly UIDocument _uidoc;
        private Document _doc;
        private ICollection<ElementId> _sids;

        private ObservableCollection<Element> _currentCswCollection = new ObservableCollection<Element>();
        public ObservableCollection<Element> CurrentCswCollection { get => _currentCswCollection; set { _currentCswCollection = value; OnPropertyChanged(nameof(CurrentCswCollection)); } }


        public CodeFix(ExternalCommandData commandData)
        {
            InitializeComponent();
            this.DataContext = this;

            _uiapp = commandData.Application;
            _uidoc = _uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;
            _sids = _uidoc.Selection.Elements.Cast<Element>().Select(e => e.Id).ToList<ElementId>();

            //CS/W 统计
            var cswcollection = new FilteredElementCollector(_doc).WherePasses(new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_Walls)))
                .Union(new FilteredElementCollector(_doc).WherePasses(new LogicalAndFilter(new ElementClassFilter(typeof(FamilyInstance)), new ElementCategoryFilter(BuiltInCategory.OST_Curtain_Systems))));
            foreach (var csw in cswcollection) CurrentCswCollection.Add(csw);


            InitializeCommand();

        }

        private void InitializeCommand()
        {
            throw new NotImplementedException();
        }
    }
}
