using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
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
    /// Interaction logic for FamilyHelper.xaml
    /// </summary>
    public partial class FamilyHelper : UserControl, INotifyPropertyChanged
    {
        public Window ParentWin { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private UIApplication uiapp;
        private UIDocument uidoc;
        private Document doc;
        private ExternalCommandData cdata;

        private DefinitionFile _currentSharedParamFile;
        public DefinitionFile CurrentSharedParamFile { get { return _currentSharedParamFile; } set { _currentSharedParamFile = value; OnPropertyChanged(nameof(CurrentSharedParamFile)); } }
        private List<DefinitionGroup> _currentDefinitionGroups;
        public List<DefinitionGroup> CurrentDefinitionGroups { get { return _currentDefinitionGroups; } set { _currentDefinitionGroups = value; OnPropertyChanged(nameof(CurrentDefinitionGroups)); } }


        public FamilyHelper(ExternalCommandData commandData)
        {
            InitializeComponent();
            this.DataContext = this;
            InitializeCommand();

            cdata = commandData;
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;

            FamilyManager fm = doc.FamilyManager;
            CurrentSharedParamFile = uiapp.Application.OpenSharedParameterFile();
            CurrentDefinitionGroups = CurrentSharedParamFile.Groups.ToList();
            #region 铝型材纠错
            var p1 = fm.get_Parameter("米重");
            var p2 = fm.get_Parameter("模图编号");

            double? v1 = 0;
            string v2 = "";
            if (p1 != null && p2 != null)
                if (!p1.IsShared && !p2.IsShared)
                {
                    using (Transaction trans = new Transaction(doc, "CreateFamilyParameters"))
                    {
                        trans.Start();
                        foreach (FamilyType ft in fm.Types)
                        {
                            if (ft.HasValue(p1)) v1 = ft.AsDouble(p1);
                            if (ft.HasValue(p2)) v2 = ft.AsString(p2);
                        }
                        doc.Delete(p1.Id);
                        doc.Delete(p2.Id);

                        DefinitionFile spfile = uiapp.Application.OpenSharedParameterFile();
                        DefinitionGroup _grp = spfile.Groups.get_Item("物理特性");
                        ExternalDefinition _def1 = _grp.Definitions.get_Item("米重") as ExternalDefinition;
                        ExternalDefinition _def2 = _grp.Definitions.get_Item("模图编号") as ExternalDefinition;
                        //添加参数
                        FamilyManager familyMgr = doc.FamilyManager;
                        bool isInstance = false;
                        FamilyParameter paramMID = familyMgr.AddParameter(_def2, BuiltInParameterGroup.PG_TEXT, isInstance);
                        FamilyParameter paramWPM = familyMgr.AddParameter(_def1, BuiltInParameterGroup.INVALID, isInstance);

                        familyMgr.Set(paramMID, v2);
                        familyMgr.Set(paramWPM, (double)v1);

                        trans.Commit();

                        txtMID.Text = v2;
                        txtWPM.Text = v1.ToString();
                    }

                }

            #endregion

        }

        private RoutedCommand cmdApplyParam = new RoutedCommand();
        private void InitializeCommand()
        {
            CommandBinding cbApplyParam = new CommandBinding(cmdApplyParam, (sender, e) =>
            {
                using (Transaction trans = new Transaction(doc, "CreateFamilyParameters"))
                {
                    trans.Start();
                    DefinitionFile spfile = uiapp.Application.OpenSharedParameterFile();
                    DefinitionGroup _grp = spfile.Groups.get_Item("物理特性");
                    ExternalDefinition _def1 = _grp.Definitions.get_Item("米重") as ExternalDefinition;
                    ExternalDefinition _def2 = _grp.Definitions.get_Item("模图编号") as ExternalDefinition;
                    //添加参数
                    FamilyManager familyMgr = doc.FamilyManager;
                    bool isInstance = false;
                    FamilyParameter paramMID = familyMgr.AddParameter(_def2, BuiltInParameterGroup.PG_TEXT, isInstance);
                    FamilyParameter paramWPM = familyMgr.AddParameter(_def1, BuiltInParameterGroup.INVALID, isInstance);

                    familyMgr.Set(paramMID, txtMID.Text);
                    familyMgr.Set(paramWPM, double.Parse(txtWPM.Text));

                    trans.Commit();
                }

                lblApplied.Visibility = System.Windows.Visibility.Visible;
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            bnApplyParam.Command = cmdApplyParam;

            ProcFH.CommandBindings.AddRange(new CommandBinding[]
            {
                cbApplyParam
            });
        }
    }
}
