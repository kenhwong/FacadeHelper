using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
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

        public FamilyHelper(ExternalCommandData commandData)
        {
            InitializeComponent();
            this.DataContext = this;
            InitializeCommand();

            cdata = commandData;
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;
        }

        private RoutedCommand cmdApplyParam = new RoutedCommand();
        private void InitializeCommand()
        {
            CommandBinding cbApplyParam = new CommandBinding(cmdApplyParam, (sender, e) =>
            {
                using (Transaction trans = new Transaction(doc, "CreateFamilyParameters"))
                {
                    trans.Start();
                    //添加参数
                    FamilyManager familyMgr = doc.FamilyManager;
                    bool isInstance = false;
                    FamilyParameter paramMID = familyMgr.AddParameter("模图编号", BuiltInParameterGroup.PG_TEXT, ParameterType.Text, isInstance);
                    FamilyParameter paramWPM = familyMgr.AddParameter("米重", BuiltInParameterGroup.INVALID, ParameterType.Number, isInstance);

                    familyMgr.Set(paramMID, txtMID.Text);
                    familyMgr.Set(paramWPM, double.Parse(txtWPM.Text));

                    trans.Commit();
                }

                lblApplied.Visibility = System.Windows.Visibility.Visible;
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });

            bnApplyParam.Command = cmdApplyParam;

            ProcPH.CommandBindings.AddRange(new CommandBinding[]
            {
                cbApplyParam
            });
        }
    }
}
