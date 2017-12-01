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
    /// Interaction logic for FacadeConfig.xaml
    /// </summary>
    public partial class FacadeConfig : UserControl, INotifyPropertyChanged
    {
        public Window ParentWin { get; set; }
        private bool isModified = false;

        private bool _isSuspend = false;
        public bool IsSuspend { get { return _isSuspend; } set { _isSuspend = value; OnPropertyChanged(nameof(IsSuspend)); } }

        public List<ElementClass> CurrentElementClassList { get; set; } = new List<ElementClass>();
        public List<ZoneLayerInfo> CurrentZoneLayerList { get; set; } = new List<ZoneLayerInfo>();

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private string currpid = string.Empty;
        private string zfile = string.Empty;
        private string ecfile = string.Empty;

        public FacadeConfig()
        {
            InitializeComponent();
            InitializeCommand();
            this.DataContext = this;
            if (Global.GetAppConfig("CurrentProjectID") is null)
            {
                CurrentElementClassList = ZoneHelper.InitElementClass();
                return;
            }
            currpid = Global.GetAppConfig("CurrentProjectID");
            txtProjectID.Text = currpid;
            //txtProjectName.Text = Global.GetAppConfig("CurrentProjectName") ?? "未设置"; //针对COCC项目测试方便，该项值用预设
            txtRVTPrecision.Text = Global.GetAppConfig("RVTPrecision") ?? "0.03";
            zfile = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), $"{currpid}.zone.xml");
            ecfile = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), $"{currpid}.class.xml");

            if ((CurrentZoneLayerList = ZoneHelper.FnZoneDataDeserialize()) is null)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: ZONE LAYER/NOT SET.");
            }
            if ((CurrentElementClassList = ZoneHelper.FnFilterClassDeserialize()) is null)
            {
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - ERR: ELEMENT CLASS/NOT SET.");
                CurrentElementClassList = ZoneHelper.InitElementClass();
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - INIT: ELEMENT CLASS.");
            }
        }

        private RoutedCommand cmdIDEdit = new RoutedCommand();
        private RoutedCommand cmdIDEditOk = new RoutedCommand();
        private RoutedCommand cmdIDEditCancel = new RoutedCommand();
        private RoutedCommand cmdPNEdit = new RoutedCommand();
        private RoutedCommand cmdPNEditOk = new RoutedCommand();
        private RoutedCommand cmdPNEditCancel = new RoutedCommand();
        private RoutedCommand cmdBrowseZoneData = new RoutedCommand();

        private RoutedCommand cmdApplyConfigure = new RoutedCommand();

        private void InitializeCommand()
        {
            #region CommandBinding : IDEdit
            CommandBinding cbIDEdit = new CommandBinding(cmdIDEdit, (sender, e) =>
            {
                txtProjectID.IsReadOnly = false;
                bnIDEditOk.IsEnabled = true;
                bnIDEditCancel.IsEnabled = true;
                bnIDEdit.IsEnabled = false;
                IsSuspend = true;
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbIDEditOk = new CommandBinding(cmdIDEditOk, (sender, e) =>
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(txtProjectID.Text, @"^BIM\d{10}[C|W]$"))
                {
                    Global.UpdateAppConfig("CurrentProjectID", txtProjectID.Text);

                    txtProjectID.IsReadOnly = true;
                    bnIDEditCancel.IsEnabled = false;
                    bnIDEdit.IsEnabled = true;
                    bnIDEditOk.IsEnabled = false;
                    IsSuspend = false;

                    lblApplied.Visibility = Visibility.Collapsed;
                    isModified = true;
                }
                else
                {
                    txtProjectID.Focus();
                    txtProjectID.SelectAll();
                    return;
                }
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbIDEditCancel = new CommandBinding(cmdIDEditCancel, (sender, e) =>
            {
                txtProjectID.IsReadOnly = true;
                bnIDEditCancel.IsEnabled = false;
                bnIDEdit.IsEnabled = true;
                bnIDEditOk.IsEnabled = false;
                IsSuspend = false;
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            #endregion

            #region CommandBinding : PNEdit
            CommandBinding cbPNEdit = new CommandBinding(cmdPNEdit, (sender, e) =>
            {
                txtProjectName.IsReadOnly = false;
                bnPNEditOk.IsEnabled = true;
                bnPNEditCancel.IsEnabled = true;
                bnPNEdit.IsEnabled = false;
                IsSuspend = true;
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbPNEditOk = new CommandBinding(cmdPNEditOk, (sender, e) =>
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(txtProjectName.Text, @"^\S.*?\S$"))
                {
                    Global.UpdateAppConfig("CurrentProjectName", txtProjectName.Text);

                    txtProjectName.IsReadOnly = true;
                    bnPNEditCancel.IsEnabled = false;
                    bnPNEdit.IsEnabled = true;
                    bnPNEditOk.IsEnabled = false;
                    IsSuspend = false;

                    lblApplied.Visibility = Visibility.Collapsed;
                    isModified = true;
                }
                else
                {
                    txtProjectName.Focus();
                    txtProjectName.SelectAll();
                    return;
                }
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbPNEditCancel = new CommandBinding(cmdPNEditCancel, (sender, e) =>
            {
                txtProjectName.IsReadOnly = true;
                bnPNEditCancel.IsEnabled = false;
                bnPNEdit.IsEnabled = true;
                bnPNEditOk.IsEnabled = false;
                IsSuspend = false;
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            #endregion

            #region BrowseZoneData
            CommandBinding cbBrowseZoneData = new CommandBinding(cmdBrowseZoneData, (sender, e) =>
            {
                Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog()
                {
                    DefaultExt = "*.zone",
                    Filter = "Zone Data Files(*.zone)|*.zone|All(*.*)|*.*"
                };
                if (ofd.ShowDialog() == false) return;
                if (Global.GetAppConfig("CurrentProjectID") is null)
                {
                    MessageBox.Show("导入分区数据之前必须先设置当前项目的项目编号。");
                    return;
                }
                var pzfile = ZoneHelper.FnZoneDataSerialize(ofd.FileName);
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - L: ZONEDATA/{ofd.FileName}.");
                listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - S: ZONEDATA/{pzfile}.");
                txtProjectZoneFile.Text = pzfile;
            }, (sender, e) => { e.CanExecute = true; e.Handled = true; });
            #endregion

            CommandBinding cbApplyConfigure = new CommandBinding(cmdApplyConfigure, (sender, e) =>
            {
                ZoneHelper.FnFilterClassSerialize(CurrentElementClassList);
                Global.UpdateAppConfig("CurrentProjectID", txtProjectID.Text);
                Global.UpdateAppConfig("CurrentProjectName", txtProjectName.Text);
                Global.UpdateAppConfig("RVTPrecision", txtRVTPrecision.Text);
                lblApplied.Visibility = Visibility.Visible;
                isModified = false;
            }, (sender, e) => { if (isModified) { e.CanExecute = true; e.Handled = true; } });

            bnIDEdit.Command = cmdIDEdit;
            bnIDEditOk.Command = cmdIDEditOk;
            bnIDEditCancel.Command = cmdIDEditCancel;
            bnPNEdit.Command = cmdPNEdit;
            bnPNEditOk.Command = cmdPNEditOk;
            bnPNEditCancel.Command = cmdPNEditCancel;

            bnBrowseZoneData.Command = cmdBrowseZoneData;
            bnApplyConfigure.Command = cmdApplyConfigure;

            ProcConfig.CommandBindings.AddRange(new CommandBinding[]
            {
                cbIDEdit,
                cbIDEditOk,
                cbIDEditCancel,
                cbPNEdit,
                cbPNEditOk,
                cbPNEditCancel,
                cbBrowseZoneData,
                cbApplyConfigure
            });


        }

        private void listInformation_SelectionChanged(object sender, SelectionChangedEventArgs e) { var lb = sender as ListBox; lb.ScrollIntoView(lb.Items[lb.Items.Count - 1]); }

        private void txtRVTPrecision_TextChanged(object sender, TextChangedEventArgs e)
        {
            lblApplied.Visibility = Visibility.Collapsed;
            isModified = true;
        }
    }
}
