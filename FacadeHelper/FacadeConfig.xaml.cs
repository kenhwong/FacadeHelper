using System;
using System.Collections.Generic;
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
    /// Interaction logic for FacadeConfig.xaml
    /// </summary>
    public partial class FacadeConfig : UserControl, INotifyPropertyChanged
    {
        public Window ParentWin { get; set; }

        private bool _isSuspend = false;
        public bool IsSuspend { get { return _isSuspend; } set { _isSuspend = value; OnPropertyChanged(nameof(IsSuspend)); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        public FacadeConfig()
        {
            InitializeComponent();
            InitializeCommand();
        }

        private RoutedCommand cmdIDEdit = new RoutedCommand();
        private RoutedCommand cmdIDEditOk = new RoutedCommand();
        private RoutedCommand cmdIDEditCancel = new RoutedCommand();
        private RoutedCommand cmdPNEdit = new RoutedCommand();
        private RoutedCommand cmdPNEditOk = new RoutedCommand();
        private RoutedCommand cmdPNEditCancel = new RoutedCommand();
        private RoutedCommand cmdBrowseZoneData = new RoutedCommand();

        private void InitializeCommand()
        {
            #region CommandBinding : IDEdit
            CommandBinding cbIDEdit = new CommandBinding(cmdIDEdit,
        (sender, e) =>
        {
            txtProjectID.IsReadOnly = false;
            bnIDEditOk.IsEnabled = true;
            bnIDEditCancel.IsEnabled = true;
            bnIDEdit.IsEnabled = false;
            IsSuspend = true;
        },
        (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbIDEditOk = new CommandBinding(cmdIDEditOk,
                (sender, e) =>
                {
                    if (System.Text.RegularExpressions.Regex.IsMatch(txtProjectID.Text, @"^BIM\d{10}[C|W]$"))
                    {
                        Global.UpdateAppConfig("CurrentProjectID", txtProjectID.Text);

                        txtProjectID.IsReadOnly = true;
                        bnIDEditCancel.IsEnabled = false;
                        bnIDEdit.IsEnabled = true;
                        bnIDEditOk.IsEnabled = false;
                        IsSuspend = false;
                    }
                    else
                    {
                        txtProjectID.Focus();
                        txtProjectID.SelectAll();
                        return;
                    }
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbIDEditCancel = new CommandBinding(cmdIDEditCancel,
                (sender, e) =>
                {
                    txtProjectID.IsReadOnly = true;
                    bnIDEditCancel.IsEnabled = false;
                    bnIDEdit.IsEnabled = true;
                    bnIDEditOk.IsEnabled = false;
                    IsSuspend = false;
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            #endregion

            CommandBinding cbPNEdit = new CommandBinding(cmdPNEdit,
                (sender, e) =>
                {
                    txtProjectName.IsReadOnly = false;
                    bnPNEditOk.IsEnabled = true;
                    bnPNEditCancel.IsEnabled = true;
                    bnPNEdit.IsEnabled = false;
                    IsSuspend = true;
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbPNEditOk = new CommandBinding(cmdPNEditOk,
                (sender, e) =>
                {
                    if (System.Text.RegularExpressions.Regex.IsMatch(txtProjectName.Text, @"^\S.*?\S$"))
                    {
                        Global.UpdateAppConfig("CurrentProjectName", txtProjectName.Text);

                        txtProjectName.IsReadOnly = true;
                        bnPNEditCancel.IsEnabled = false;
                        bnPNEdit.IsEnabled = true;
                        bnPNEditOk.IsEnabled = false;
                        IsSuspend = false;
                    }
                    else
                    {
                        txtProjectName.Focus();
                        txtProjectName.SelectAll();
                        return;
                    }
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });
            CommandBinding cbPNEditCancel = new CommandBinding(cmdPNEditCancel,
                (sender, e) =>
                {
                    txtProjectName.IsReadOnly = true;
                    bnPNEditCancel.IsEnabled = false;
                    bnPNEdit.IsEnabled = true;
                    bnPNEditOk.IsEnabled = false;
                    IsSuspend = false;
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            CommandBinding cbBrowseZoneData = new CommandBinding(cmdBrowseZoneData,
                (sender, e) =>
                {
                    Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog()
                    {
                        InitialDirectory = System.IO.Path.GetDirectoryName(Global.DataFile),
                        DefaultExt = "*.zone",
                        Filter = "Zone Data Files(*.zone)|*.zone|All(*.*)|*.*"
                    };
                    ZoneHelper.FnLoadZoneScheduleData(ofd.FileName);
                    listInformation.SelectedIndex = listInformation.Items.Add($"{DateTime.Now:HH:mm:ss} - L: {ofd.FileName}.");
                },
                (sender, e) => { e.CanExecute = true; e.Handled = true; });

            bnIDEdit.Command = cmdIDEdit;
            bnIDEditOk.Command = cmdIDEditOk;
            bnIDEditCancel.Command = cmdIDEditCancel;
            bnPNEdit.Command = cmdPNEdit;
            bnPNEditOk.Command = cmdPNEditOk;
            bnPNEditCancel.Command = cmdPNEditCancel;

            bnBrowseZoneData.Command = cmdBrowseZoneData;

            ProcConfig.CommandBindings.AddRange(new CommandBinding[]
            {
                cbIDEdit,
                cbIDEditOk,
                cbIDEditCancel,
                cbPNEdit,
                cbPNEditOk,
                cbPNEditCancel,
                cbBrowseZoneData
            });


        }

        private void listInformation_SelectionChanged(object sender, SelectionChangedEventArgs e) { var lb = sender as ListBox; lb.ScrollIntoView(lb.Items[lb.Items.Count - 1]); }
    }
}
