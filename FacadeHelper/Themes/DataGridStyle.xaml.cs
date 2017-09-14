using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace FacadeHelper
{
    partial class DataGridStyle : ResourceDictionary
    {
        private void DataGridHostItem_GotFocus(object sender, RoutedEventArgs e)
        {
            DataGridRow hostrow = ((DataGridRow)sender);
            if (hostrow.Item is CurtainPanelInfo)
            {
                CurtainPanelInfo currentcs = hostrow.Item as CurtainPanelInfo;
                
            }
        }
    }
}
