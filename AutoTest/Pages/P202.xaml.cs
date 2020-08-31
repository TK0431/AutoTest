using AutoTest.ViewModels;
using FrameWork.Utility;
using System;
using System.Collections.Generic;
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

namespace AutoTest.Pages
{
    /// <summary>
    /// P202.xaml の相互作用ロジック
    /// </summary>
    public partial class P202 : Page
    {
        private P202ViewModel _model;

        public P202()
        {
            InitializeComponent();

            _model = new P202ViewModel();
            this.DataContext = _model;
        }

        private void Page_Unloaded(object sender, RoutedEventArgs e)
        {
            ProcessUtility.KillProcess("chromedriver");
        }
    }
}
