using AutoTest.ViewModels;
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
    /// P201.xaml の相互作用ロジック
    /// </summary>
    public partial class P201 : Page
    {
        private P201ViewModel _model;
        public P201()
        {
            InitializeComponent();

            _model = new P201ViewModel();
            this.DataContext = _model;
        }
    }
}
