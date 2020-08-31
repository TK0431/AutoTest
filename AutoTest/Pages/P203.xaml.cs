using AutoTest.ViewModels;
using FrameWork.Utility;
using System;
using System.Collections.Generic;
using System.Globalization;
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
    public partial class P203 : Page
    {
        private P203ViewModel _model;

        public P203()
        {
            InitializeComponent();

            _model = new P203ViewModel();
            this.DataContext = _model;
        }

        private void Page_Unloaded(object sender, RoutedEventArgs e)
        {
            ProcessUtility.KillProcess("chromedriver");
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            ProcessUtility.KillProcess("chromedriver");
        }
    }

    public class StringToImageSourceConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string path = (string)value;
            return !string.IsNullOrEmpty(path) ? new BitmapImage(new Uri(path, UriKind.Absolute)) : null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }
}
