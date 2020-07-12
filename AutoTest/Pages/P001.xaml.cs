using AutoTest.ViewModels;
using System.Windows.Controls;

namespace AutoTest.Pages
{
    /// <summary>
    /// EXE解析
    /// </summary>
    public partial class P001 : Page
    {
        /// <summary>
        /// Model
        /// </summary>
        private P001ViewModel _model;

        public P001()
        {
            InitializeComponent();

            _model = new P001ViewModel();
            this.DataContext = _model;
        }
    }
}
