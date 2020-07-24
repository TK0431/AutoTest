using AutoTest.Consts;
using AutoTest.ViewModels;
using FrameWork.Models;
using FrameWork.Utility;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace AutoTest.Pages
{
    /// <summary>
    /// 主画面
    /// </summary>
    public partial class W000 : Window
    {
        /// <summary>
        /// Model
        /// </summary>
        private W000ViewModel _model;

        /// <summary>
        /// 初期化
        /// </summary>
        public W000()
        {
            InitializeComponent();

            _model = new W000ViewModel();
            this.DataContext = _model;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            App.MessageQueue = SnackbarThree.MessageQueue;
        }

        /// <summary>
        /// 关闭事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ProcessUtility.KillProcess("chromedriver");

            System.Environment.Exit(0);
        }

        /// <summary>
        /// 检索Page
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Search_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox tb = sender as TextBox;
            List<EnumItem> list = typeof(PageEnum).GetList();
            _model.MenuItems = new ObservableCollection<EnumItem>(list.Where(x => x.Description.Contains(tb.Text)));
        }

        /// <summary>
        /// 切换Page
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuList_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            EnumItem selected = ((ListBox)sender).SelectedItem as EnumItem;

            if (selected != null)
            {
                switch ((PageEnum)selected.Index)
                {
                    case PageEnum.P001:
                        _model.MainPage = new P001();
                        break;
                    case PageEnum.P101:
                        _model.MainPage = new P101();
                        break;
                    case PageEnum.P102:
                        _model.MainPage = new P102();
                        break;
                    case PageEnum.P103:
                        _model.MainPage = new P103();
                        break;
                    case PageEnum.P201:
                        _model.MainPage = new P201();
                        break;
                    case PageEnum.P202:
                        _model.MainPage = new P202();
                        break;
                    default:
                        break;
                }
            }

            MenuToggleButton.IsChecked = false;
        }

        /// <summary>
        /// 关闭窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 最小化
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Mini_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }
    }
}
