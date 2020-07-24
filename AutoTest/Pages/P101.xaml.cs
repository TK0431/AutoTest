using AutoTest.ViewModels;
using FrameWork.Consts;
using FrameWork.Models;
using FrameWork.Utility;
using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;

namespace AutoTest.Pages
{
    /// <summary>
    /// EXE解析
    /// </summary>
    public partial class P101 : Page
    {
        /// <summary>
        /// Model
        /// </summary>
        private P101ViewModel _model;
        MouseHook mh;

        public P101()
        {
            InitializeComponent();

            _model = new P101ViewModel();
            this.DataContext = _model;

            mh = new MouseHook();

            mh.MouseMoveEvent += mh_MouseMoveEvent;
            mh.MouseClickEvent += mh_MouseClickEvent;
        }

        private void TreeView_SelectedItemChanged(object sender, System.Windows.RoutedPropertyChangedEventArgs<object> e)
        {
            System.Windows.Controls.TreeView view = sender as System.Windows.Controls.TreeView;
            HwndItem selected = view.SelectedItem as HwndItem;

            if (selected == null) return;

            _model.Hwnd = selected.TreeTextHwnd;
            _model.Value = selected.HModel.Value;
            _model.Class = selected.HModel.Class;
            _model.PointX = selected.HModel.ExeX;
            _model.PointY = selected.HModel.ExeY;
            _model.Width = selected.HModel.Width;
            _model.Height = selected.HModel.Height;
            _model.SelectedHwndItem = selected;

            _model.TopHwnd = HwndUtility.GetTopParentHwnd(_model.SelectedHwndItem.HModel.HwndId).ToString();
        }

        HwndModel _fmodel = null;
        private void Point_Click(object sender, RoutedEventArgs e)
        {
            if (_model.SelectedHwndItem == null)
            {
                Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"请选择一个程序控件", "OK", delegate () { }));
                LogUtility.WriteError($"程序控件未选择", null);
                return;
            }

            IntPtr hwnd = HwndUtility.GetTopParentHwnd(_model.SelectedHwndItem.HModel.HwndId);
            _model.TopHwnd = HwndUtility.GetTopParentHwnd(_model.SelectedHwndItem.HModel.HwndId).ToString();
            _fmodel = HwndUtility.GetHwndModel(hwnd);

            _model.Hwnd = "";
            _model.Value = "";
            _model.Class = "";
            _model.PointX = 0;
            _model.PointY = 0;
            _model.Width = 0;
            _model.Height = 0;

            mh.SetHook();
        }

        private void mh_MouseClickEvent(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _model.PointX = e.X - _fmodel.DeskX;
                _model.PointY = e.Y - _fmodel.DeskY;
                _model.Class = "$CustomControl";

                btnAdd.IsEnabled = true;

                mh.UnHook();
            }
        }

        private void mh_MouseMoveEvent(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            _model.PointX = e.Location.X - _fmodel.DeskX;
            _model.PointY = e.Location.Y - _fmodel.DeskY;
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            _model.AddControl.Add(new HwndModel()
            {
                Value = _model.Value,
                Type = HwndItemType.Control,
                Class = "$CustomControl",
                ExeX = _model.PointX,
                ExeY = _model.PointY,
            });
            LogUtility.WriteInfo($"EXE解析:追加控件{_model.PointX}-{_model.PointY}");

            tbkAdd.Text = $"追加控件：{_model.AddControl.Count}件";

            btnAdd.IsEnabled = false;
            _fmodel = null;
        }

        private void Find_Click(object sender, RoutedEventArgs e)
        {
            User32Utility.SetForegroundWindow(_model.SelectedHwndItem.HModel.HwndId);
            Thread.Sleep(100);

            W001 range = new W001();

            //range.StartPosition = FormStartPosition.Manual;
            range.Left = _model.SelectedHwndItem.HModel.DeskX;
            range.Top = _model.SelectedHwndItem.HModel.DeskY;
            range.Show();
            range.Width = _model.SelectedHwndItem.HModel.Width;
            range.Height = _model.SelectedHwndItem.HModel.Height;
            range.Topmost = true;
        }

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            var eventArg = new MouseWheelEventArgs(e.MouseDevice, e.Timestamp, e.Delta);
            eventArg.RoutedEvent = UIElement.MouseWheelEvent;
            eventArg.Source = sender;
            scrollViewer.RaiseEvent(eventArg);
        }
    }
}
