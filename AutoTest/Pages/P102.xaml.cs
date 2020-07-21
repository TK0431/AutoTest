using AutoTest.ViewModels;
using FrameWork.Utility;
using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;

namespace AutoTest.Pages
{
    /// <summary>
    /// EXE测试
    /// </summary>
    public partial class P102 : Page
    {
        private P102ViewModel _model;
        KeyboardHook k_hook;

        public P102()
        {
            InitializeComponent();

            _model = new P102ViewModel();
            this.DataContext = _model;
        }

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            // Hook
            k_hook = new KeyboardHook();
            k_hook.KeyDownEvent += new KeyEventHandler(hook_KeyDown);
            k_hook.Start();
        }

        /// <summary>
        /// 截图
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void hook_KeyDown(object sender, KeyEventArgs e)
        {
            // Control + Control
            if (e.KeyValue == (int)Keys.RControlKey && (int)System.Windows.Forms.Control.ModifierKeys == (int)Keys.Control)
            {
                // Alt + PS
                User32Utility.keybd_event((byte)Keys.Menu, 0, 0x0, IntPtr.Zero);
                User32Utility.keybd_event((byte)0x2c, 0, 0x0, IntPtr.Zero);
                System.Windows.Forms.Application.DoEvents();
                System.Windows.Forms.Application.DoEvents();
                User32Utility.keybd_event((byte)0x2c, 0, 0x2, IntPtr.Zero);
                User32Utility.keybd_event((byte)Keys.Menu, 0, 0x2, IntPtr.Zero);
                System.Windows.Forms.Application.DoEvents();
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);

                // Path Create
                if (!Directory.Exists(_model.PicPath))
                {
                    Directory.CreateDirectory(_model.PicPath);
                    LogUtility.WriteInfo($"创建路径{_model.PicPath}");
                }

                if (!Directory.Exists(_model.PicPath + @"\" + _model.CaseNo))
                {
                    Directory.CreateDirectory(_model.PicPath + @"\" + _model.CaseNo);
                    LogUtility.WriteInfo($"创建路径{_model.PicPath + @"\" + _model.CaseNo}");
                }

                // Pic save
                string picFullPath = _model.PicPath + @"\" + _model.CaseNo + @"\" + _model.PicNum2.ToString().PadLeft(3, '0') + ".png";
                System.Windows.Forms.IDataObject newObject = null;
                Bitmap newBitmap = null;
                System.Windows.Forms.Application.DoEvents();
                newObject = System.Windows.Forms.Clipboard.GetDataObject();
                if (System.Windows.Forms.Clipboard.ContainsImage())
                {
                    newBitmap = (Bitmap)(System.Windows.Forms.Clipboard.GetImage().Clone());
                    newBitmap.Save(picFullPath);
                }
                newBitmap.Dispose();

                // No + 1
                _model.PicNum2 = _model.PicNum2 + 1;

                LogUtility.WriteInfo($"图片保存成功{picFullPath}");
            }
        }
    }
}
