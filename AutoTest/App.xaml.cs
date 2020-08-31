using AutoTest.Pages;
using FrameWork.Consts;
using FrameWork.Utility;
using MaterialDesignThemes.Wpf;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

namespace AutoTest
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// 消息机制
        /// </summary>
        public static SnackbarMessageQueue MessageQueue { get; set; }

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            bool Exist;
            System.Threading.Mutex newMutex = new System.Threading.Mutex(true, "AutoTest", out Exist);
            if (Exist)//如果没有运行
            {
                this.DispatcherUnhandledException += new DispatcherUnhandledExceptionEventHandler(App_DispatcherUnhandledException);
                AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
                // 启动主画面
                W000 win = new W000();
                this.MainWindow = win;
                win.Show();
            }
            else
            {
                MessageBox.Show("程序已经打开，请不要重复启动", "提示");//弹出提示信息
                Environment.Exit(0);
            }
        }

        public static void ShowMessage(string message, string OK = "OK", EnumMessageType type = EnumMessageType.Info, Exception ex = null)
        {
            Task.Factory.StartNew(() => MessageQueue.Enqueue(message, OK, delegate (object x) { }, null, false, true, TimeSpan.FromMilliseconds(10000)));

            switch (type)
            {
                case EnumMessageType.Info:
                    LogUtility.WriteInfo(message);
                    break;
                case EnumMessageType.Warn:
                    LogUtility.WarnInfo(message);
                    break;
                case EnumMessageType.Error:
                    LogUtility.WriteError(message, ex);
                    break;
                default:
                    break;
            }
        }

        private static void ShowOKMessage()
        {
            Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"读取完毕", "OK", delegate (object x) { }, null, false, true, TimeSpan.FromMilliseconds(10000)));
        }

        public static void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            if (e.Exception.Message.StartsWith("Err:"))
            {
                ShowMessage(e.Exception.Message, "异常");
            }
            if (e.Exception.Message == "【停止】")
            {
                ShowMessage($"已手动停止", "OK");
            }
            else if (e.Exception.Message == "Stop")
            {
            }
            else
            {
                ShowMessage($"程序异常", "OK", EnumMessageType.Error, e.Exception);
            }

            e.Handled = true;

            Win_top();
        }

        /// <summary>
        /// 非UI线程抛出全局异常事件处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                Exception exception = e.ExceptionObject as Exception;
                if (exception != null)
                {
                    ShowMessage($"程序异常", "OK", EnumMessageType.Error, exception);
                }
            }
            catch (Exception ex)
            {
                ShowMessage($"程序异常", "OK", EnumMessageType.Error, ex);
            }
            Win_top();
        }

        public static void Win_top()
        {
            IntPtr hwnd = new WindowInteropHelper(Application.Current.MainWindow).Handle;
            User32Utility.SetForegroundWindow(hwnd);
        }
    }
}
