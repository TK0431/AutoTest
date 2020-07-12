using AutoTest.Pages;
using FrameWork.Models;
using FrameWork.Utility;
using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
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

        public static void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        { 
            if(e.Exception.Message == "【停止】")
            { 
                Task.Factory.StartNew(() => MessageQueue.Enqueue($"已手动停止", "OK", delegate () { }));
                LogUtility.WriteInfo($"EXE测试：手动停止");
            }
            else if (e.Exception.Message == "Stop")
            {
            }
            else
            { 
                Task.Factory.StartNew(() => MessageQueue.Enqueue($"程序异常,找俞解决", "OK", delegate () { }));
                LogUtility.WriteError($"系统异常发生",e.Exception);
            }

            e.Handled = true;
        }
    }
}
