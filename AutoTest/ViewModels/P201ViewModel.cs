using AutoTest.Logic;
using FrameWork.Models;
using FrameWork.ViewModels.Base;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Input;

namespace AutoTest.ViewModels
{
    /// <summary>
    /// EXE测试
    /// </summary>
    public class P201ViewModel : BaseViewModel
    {
        /// <summary>
        /// Logic
        /// </summary>
        private P201Logic _logic = new P201Logic();

        /// <summary>
        /// Exel脚本Model
        /// </summary>
        public SeleniumScriptModel ExcelModel { get; set; }

        /// <summary>
        /// Exel脚本Model
        /// </summary>
        public bool FlgStop { get; set; } = false;

        /// <summary>
        /// 文件 下拉框
        /// </summary>
        public ObservableCollection<FileInfo> Files { get; set; } = new ObservableCollection<FileInfo>();
        public FileInfo SelectedFile { get; set; }

        /// <summary>
        /// 文件按钮
        /// </summary>
        public ICommand BtnReadFile { get; set; }

        /// <summary>
        /// 日期控件
        /// </summary>
        public string Arg1 { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");

        /// <summary>
        /// 开始按钮
        /// </summary>
        public ICommand BtnStart { get; set; }

        /// <summary>
        /// Events 下拉框
        /// </summary>
        public ObservableCollection<string> WebElements { get; set; }
        public string SelectedElement { get; set; }

        /// <summary>
        /// 继续按钮
        /// </summary>
        public ICommand BtnContinu { get; set; }

        /// <summary>
        /// 出力路径
        /// </summary>
        public string OutPath { get; set; }

        /// <summary>
        /// 制御
        /// </summary>
        public bool FlgFile { get; set; }
        public bool FlgDate { get; set; }
        public bool FlgStart { get; set; }
        public bool FlgContinue { get; set; }
        public Visibility FlgDoing { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Msg { get; set; }

        /// <summary>
        /// 强制终了按钮
        /// </summary>
        public ICommand BtnStop { get; set; }

        /// <summary>
        /// 初始化
        /// </summary>
        public P201ViewModel()
        {
            // 初始化
            this._logic.Init(this);
            // 按钮事件
            this.BtnReadFile = new RelayTCommand<P201ViewModel>(_logic.ReadFile);
            this.BtnStart = new RelayTCommand<P201ViewModel>(_logic.Start);
            this.BtnContinu = new RelayTCommand<P201ViewModel>(_logic.BtnContinu);
            this.BtnStop = new RelayTCommand<P201ViewModel>(_logic.Stop);
        }
    }
}
