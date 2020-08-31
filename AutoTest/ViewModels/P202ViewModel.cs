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
    /// Web测试
    /// </summary>
    public class P202ViewModel : BaseViewModel
    {
        /// <summary>
        /// Logic
        /// </summary>
        private P202Logic _logic = new P202Logic();

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
        /// 开始按钮
        /// </summary>
        public ICommand BtnStart { get; set; }

        public ICommand OutStart { get; set; }

        /// <summary>
        /// 出力路径
        /// </summary>
        public string OutPath { get; set; }

        public Visibility FlgDoing { get; set; }

        /// <summary>
        /// 强制终了按钮
        /// </summary>
        public ICommand BtnStop { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Msg { get; set; }

        /// <summary>
        /// 初始化
        /// </summary>
        public P202ViewModel()
        {
            // 初始化
            this._logic.Init(this);
            // 按钮事件
            this.BtnReadFile = new RelayTCommand<P202ViewModel>(_logic.ReadFile);
            this.BtnStart = new RelayTCommand<P202ViewModel>(_logic.Start);
            this.OutStart = new RelayTCommand<P202ViewModel>(_logic.CreateExel);
            this.BtnStop = new RelayTCommand<P202ViewModel>(_logic.Stop);
        }
    }
}
