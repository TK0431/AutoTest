using AutoTest.Logic;
using FrameWork.Models;
using FrameWork.ViewModels.Base;
using System;
using System.Collections.ObjectModel;
using System.IO;
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


        public SeleniumScriptModel ExcelModel { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public ObservableCollection<FileInfo> Files { get; set; } = new ObservableCollection<FileInfo>();

        public FileInfo SelectedFile { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public ICommand BtnReadFile { get; set; }

        public ICommand BtnStart { get; set; }

        public string Arg1 { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");

        /// <summary>
        /// 初始化
        /// </summary>
        public P201ViewModel()
        {
            // 初始化
            this._logic.Init(this);
            //
            this.BtnReadFile = new RelayTCommand<P201ViewModel>(_logic.ReadFile);
            this.BtnStart = new RelayTCommand<P201ViewModel>(_logic.Start);
        }
    }
}
