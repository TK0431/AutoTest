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
    public class P102ViewModel : BaseViewModel
    {
        /// <summary>
        /// Logic
        /// </summary>
        private P102Logic _logic = new P102Logic();

        /// <summary>
        /// 
        /// </summary>
        public ObservableCollection<FileInfo> Files { get; set; } = new ObservableCollection<FileInfo>();

        public FileInfo SelectedFile { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public ICommand BtnReadFile { get; set; }

        public bool FlgNewDir { get; set; } = true;

        /// <summary>
        /// 
        /// </summary>
        public bool FlgCodeOld { get; set; } = true;

        /// <summary>
        /// 
        /// </summary>
        public ICommand BtnStart { get; set; }

        public bool FlgStart { get; set; } = false;

        public ICommand BtnContinue { get; set; }

        public int CbxRowHeight { get; set; } = 15;

        public bool FlgContinue { get; set; } = false;

        public OrderItem ContinueOrder { get; set; }

        public string OutResultPath { get; set; }

        public int KeySleep { get; set; } = 100;

        /// <summary>
        /// 
        /// </summary>
        public int PicNum { get; set; } = 1;

        public string CaseNo { get; set; } = "Case001";

        public string PicPath { get; set; } = Environment.CurrentDirectory + @"\Pics";

        public ExcelScriptModel ExcelData { get; set; }

        public string ComparePath { get; set; } = Environment.CurrentDirectory + @"\Result";

        public int PicTop { get; set; } = 7;
        public int PicButtom { get; set; } = 7;
        public int PicLeft { get; set; } = 7;
        public int PicRight { get; set; } = 7;

        public ICommand BtnCompare { get; set; }

        /// <summary>
        /// 初始化
        /// </summary>
        public P102ViewModel()
        {
            // 初始化
            this._logic.Init(this);

            //
            this.BtnReadFile = new RelayTCommand<P102ViewModel>(_logic.ReadFile);
            // 
            this.BtnStart = new RelayTCommand<P102ViewModel>(_logic.Start);

            this.BtnContinue = new RelayTCommand<P102ViewModel>(_logic.Continue);

            this.BtnCompare = new RelayTCommand<P102ViewModel>(_logic.Compare);
        }
    }
}
