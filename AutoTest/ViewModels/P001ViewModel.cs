using AutoTest.Logic;
using FrameWork.Models;
using FrameWork.ViewModels.Base;
using System.Collections.ObjectModel;
using System.Windows.Input;

namespace AutoTest.ViewModels
{
    /// <summary>
    /// EXE解析
    /// </summary>
    public class P001ViewModel : BaseViewModel
    {
        /// <summary>
        /// Logic
        /// </summary>
        private P001Logic _logic = new P001Logic();

        /// <summary>
        /// 检索
        /// </summary>
        public string StrSearch { get; set; }

        /// <summary>
        /// 检索按钮
        /// </summary>
        public ICommand BtnSearch { get; set; }

        /// <summary>
        /// 是否全体搜索
        /// </summary>
        public bool IsAllFind { get; set; }

        /// <summary>
        /// 句柄
        /// </summary>
        public ObservableCollection<HwndModel> HwndItems { get; set; }

        /// <summary>
        /// 句柄
        /// </summary>
        public string Hwnd { get; set; }

        /// <summary>
        /// 值
        /// </summary>
        public string Value { get; set; }

        /// <summary>
        /// 类
        /// </summary>
        public string Class { get; set; }

        /// <summary>
        /// 坐标X
        /// </summary>
        public int PointX { get; set; }

        /// <summary>
        /// 坐标Y
        /// </summary>
        public int PointY { get; set; }

        /// <summary>
        /// 宽度
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// 高度
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// 文件出力按钮
        /// </summary>
        public ICommand BtnFileOut { get; set; }

        /// <summary>
        /// 初始化
        /// </summary>
        public P001ViewModel()
        {
            // 初始化
            this._logic.Init(this);

            // 检索按钮
            this.BtnSearch = new RelayTCommand<P001ViewModel>(_logic.BtnSearch);
            // 文件出力按钮
            this.BtnSearch = new RelayTCommand<P001ViewModel>(_logic.BtnFileOut);
        }
    }
}
