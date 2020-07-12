using AutoTest.Logic;
using FrameWork.Consts;
using FrameWork.Models;
using FrameWork.Utility;
using FrameWork.ViewModels.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;

namespace AutoTest.ViewModels
{
    /// <summary>
    /// EXE解析
    /// </summary>
    public class P101ViewModel : BaseViewModel
    {
        /// <summary>
        /// Logic
        /// </summary>
        private P101Logic _logic = new P101Logic();

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
        public ObservableCollection<HwndItem> HwndItems { get; set; }

        public List<HwndModel> AddControl { get; set; } = new List<HwndModel>();

        public HwndItem SelectedHwndItem { get; set; }

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
        public P101ViewModel()
        {
            // 初始化
            this._logic.Init(this);

            // 检索按钮
            this.BtnSearch = new RelayTCommand<P101ViewModel>(_logic.BtnSearch);

            // 文件出力按钮
            this.BtnFileOut = new RelayTCommand<P101ViewModel>(_logic.BtnFileOut);
        }
    }

    public class HwndItem : BaseViewModel
    {
        public HwndItem(HwndModel _model)
        {
            // Create commands
            this.ExpandCommand = new RelayCommand(Expand);

            // Set path and Type
            this.HModel = _model;

            // Setup the children as needed
            this.ClearChildren();
        }

        public HwndModel HModel { get; set; }

        public ObservableCollection<HwndItem> Children { get; set; }

        public ICommand ExpandCommand { get; set; }

        public bool IsExpanded
        {
            get
            {
                return Children?.Count(x => x != null) > 0;
            }
            set
            {
                if (value == true)
                    Expand();
                else
                    this.ClearChildren();
            }
        }

        private void ClearChildren()
        {
            this.Children = new ObservableCollection<HwndItem>();

            if (HModel.Type != HwndItemType.Control && HwndUtility.GetChildHwnd(HModel.HwndId) != IntPtr.Zero)
                this.Children.Add(null);
        }

        private void Expand()
        {
            if (HModel.Type == HwndItemType.Control)
                return;

            Children = new ObservableCollection<HwndItem>();
            HwndUtility.GetChildrenModels(HModel.HwndId).ForEach(x => Children.Add(new HwndItem(x)));
        }

        /// <summary>
        /// Tree显示
        /// </summary>
        public string TreeTextHwnd
        {
            get
            {
                return "0x" + HModel.HwndId.ToString().PadLeft(8, '0') + "(" + HModel.HwndId.ToString("D10") + ")";
            }
        }

        /// <summary>
        /// Tree显示
        /// </summary>
        public string TreeTextValue
        {
            get
            {
                return "【 " + HModel.Value + " 】";
            }
        }

        /// <summary>
        /// Tree显示
        /// </summary>
        public string TreeTextClass
        {
            get
            {
                return HModel.Class;
            }
        }
    }
}
