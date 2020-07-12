using AutoTest.Consts;
using AutoTest.Logic;
using AutoTest.Pages;
using FrameWork.Models;
using FrameWork.Utility;
using FrameWork.ViewModels.Base;
using System.Collections.ObjectModel;
using System.Windows.Controls;
using System.Windows.Input;

namespace AutoTest.ViewModels
{
    /// <summary>
    /// 主画面
    /// </summary>
    public class W000ViewModel:BaseViewModel
    {
        /// <summary>
        /// Logic
        /// </summary>
        private W000Logic _logic = new W000Logic();

        /// <summary>
        /// 主Page
        /// </summary>
        public Page MainPage { get; set; } = new P102();

        /// <summary>
        /// 用户名
        /// </summary>
        public string UName { get; set; }

        /// <summary>
        /// Wbs 菜单
        /// </summary>
        public ObservableCollection<EnumItem> MenuItems { get; set; }

        /// <summary>
        /// 上传Excel
        /// </summary>
        public ICommand Login { get; set; }

        /// <summary>
        /// 初期化
        /// </summary>
        public W000ViewModel()
        {
            // 初期化
            this._logic.Init(this);

            // Wbs 菜单
            MenuItems = new ObservableCollection<EnumItem>()
            {
                //PageEnum.P001.GetItem(), // 设置
                PageEnum.P101.GetItem(), // EXE解析
                PageEnum.P102.GetItem(), // EXE测试
                PageEnum.P103.GetItem(), // EXE结果
                PageEnum.P201.GetItem(), // WEB测试
                PageEnum.P202.GetItem(), // WEB结果
            };
        }
    }
}
