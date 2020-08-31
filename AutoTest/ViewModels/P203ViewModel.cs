using AutoTest.Logic;
using FrameWork.Models;
using FrameWork.Utility;
using FrameWork.ViewModels.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace AutoTest.ViewModels
{
    /// <summary>
    /// Web测试
    /// </summary>
    public class P203ViewModel : BaseViewModel
    {
        /// <summary>
        /// Logic
        /// </summary>
        private P203Logic _logic = new P203Logic();

        public SeleniumUtility Su { get; set; }

        public string Url { get; set; }

        public ObservableCollection<string> Types { get; set; }

        public BitmapImage Image { get; set; }

        public ICommand StartCommand { get; set; }

        public ICommand AddCommand { get; set; }

        public ICommand ShowCommand { get; set; }

        public ObservableCollection<P203ItemViewModel> Items { get; set; } = new ObservableCollection<P203ItemViewModel>();

        /// <summary>
        /// 初始化
        /// </summary>
        public P203ViewModel()
        {
            _logic.Init(this);

            this.StartCommand = new RelayTCommand<P203ViewModel>(_logic.StartButton);

            this.AddCommand = new RelayTCommand<P203ViewModel>(_logic.AddButton);

            this.ShowCommand = new RelayTCommand<P203ViewModel>(_logic.ShowButton);
        }
    }

    public class P203ItemViewModel : BaseViewModel
    {
        public int Id { get; set; }
        public string Type { get; set; }
        public string Value { get; set; }
    }
}
