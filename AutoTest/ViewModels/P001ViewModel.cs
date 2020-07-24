using AutoTest.Logic;
using FrameWork.Models;
using FrameWork.ViewModels.Base;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Media;
using System.Windows.Input;
using FrameWork.Utility;

namespace AutoTest.ViewModels
{
    /// <summary>
    /// 设置
    /// </summary>
    public class P001ViewModel : BaseViewModel
    {
        private readonly PaletteHelper _paletteHelper = new PaletteHelper();

        /// <summary>
        /// Logic
        /// </summary>
        private P001Logic _logic = new P001Logic();

        private Color? _selectedColor;
        public Color? SelectedColor { set {
                if (_selectedColor != value)
                {
                    _selectedColor = value;
                    OnPropertyChanged("SelectedColor");

                    if (value is Color color)
                    {
                        ChangeCustomColor(color);
                    }
                }
            } get => _selectedColor; }

        public bool FlgDark { get; set; }

        public IEnumerable<ISwatch> Swatches { get; } = SwatchHelper.Swatches;

        public ICommand ChangeHueCommand { get; }

        public ICommand ToggleBaseCommand { get; }

        public ICommand BtnSaveTheme { get; }

        /// <summary>
        /// 初始化
        /// </summary>
        public P001ViewModel()
        {
            // 初始化
            this._logic.Init(this);

            ChangeHueCommand = new RelayTCommand<Color>(ChangeHue);
            ToggleBaseCommand = new RelayTCommand<P001ViewModel>(_logic.ApplyBase);
            BtnSaveTheme = new RelayTCommand<P001ViewModel>(_logic.SaveTheme);
        }

        private void ChangeHue(Color obj)
        {
            //var hue = (Color)obj;

            //ITheme theme = _paletteHelper.GetTheme();
            //Color? _primaryColor = (Color?)theme.PrimaryMid.Color;
            //SelectedColor = _primaryColor;

            SelectedColor = obj;
            _paletteHelper.ChangePrimaryColor(obj);
            //_primaryColor = hue;
        }

        private void ChangeCustomColor(object obj)
        {
            var color = (Color)obj;


                _paletteHelper.ChangePrimaryColor(color);
                //_primaryColor = color;


                //_paletteHelper.ChangeSecondaryColor(color);
                //_secondaryColor = color;


                //SetPrimaryForegroundToSingleColor(color);
                //_primaryForegroundColor = color;

                //SetSecondaryForegroundToSingleColor(color);
                //_secondaryForegroundColor = color;

        }
    }
}
