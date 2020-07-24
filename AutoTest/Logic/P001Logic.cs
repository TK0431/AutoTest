using AutoTest.ViewModels;
using FrameWork.Utility;
using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTest.Logic
{
    /// <summary>
    /// EXE解析
    /// </summary>
    public class P001Logic
    {
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="model"></param>
        public void Init(P001ViewModel model)
        {
            model.FlgDark = XmlUtility.GetXmValue("isdark") == "True" ? true : false;
        }

        /// <summary>
        /// 检索按钮
        /// </summary>
        /// <param name="model"></param>
        public void BtnSearch(P001ViewModel model)
        { }

        /// <summary>
        /// 文件出力按钮
        /// </summary>
        /// <param name="model"></param>
        public void BtnFileOut(P001ViewModel model)
        { }

        public void ApplyBase(P001ViewModel model)
        {
            PaletteHelper _paletteHelper = new PaletteHelper();
            ITheme theme = _paletteHelper.GetTheme();
            IBaseTheme baseTheme = model.FlgDark ? new MaterialDesignDarkTheme() : (IBaseTheme)new MaterialDesignLightTheme();
            theme.SetBaseTheme(baseTheme);
            _paletteHelper.SetTheme(theme);
        }

        public void SaveTheme(P001ViewModel model)
        {
            if(model.SelectedColor != null)
                XmlUtility.SetXmValue("theme", model.SelectedColor.ToString());
            XmlUtility.SetXmValue("isdark", model.FlgDark.ToString());

            App.ShowMessage($"保存成功", "OK");
        }
    }
}
