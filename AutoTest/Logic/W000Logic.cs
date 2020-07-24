using AutoTest.ViewModels;
using FrameWork.Utility;
using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace AutoTest.Logic
{
    /// <summary>
    /// 主窗口
    /// </summary>
    public class W000Logic
    {
        /// <summary>
        /// 初期化
        /// </summary>
        /// <param name="model"></param>
        public void Init(W000ViewModel model)
        {
            // 获取项目
            //model.UName = XmlUtility.GetXmValue("uname");
            model.Version = XmlUtility.GetXmValue("version");

            new PaletteHelper().ChangePrimaryColor((Color)ColorConverter.ConvertFromString(XmlUtility.GetXmValue("theme")));

            if (XmlUtility.GetXmValue("isdark") == "True")
            {
                PaletteHelper _paletteHelper = new PaletteHelper();
                ITheme theme = _paletteHelper.GetTheme();
                IBaseTheme baseTheme = new MaterialDesignDarkTheme();
                theme.SetBaseTheme(baseTheme);
                _paletteHelper.SetTheme(theme);
            }
        }
    }
}
