using AUTOSYS.Utility;
using AutoTest.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        }
    }
}
