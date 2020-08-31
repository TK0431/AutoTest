using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTest.Consts
{
    /// <summary>
    /// 画面
    /// </summary>
    public enum PageEnum
    {
        [Description("")]
        ALL,
        [Description("主程序")]
        W000,
        [Description("设置")]
        P001,
        [Description("EXE解析")]
        P101,
        [Description("EXE测试")]
        P102,
        [Description("EXE结果")]
        P103,
        [Description("猫超")]
        P201,
        [Description("WEB测试")]
        P202,
        [Description("WEB截图辅助")]
        P203,
    }
}
