using AutoTest.ViewModels;
using FrameWork.Consts;
using FrameWork.Models;
using FrameWork.Utility;
using NPOI.SS.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace AutoTest.Logic
{
    public class P201Logic
    {
        private SeleniumUtility _su;

        /// <summary>
        /// 权限设置
        /// </summary>
        /// <param name="model"></param>
        /// <param name="tp"></param>
        private void SetFlgs(P201ViewModel model, int tp)
        {
            switch (tp)
            {
                // 初期化
                case 1:
                    model.FlgFile = true;
                    model.FlgDate = false;
                    model.FlgStart = false;
                    model.FlgContinue = false;
                    model.FlgDoing = Visibility.Hidden;
                    model.FlgStop = false;
                    break;
                // 读取后
                case 2:
                    model.FlgFile = true;
                    model.FlgDate = true;
                    model.FlgStart = true;
                    model.FlgContinue = false;
                    model.FlgDoing = Visibility.Hidden;
                    model.FlgStop = false;
                    break;
                // 执行中
                case 3:
                    model.FlgFile = false;
                    model.FlgDate = false;
                    model.FlgStart = false;
                    model.FlgContinue = false;
                    model.FlgDoing = Visibility.Visible;
                    model.FlgStop = false;
                    break;
                // 异常中断后
                case 4:
                    model.FlgFile = true;
                    model.FlgDate = true;
                    model.FlgStart = false;
                    model.FlgContinue = true;
                    model.FlgDoing = Visibility.Hidden;
                    model.FlgStop = false;
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="model"></param>
        public void Init(P201ViewModel model)
        {
            // 脚本路径创建
            if (!Directory.Exists(Environment.CurrentDirectory + @"\ExcelScript\"))
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + @"\ExcelScript");
                Directory.CreateDirectory(Environment.CurrentDirectory + @"\ExcelScript\DB");
            }

            // 获取内部.xlsx文件
            DirectoryInfo di = new DirectoryInfo(Environment.CurrentDirectory + @"\ExcelScript\");
            di.GetFiles().ToList().Where(x => x.Extension == ".xlsx").ToList().ForEach(x => model.Files.Add(x));

            // 初期化
            SetFlgs(model, 1);

            LogUtility.WriteInfo("初始完了");
        }

        public void ReadFile(P201ViewModel model)
        {
            // 脚本文件选择Check
            if (model.SelectedFile == null)
            {
                App.ShowMessage("文件未选择", "Check");
                return;
            }

            // 下拉框初始化
            model.WebElements = new ObservableCollection<string>();

            // 脚本文件读取
            model.ExcelModel = new SeleniumScriptModel();
            using (ExcelUtility excel = new ExcelUtility(model.SelectedFile.FullName, model.ExcelPassWord))
            {
                // Menu Sheet读取
                ExcelWorksheet sh = excel.GetSheet("Menu");

                // 获取全部Order
                int maxRow = sh.GetMaxRow(1);
                for (int i = 2; i <= maxRow; i++)
                {
                    // A列空跳过
                    if (!string.IsNullOrWhiteSpace(sh.Cells[i, 1].Text))
                    {
                        // 添加Order
                        SeleniumOrder tempOrder = new SeleniumOrder()
                        {
                            Case = sh.Cells[i, 2].Text,
                            View = sh.Cells[i, 3].Text,
                            ViewName = sh.Cells[i, 4].Text,
                            Event = sh.Cells[i, 5].Text,
                        };
                        model.ExcelModel.Orders.Add(tempOrder);

                        LogUtility.WriteInfo($"【Menu】Sheet读取数据【{tempOrder.Case}-{tempOrder.View}-{tempOrder.ViewName}-{tempOrder.Event}】");
                    }
                }
                LogUtility.WriteInfo("【Menu】Sheet读取完毕");

                // 统计Order涉及Sheet
                List<string> sheets = model.ExcelModel.Orders.Select(x => x.View).ToList();
                sheets = sheets.Distinct().ToList();
                sheets.Sort();

                // 遍历Order涉及的Sheet
                foreach (string shName in sheets)
                {
                    // 读取Sheet
                    ExcelWorksheet shTemp = excel.GetSheet(shName);
                    Dictionary<string, List<SeleniumEvent>> events = new Dictionary<string, List<SeleniumEvent>>();
                    if (shTemp == null)
                    {
                        App.ShowMessage("Sheet不存在[{shName}]", "异常", EnumMessageType.Error);
                        return;
                    }

                    // 遍历每个页面
                    int colMax = shTemp.GetMaxColumn(1);
                    for (int col = 1; col <= colMax; col += 4)
                    {
                        // 如果页面未定义编号跳过
                        string webId = shTemp.Cells[1, col].Text;
                        if (string.IsNullOrWhiteSpace(webId)) continue;

                        // 遍历每个页面的Event
                        List<SeleniumEvent> colEvents = new List<SeleniumEvent>();
                        for (int row = 3; row <= shTemp.GetMaxRow(col + 1); row++)
                        {
                            // Event空白跳过
                            if (!string.IsNullOrWhiteSpace(shTemp.Cells[row, col + 2].Text))
                            {
                                // 添加Event
                                SeleniumEvent tempEvent = new SeleniumEvent()
                                {
                                    No = shTemp.Cells[row, col].Text,
                                    Key = shTemp.Cells[row, col + 1].Text,
                                    Event = shTemp.Cells[row, col + 2].Text,
                                    Back = shTemp.Cells[row, col + 3].Text,
                                };
                                colEvents.Add(tempEvent);

                                if (!string.IsNullOrWhiteSpace(tempEvent.Back))
                                    model.WebElements.Add(tempEvent.Back);

                                LogUtility.WriteInfo($"【{shName}-{webId}】Sheet读取数据【{tempEvent.Key}-{tempEvent.Event}】");
                            }
                        }

                        // 添加Web Events
                        events.Add(webId, colEvents);
                        LogUtility.WriteInfo($"【{shName}-{webId}】Sheet读取完毕");
                    }

                    // 添加Sheet Webs Events
                    model.ExcelModel.Events.Add(shName, events);

                    LogUtility.WriteInfo($"【{shName}】Sheet读取完毕");
                }

                // 读取Check Sheet
                ExcelWorksheet shChk = excel.GetSheet("CHK");

                // 采购入库
                int maxCol1 = shChk.GetMaxRow(1);
                for (int i = 4; i <= maxCol1; i++)
                {
                    SeleniumCheckItem checkTemp = new SeleniumCheckItem()
                    {
                        col1 = int.Parse(shChk.Cells[i, 1].Text),
                        name1 = shChk.Cells[i, 2].Text,
                        col2 = int.Parse(shChk.Cells[i, 3].Text),
                        name2 = shChk.Cells[i, 4].Text,
                    };
                    model.ExcelModel.Initems.Add(checkTemp);
                    LogUtility.WriteInfo($"【CHK】Sheet读取数据【{checkTemp.col1}-{checkTemp.name1}-{checkTemp.col2}-{checkTemp.name2}】");
                }

                // 采购退货
                int maxCol2 = shChk.GetMaxRow(5);
                for (int i = 4; i <= maxCol2; i++)
                {
                    SeleniumCheckItem checkTemp = new SeleniumCheckItem()
                    {
                        col1 = int.Parse(shChk.Cells[i, 5].Text),
                        name1 = shChk.Cells[i, 6].Text,
                        col2 = int.Parse(shChk.Cells[i, 7].Text),
                        name2 = shChk.Cells[i, 8].Text,
                    };
                    model.ExcelModel.Outitems.Add(checkTemp);
                    LogUtility.WriteInfo($"【CHK】Sheet读取数据【{checkTemp.col1}-{checkTemp.name1}-{checkTemp.col2}-{checkTemp.name2}】");
                }
                LogUtility.WriteInfo($"【CHK】Sheet读取完毕");
            }

            // 读取完文件后
            SetFlgs(model, 2);

            App.ShowMessage("Excel读取完毕");
        }

        /// <summary>
        /// 自动脚本执行开始
        /// </summary>
        /// <param name="model"></param>
        public void Start(P201ViewModel model)
        {
            ProcessUtility.KillProcess("chromedriver");

            Task task = Task.Run(() => DoSelenium(model));

            // 执行中
            SetFlgs(model, 3);
        }

        private void DoSelenium(P201ViewModel model, string continuEvent = null)
        {
            model.Msg = "浏览器模拟操作中...";

            // 参数设定
            //string arg1 = (model.Arg1 + "~" + model.Arg1).Replace("/", "-");
            //string arg2 = model.Arg1.Replace("/", "-");
            string arg2 = Convert.ToDateTime(model.Arg1).ToString("yyyy-MM-dd");
            string arg1 = arg2 + "~" + arg2;

            if (continuEvent == null)
            {
                _su = new SeleniumUtility();

                _su.AddArg("$arg(1)", arg1);
                _su.AddArg("$arg(2)", arg2);

                // 输出路径
                _su.SuTime = arg2.Replace("-", "") + "_" + _su.SuTime;
                model.OutPath = Environment.CurrentDirectory + @"\" + _su.SuTime;
            }
            else
            {
                _su.SetArg("$arg(1)", arg1);
                _su.SetArg("$arg(2)", arg2);
            }

            LogUtility.WriteInfo($"【参数】-【$arg(1)】-【{arg1}】");
            LogUtility.WriteInfo($"【参数】-【$arg(2)】-【{arg2}】");

            // 遍历Orders
            foreach (SeleniumOrder order in model.ExcelModel.Orders)
            {
                if (model.FlgStop)
                {
                    // 强制终了
                    SetFlgs(model, 4);
                    App.ShowMessage("已强制终了", "OK");
                    return;
                }

                LogUtility.WriteInfo($"【Order開始】-【{order.Case}-{order.View}-{order.ViewName}-{order.Event}】");

                // 清除Elements
                _su.ClearElements();

                try
                {
                    _su.DoEvents(model.ExcelModel.Events[order.View][order.Event], ref continuEvent);
                }
                catch (Exception e)
                {
                    // 异常中断后
                    SetFlgs(model, 4);
                    App.ShowMessage(e.Message, "异常", EnumMessageType.Error, e);
                    return;
                }

                LogUtility.WriteInfo($"【Order终了】-【{order.Case}-{order.View}-{order.ViewName}-{order.Event}】");
            }

            _su.Dispose();

            //model.OutPath = @"E:\GitHub\AutoTest\AutoTest\bin\Debug\20200724124612";

            // Excel 整理
            CreateExel(model, model.OutPath);

            // 执行完了
            SetFlgs(model, 1);

            App.ShowMessage("执行完了");
        }

        /// <summary>
        /// Excel 整理
        /// </summary>
        /// <param name="model"></param>
        /// <param name="path"></param>
        private void CreateExel(P201ViewModel model, string path)
        {
            model.Msg = "透视表数据做成中...";
            // 透视表
            FileInfo file = new FileInfo(path + @"\透视表.xlsx");

            // 创建透视表
            int maxRow;
            using (ExcelUtility excel = new ExcelUtility(Environment.CurrentDirectory + @"\ExcelScript\Sample.xlsx"))
            {
                // 获取【对账单明细报表】Sheet
                ExcelWorksheet sh = excel.GetSheet("对账单明细报表");

                // 遍历获取下载文件
                List<FileInfo> dataFiles = new DirectoryInfo(path).GetFiles().Where(f => f.Name.StartsWith(DateTime.Now.ToString("yyyyMMdd"))).ToList();
                if (dataFiles.Count != 3)
                {
                    SeleniumErr(model, "未能成功下载3个文件");
                    return;
                }

                // 遍历下载文件
                foreach (FileInfo f in dataFiles)
                {
                    using (ExcelUtility tempExcel = new ExcelUtility(f.FullName))
                    {
                        // 获取数据Sheet
                        ExcelWorksheet tempSh = tempExcel.GetSheet();
                        maxRow = tempSh.GetMaxRow(1);

                        // 复制数据
                        if (maxRow >= 2)
                        {
                            int max = sh.GetMaxRow(1) + 1;
                            for (int i = 2; i <= maxRow; i++)
                                for (int j = 1; j <= 23; j++)
                                    sh.Cells[max + i - 2, j].Value = tempSh.Cells[i, j].Text;
                        }
                    }
                    LogUtility.WriteInfo($"【{f.FullName}】数据文件读取完了");
                }

                maxRow = sh.GetMaxRow(1);
                for (int i = 2; i <= maxRow; i++)
                {
                    // L列（含税金额）数值转换
                    double val1;
                    if (double.TryParse(sh.Cells[i, 12].Text, out val1))
                        sh.Cells[i, 12].Value = val1;

                    // S列（商品数量）数值转换
                    double val2;
                    if (double.TryParse(sh.Cells[i, 19].Text, out val2))
                        sh.Cells[i, 19].Value = val2;

                    // 复制公式
                    if (i > 2)
                        sh.Cells["X2:AW2"].Copy(sh.Cells[i, 24, i, 47]);
                }

                LogUtility.WriteInfo($"【透视表】对账单明细报表 数据做成完了");
                excel.SaveAs(file);
                LogUtility.WriteInfo($"【透视表】对账单明细报表 文件保存完了");
            }

            model.Msg = "透视表数据刷新中...";

            //刷新透视表
            int cnt = 0;
            do
            {
                Thread.Sleep(500);
                LogUtility.WriteInfo($"透视表数据刷新中...");
            } while (!ExcelHelper.RefreshPivotTable(path + @"\透视表.xlsx") && cnt++ < 5);
            if (cnt >= 6)
                throw new Exception("Err:透视表数据刷新失败");

            LogUtility.WriteInfo($"【透视表】透视表刷新完了");

            // 采购入库/采购退货 做成
            using (ExcelUtility excel = new ExcelUtility(path + @"\透视表.xlsx"))
            {
                ExcelWorksheet sh = excel.GetSheet("对账单明细报表");

                // 采购入库
                ExcelWorksheet shOrgIn = excel.GetSheet("采购入库");
                int maxOrgInRow = shOrgIn.GetMaxRow(2);

                if (shOrgIn.Cells[maxOrgInRow + 1, 1].Text != "总计")
                {

                }

                model.Msg = "电商销售单导入做成中...";
                using (ExcelXlsUtility excelIn = new ExcelXlsUtility(Environment.CurrentDirectory + @"\ExcelScript\In.xls"))
                {
                    ISheet shIn = excelIn.GetSheet();
                    int shInColCnt = shIn.GetRow(0).Cells.Count;
                    List<string> codes = new List<string>();
                    int indexCode0 = -1;
                    int indexCode1 = -1;

                    foreach (SeleniumCheckItem item in model.ExcelModel.Initems)
                    {
                        string name1 = shIn.GetRow(0).GetCell(item.col1 - 1).StringCellValue;
                        string name2 = shOrgIn.Cells[3, item.col2].Text;
                        if (name1 != item.name1)
                        {
                            SeleniumErr(model, $"【销售】模板已发生变动【{name1}-{item.name2}】");
                            return;
                        }
                        if (name2 != item.name2)
                        {
                            SeleniumErr(model, $"【销售】模板已发生变动【{name2}-{item.name2}】");
                            return;
                        }

                        for (int i = 4; i <= maxOrgInRow; i++)
                        {
                            shIn.CreateRowCells(shInColCnt);
                            shIn.GetRow(i - 3).GetCell(item.col1 - 1).SetCellValue(shOrgIn.Cells[i, item.col2].Text);
                        }

                        if (name1 == "产品编码") indexCode0 = item.col1 - 1;
                        if (name1 == "客户订单号") indexCode1 = item.col1 - 1;

                        LogUtility.WriteInfo($"【退货】数据复制：{item.col2}-{item.name2}->{item.col1}-{item.name1}");
                    }

                    // 重复【产品编码】的【客户订单号】+【.1】
                    for (int i = 4; i <= maxOrgInRow; i++)
                    {
                        string value0 = shIn.GetRow(i - 3).GetCell(indexCode0).StringCellValue;
                        string value1 = shIn.GetRow(i - 3).GetCell(indexCode1).StringCellValue;
                        if (codes.Contains(value0))
                            shIn.GetRow(i - 3).GetCell(indexCode1).SetCellValue(value1 + ".1");
                        else
                            codes.Add(value0);
                    }

                    excelIn.SaveAs(path + @"\电商销售单导入.xls");
                }
                LogUtility.WriteInfo($"【电商销售单导入】做成完了");

                // 采购退货
                model.Msg = "退货单批量导入做成中...";
                ExcelWorksheet shOrgOut = excel.GetSheet("采购退货");
                int maxOrgOutRow = shOrgOut.GetMaxRow(2);

                if (shOrgOut.Cells[maxOrgOutRow + 1, 1].Text != "总计")
                {

                }

                using (ExcelXlsUtility excelOut = new ExcelXlsUtility(Environment.CurrentDirectory + @"\ExcelScript\Out.xls"))
                {
                    ISheet shOut = excelOut.GetSheet();
                    int shOutColCnt = shOut.GetRow(0).Cells.Count;

                    foreach (SeleniumCheckItem item in model.ExcelModel.Outitems)
                    {
                        string name1 = shOut.GetRow(0).GetCell(item.col1 - 1).StringCellValue;
                        string name2 = shOrgOut.Cells[3, item.col2].Text;
                        if (name1 != item.name1)
                        {
                            SeleniumErr(model, $"【退货】模板已发生变动【{name1}-{item.name2}】");
                            return;
                        }
                        if (name2 != item.name2)
                        {
                            SeleniumErr(model, $"【退货】模板已发生变动【{name2}-{item.name2}】");
                            return;
                        }

                        for (int i = 4; i <= maxOrgOutRow; i++)
                        {
                            shOut.CreateRowCells(shOutColCnt);
                            shOut.GetRow(i - 3).GetCell(item.col1 - 1).SetCellValue(shOrgOut.Cells[i, item.col2].Text);
                        }
                        LogUtility.WriteInfo($"【退货】数据复制：{item.col2}-{item.name2}->{item.col1}-{item.name1}");
                    }

                    excelOut.SaveAs(path + @"\退货单批量导入.xls");
                }
                LogUtility.WriteInfo($"【退货单批量导入】做成完了");

                excel.Save();
            }
        }

        private void SeleniumErr(P201ViewModel model, string msg)
        {
            App.ShowMessage(msg, "异常", EnumMessageType.Error);
            model.FlgContinue = true;
        }

        /// <summary>
        /// 继续按钮
        /// </summary>
        /// <param name="model"></param>
        public void BtnContinu(P201ViewModel model)
        {
            if (string.IsNullOrEmpty(model.SelectedElement))
            {
                App.ShowMessage("请选择后续动作", "OK");
                return;
            }

            _su.FlgStop = false;
            Task task = Task.Run(() => DoSelenium(model, model.SelectedElement));

            // 执行中
            SetFlgs(model, 3);
        }

        /// <summary>
        /// 强制终了
        /// </summary>
        /// <param name="model"></param>
        public void Stop(P201ViewModel model)
        {
            model.FlgStop = true;

            if (_su != null)
                _su.FlgStop = true;
        }
    }
}
