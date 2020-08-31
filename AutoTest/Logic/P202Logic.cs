using AutoTest.ViewModels;
using FrameWork.Consts;
using FrameWork.Models;
using FrameWork.Utility;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace AutoTest.Logic
{
    public class P202Logic
    {
        private SeleniumUtility _su;

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="model"></param>
        public void Init(P202ViewModel model)
        {
            string path = Environment.CurrentDirectory + @"\ExcelScript";

            // 脚本路径创建
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            // 获取内部.xlsx .xlsm文件
            DirectoryInfo di = new DirectoryInfo(path);
            di.GetFiles().ToList().Where(x => x.Extension == ".xlsx" || x.Extension == ".xlsm").ToList().ForEach(x => model.Files.Add(x));

            LogUtility.WriteInfo("初始完了");
        }

        public void ReadFile(P202ViewModel model)
        {
            // 脚本文件选择Check
            if (model.SelectedFile == null)
            {
                App.ShowMessage("文件未选择", "Check");
                return;
            }

            // 脚本文件读取
            model.ExcelModel = new SeleniumScriptModel();
            using (ExcelUtility excel = new ExcelUtility(model.SelectedFile.FullName))
            {
                // WSet读取
                ExcelWorksheet sh = excel.GetSheet("WSet");
                int maxRow = sh.GetMaxRow(1);
                for (int i = 2; i <= maxRow; i++)
                {
                    string value = sh.Cells[i, 2].Text;
                    if (string.IsNullOrWhiteSpace(value)) continue;
                    switch (sh.Cells[i, 1].Text)
                    {
                        case "超时(秒)":
                            model.ExcelModel.OutTime = int.Parse(value);
                            break;
                        case "延时(毫秒)":
                            model.ExcelModel.ReTry = int.Parse(value);
                            break;
                        case "导出路径":
                                model.ExcelModel.OutPath = value;
                            break;
                        case "模板名":
                            model.ExcelModel.TemplateFile = value;
                            break;
                        case "模板页":
                            model.ExcelModel.TemplateSheet = value;
                            break;
                        case "图开始行":
                            model.ExcelModel.StartRow = int.Parse(value);
                            break;
                        case "图开始列":
                            model.ExcelModel.StartCol = int.Parse(value);
                            break;
                        default:
                            if (sh.Cells[i, 1].Text.StartsWith("$"))
                                model.ExcelModel.Args.Add(sh.Cells[i, 1].Text, value);
                            break;
                    }
                }

                // WMenu读取
                sh = excel.GetSheet("WMenu");

                // 获取全部Order
                maxRow = sh.GetMaxRow(1);
                int maxCol = sh.GetMaxColumn(1);
                for (int i = 2; i <= maxRow; i++)
                {
                    // A列空跳过
                    if (!string.IsNullOrWhiteSpace(sh.Cells[i, 1].Text))
                    {
                        // 添加Order
                        SeleniumOrder tempOrder = new SeleniumOrder()
                        {
                            File = sh.Cells[i, 2].Text,
                            Case = sh.Cells[i, 3].Text,
                            Sid = sh.Cells[i, 4].Text,
                            Back = sh.Cells[i, 5].Text,
                        };

                        // Add Args
                        if (maxCol > 5)
                            for (int j = 6; j <= maxCol; j++)
                                if (!string.IsNullOrWhiteSpace(sh.Cells[1, j].Text) &&
                                   !string.IsNullOrWhiteSpace(sh.Cells[i, j].Text))
                                    tempOrder.Args.Add(sh.Cells[1, j].Text, sh.Cells[i, j].Text);

                        model.ExcelModel.Orders.Add(tempOrder);

                        LogUtility.WriteInfo($"【Menu】Sheet读取数据【{tempOrder.File}-{tempOrder.Case}-{tempOrder.Sid}-{tempOrder.Back}】");
                    }
                }
                LogUtility.WriteInfo("【Menu】Sheet读取完毕");

                // 统计Order涉及Sheet
                List<string> sheets = model.ExcelModel.Orders.Select(x => x.Sid).ToList();
                sheets = sheets.Distinct().ToList();
                sheets.Sort();

                // 遍历Order涉及的Sheet
                foreach (string shName in sheets)
                {
                    // 读取Sheet
                    ExcelWorksheet shTemp = excel.GetSheet(shName);
                    if (shTemp == null)
                    {
                        App.ShowMessage("Sheet不存在[{shName}]", "异常", EnumMessageType.Error);
                        return;
                    }

                    Dictionary<string, List<SeleniumEvent>> shEvents = new Dictionary<string, List<SeleniumEvent>>();

                    // 遍历每个页面
                    int colMax = shTemp.GetMaxColumn(1);
                    for (int col = 1; col <= colMax; col += 6)
                    {
                        // 如果页面未定义编号跳过
                        string caseId = shTemp.Cells[1, col].Text;
                        if (string.IsNullOrWhiteSpace(caseId)) continue;

                        // 遍历每个页面的Event
                        List<SeleniumEvent> colEvents = new List<SeleniumEvent>();
                        for (int row = 3; row <= shTemp.GetMaxRow(col + 2); row++)
                        {
                            // Event空白跳过
                            if (!string.IsNullOrWhiteSpace(shTemp.Cells[row, col + 2].Text) ||
                                !string.IsNullOrWhiteSpace(shTemp.Cells[row, col + 3].Text))
                            {
                                // 添加Event
                                SeleniumEvent tempEvent = new SeleniumEvent()
                                {
                                    No = shTemp.Cells[row, col].Text,
                                    Id = shTemp.Cells[row, col + 1].Text,
                                    Cmd = shTemp.Cells[row, col + 2].Text,
                                    Value = shTemp.Cells[row, col + 3].Text,
                                    Range = shTemp.Cells[row, col + 4].Text,
                                };
                                colEvents.Add(tempEvent);

                                LogUtility.WriteInfo($"【{shName}-{caseId}】Sheet读取数据【{tempEvent.Cmd}-{tempEvent.Value}】");
                            }
                        }

                        // 添加Web Events
                        shEvents.Add(caseId, colEvents);
                        LogUtility.WriteInfo($"【{shName}-{caseId}】Sheet读取完毕");
                    }

                    // 添加Sheet Webs Events
                    model.ExcelModel.Events.Add(shName, shEvents);

                    LogUtility.WriteInfo($"【{shName}】Sheet读取完毕");
                }

                App.ShowMessage("Excel读取完毕");
            }
        }

        /// <summary>
        /// 自动脚本执行开始
        /// </summary>
        /// <param name="model"></param>
        public void Start(P202ViewModel model)
        {
            ProcessUtility.KillProcess("chromedriver");

            Task task = Task.Run(() => DoSelenium(model));
        }

        private void DoSelenium(P202ViewModel model)
        {
            model.Msg = "浏览器模拟操作中...";

            // Add Path
            model.ExcelModel.OutPath = Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss");

            //if (_su == null) 
            _su = new SeleniumUtility(model.ExcelModel);

            _su.DoEvents(model.ExcelModel);

            _su.Dispose();
            _su = null;

            App.ShowMessage("执行完了");
        }

        /// <summary>
        /// Excel 整理
        /// </summary>
        /// <param name="model"></param>
        /// <param name="path"></param>
        public void CreateExel(P202ViewModel model)
        {
            using (ExcelUtility excel = new ExcelUtility(Environment.CurrentDirectory + @"\ExcelScript\" + model.ExcelModel.TemplateFile))
            {
                string picPath = model.ExcelModel.OutPath + @"\pics";
                List<PicRectModel> rects = PicRectModel.GetList(picPath + @"\log.txt");

                foreach(string dir in Directory.GetDirectories(picPath))
                {
                    string caseName = (new DirectoryInfo(dir)).Name;
                    ExcelWorksheet sh = excel.AddSheet(caseName, "TMP");

                    int row = model.ExcelModel.StartRow;
                    int col = model.ExcelModel.StartCol;
                    int rowHeight = sh.GetHeightPix();
                    int colWidth = sh.GetWidthPix();
                    int picHeight;

                    foreach (string file in Directory.GetFiles(dir))
                    {
                        FileInfo picFile = new FileInfo(file);
                        ExcelPicture pic = sh.Drawings.AddPicture(picFile.Name, picFile);
                        pic.SetPosition(row, 0, col, 0);
                        pic.SetSize(pic.Image.Width, pic.Image.Height);
                        picHeight = pic.Image.Height;

                        ExcelShape sps0 = sh.AddShape("000");
                        sps0.SetSize(79, 19);
                        sps0.SetPosition(19, 79);

                        List<PicRectModel> prm = rects.Where(x => x.Sheet == sh.Name && x.PicNo.ToString() == picFile.Name.Replace(picFile.Extension,"")).ToList();

                        int i = 0;
                        foreach(PicRectModel rect in prm)
                        { 
                            ExcelShape sps = sh.AddShape(picFile.Name + (i++).ToString());
                            sps.SetSize((int)(rect.Width * (4/4.625)), rect.Height);
                            sps.SetPosition((int)(pic.From.Row * rowHeight + rect.Y), (int)(pic.From.Column * colWidth + rect.X * (4 / 4.625)));
                        }

                        int cnt = 0;
                        do
                        {
                        } while (rowHeight * cnt++ < picHeight);
                        row += cnt;
                    }
                }

                //foreach (PicRectModel rect in rects)
                //{
                //    ExcelWorksheet sh = excel.GetSheet(rect.Sheet);
                //    if (sh == null)
                //        excel.AddSheet()
                //}

                excel.DelSheet("TMP");

                string fileName = model.ExcelModel.OutPath + @"\" + model.ExcelModel.TemplateFile;
                excel.SaveAs(fileName);
            }
        }

        /// <summary>
        /// 强制终了
        /// </summary>
        /// <param name="model"></param>
        public void Stop(P202ViewModel model)
        {
        }
    }
}
