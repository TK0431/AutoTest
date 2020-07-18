using AutoTest.ViewModels;
using FrameWork.Consts;
using FrameWork.Models;
using FrameWork.Utility;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTest.Logic
{
    public class P201Logic
    {
        public void Init(P201ViewModel model)
        {
            if (!Directory.Exists(Environment.CurrentDirectory + @"\ExcelScript\"))
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + @"\ExcelScript");
                Directory.CreateDirectory(Environment.CurrentDirectory + @"\ExcelScript\DB");
            }

            DirectoryInfo di = new DirectoryInfo(Environment.CurrentDirectory + @"\ExcelScript\");
            di.GetFiles().ToList().ForEach(x => model.Files.Add(x));

            LogUtility.WriteInfo("初始完了");
        }

        public void ReadFile(P201ViewModel model)
        {
            if (model.SelectedFile == null)
            {
                App.ShowMessage("文件未选择", "Check");
                return;
            }

            model.ExcelModel = new SeleniumScriptModel();
            using (ExcelUtility excel = new ExcelUtility(model.SelectedFile.FullName))
            {
                ExcelWorksheet sh = excel.GetSheet("Menu");

                for (int i = 2; i <= sh.GetMaxRow(1); i++)
                {
                    if (!string.IsNullOrWhiteSpace(sh.Cells[i, 1].Text))
                    {
                        model.ExcelModel.Orders.Add(new SeleniumOrder()
                        {
                            Case = sh.Cells[i, 2].Text,
                            View = sh.Cells[i, 3].Text,
                            ViewName = sh.Cells[i, 4].Text,
                            Event = sh.Cells[i, 5].Text,
                        });
                    }
                }

                List<string> sheets = model.ExcelModel.Orders.Select(x => x.View).ToList();
                sheets = sheets.Distinct().ToList();
                sheets.Sort();

                foreach (string shName in sheets)
                {
                    ExcelWorksheet shTemp = excel.GetSheet(shName);
                    Dictionary<string, List<FrameWork.Models.SeleniumEvent>> events = new Dictionary<string, List<FrameWork.Models.SeleniumEvent>>();
                    if (shTemp == null)
                    {
                        App.ShowMessage("Sheet不存在[{shName}]", "异常", EnumMessageType.Error);
                        return;
                    }

                    for (int col = 1; col <= shTemp.GetMaxColumn(1); col += 3)
                    {
                        if (string.IsNullOrWhiteSpace(shTemp.Cells[1, col].Text)) continue;

                        List<FrameWork.Models.SeleniumEvent> colEvents = new List<FrameWork.Models.SeleniumEvent>();
                        for (int row = 3; row <= shTemp.GetMaxRow(col + 1); row++)
                        {
                            if (!string.IsNullOrWhiteSpace(shTemp.Cells[row, col + 1].Text))
                                colEvents.Add(new FrameWork.Models.SeleniumEvent()
                                {
                                    No = shTemp.Cells[row, col].Text,
                                    Event = shTemp.Cells[row, col + 1].Text,
                                });
                        }

                        events.Add(shTemp.Cells[1, col].Text, colEvents);
                    }

                    model.ExcelModel.Events.Add(shName, events);
                }
            }

            App.ShowMessage("读取完毕");
        }

        public void Start(P201ViewModel model)
        {
            //using (SeleniumUtility su = new SeleniumUtility())
            //{
            //    su.AddArg("$arg(1)", (model.Arg1 + "~" + model.Arg1).Replace("/","-"));

            //    foreach (SeleniumOrder order in model.ExcelModel.Orders)
            //    {
            //        su.ClearElements();
            //        foreach (SeleniumEvent even in model.ExcelModel.Events[order.View][order.Event])
            //        {
            //            su.DoCommand(even);
            //        }
            //    }
            //}

            CreateExel(model, @"E:\GitHub\AutoTest\AutoTest\bin\Debug\20200718154800");

            App.ShowMessage("执行完了");
        }

        private void CreateExel(P201ViewModel model, string path)
        {
            FileInfo file = new FileInfo(path + @"\Sample.xlsx");

            using (ExcelUtility excel = new ExcelUtility(Environment.CurrentDirectory + @"\ExcelScript\Sample.xlsx"))
            {
                ExcelWorksheet sh = excel.GetSheet("对账单明细报表");

                foreach (FileInfo f in new DirectoryInfo(path).GetFiles())
                {
                    if (!f.Name.StartsWith(DateTime.Now.ToString("yyyyMMdd"))) continue;

                    using (ExcelUtility tempExcel = new ExcelUtility(f.FullName))
                    {
                        ExcelWorksheet tempSh = tempExcel.GetSheet();

                        int maxRow = tempSh.GetMaxRow(1);

                        if (maxRow >= 2)
                        {
                            int max = sh.GetMaxRow(1) + 1;
                            tempSh.Cells[2, 1, maxRow, 23].Copy(sh.Cells[max, 1, max + maxRow - 2, 23]);
                        }
                    }
                }

                //刷新透视表
                excel.RefreshAll();

                excel.SaveAs(file);
            }
        }
    }
}
