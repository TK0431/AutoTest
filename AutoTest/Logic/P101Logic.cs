using AutoTest.ViewModels;
using FrameWork.Models;
using FrameWork.Utility;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace AutoTest.Logic
{
    /// <summary>
    /// EXE解析
    /// </summary>
    public class P101Logic
    {
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="model"></param>
        public void Init(P101ViewModel model)
        {
            model.HwndItems = new ObservableCollection<HwndItem>();
            List<HwndModel> list = HwndUtility.GetDeskHwndModels();
            list.Where(x => !string.IsNullOrEmpty(x.Value)).ToList().ForEach(x => model.HwndItems.Add(new HwndItem(x)));

            LogUtility.WriteInfo("EXE解析:初始化完了");
        }

        /// <summary>
        /// 检索按钮
        /// </summary>
        /// <param name="model"></param>
        public void BtnSearch(P101ViewModel model)
        {
            model.HwndItems = new ObservableCollection<HwndItem>();
            if (!string.IsNullOrWhiteSpace(model.StrSearch))
                HwndUtility.GetDeskHwndModels().Where(x => x.Value.Contains(model.StrSearch)).ToList().ForEach(x => model.HwndItems.Add(new HwndItem(x)));
            else if (model.IsAllFind)
                HwndUtility.GetDeskHwndModels().ForEach(x => model.HwndItems.Add(new HwndItem(x)));
            else
                HwndUtility.GetDeskHwndModels().Where(x => !string.IsNullOrEmpty(x.Value)).ToList().ForEach(x => model.HwndItems.Add(new HwndItem(x)));

            LogUtility.WriteInfo($"EXE解析:检索完了{model.StrSearch}");
        }

        /// <summary>
        /// 文件出力按钮
        /// </summary>
        /// <param name="model"></param>
        public void BtnFileOut(P101ViewModel model)
        {
            if (model.SelectedHwndItem == null)
            {
                Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"未选中控件", "OK", delegate () { }));
                LogUtility.WarnInfo($"EXE解析:未选中控件");
                return;
            }

            IntPtr hwnd = HwndUtility.GetTopParentHwnd(model.SelectedHwndItem.HModel.HwndId);

            List<HwndModel> list = HwndUtility.GetAllModels(hwnd);
            LogUtility.WriteInfo($"EXE解析:控件{hwnd}检索,获取{list.Count}自控件");

            // TC Templet Copy
            string outPath = Environment.CurrentDirectory + @"\ID_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";

            // Excel 操作
            using (ExcelUtility db = new ExcelUtility(outPath))
            {
                ExcelWorksheet sheet = db.AddSheet("ID");

                sheet.Cells[1, 1].Value = "No";
                sheet.Cells[1, 2].Value = "Value";
                sheet.Cells[1, 3].Value = "Class";
                sheet.Cells[1, 4].Value = "Left";
                sheet.Cells[1, 5].Value = "Top";
                sheet.Cells["A1:E1"].SetRangeColor(Color.FromArgb(68, 114, 196));

                Enumerable.Range(2, list.Count).ToList().ForEach(i =>
                {
                    sheet.Cells[i, 1].Value = i - 1;
                    sheet.Cells[i, 2].Value = list[i - 2].Value;
                    sheet.Cells[i, 3].Value = list[i - 2].Class;
                    sheet.Cells[i, 4].Value = list[i - 2].ExeX;
                    sheet.Cells[i, 5].Value = list[i - 2].ExeY;
                    LogUtility.WriteInfo($"EXE解析:获取{i - 1}控件");
                });

                Enumerable.Range(2 + list.Count, model.AddControl.Count).ToList().ForEach(i =>
                {
                    sheet.Cells[i, 1].Value = i - 1;
                    sheet.Cells[i, 2].Value = model.AddControl[i - 2 - list.Count].Value;
                    sheet.Cells[i, 3].Value = model.AddControl[i - 2 - list.Count].Class;
                    sheet.Cells[i, 4].Value = model.AddControl[i - 2 - list.Count].ExeX;
                    sheet.Cells[i, 5].Value = model.AddControl[i - 2 - list.Count].ExeY;
                    LogUtility.WriteInfo($"EXE解析:获取自主追加{i - 1}控件");
                });

                db.Save();
            }

            Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"出力完了", null, null, null, false, true, TimeSpan.FromMilliseconds(500)));
            LogUtility.WriteInfo($"EXE解析:文件出力完了{outPath}");
        }
    }
}
