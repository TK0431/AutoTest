﻿using AutoTest.ViewModels;
using FrameWork.Models;
using FrameWork.Utility;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoTest.Logic
{
    /// <summary>
    /// EXE测试
    /// </summary>
    public class P102Logic
    {
        public static string SOLD = "現行";
        public static string SNEW = "次期";

        public void Init(P102ViewModel model)
        {
            DirectoryInfo di = new DirectoryInfo(Environment.CurrentDirectory + @"\ExcelScript\");
            di.GetFiles().ToList().ForEach(x => model.Files.Add(x));
            LogUtility.WriteInfo("EXE测试：初始化完了");
        }

        /// <summary>
        /// 读取Excel文件
        /// </summary>
        /// <param name="model"></param>
        public void ReadFile(P102ViewModel model)
        {
            model.ExcelData = new ExcelScriptModel();
            model.FlgStart = false;
            model.FlgContinue = false;

            if (model.SelectedFile == null)
            {
                Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"请选择文件", "OK", delegate () { }));
                LogUtility.WarnInfo("EXE测试：文件未选择");
                return;
            }

            using (ExcelUtility excel = new ExcelUtility(model.SelectedFile.FullName))
            {
                ExcelWorksheet sh = excel.GetSheet("実行シナリオ");
                if (sh == null)
                {
                    Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"[実行シナリオ]不存在", "OK", delegate () { }));
                    LogUtility.WriteError($"[実行シナリオ]Sheet不存在", null);
                    return;
                }

                model.ExcelData.Sheets = new List<string>();
                model.ExcelData.OrderList = new Dictionary<string, List<OrderItem>>();
                string CaseNo = null;
                for (int i = 2; i <= sh.GetMaxRow(); i++)
                {
                    if (string.IsNullOrWhiteSpace(sh.Cells[i, 1].Text)) continue;

                    OrderItem order = new OrderItem()
                    {
                        CaseNo = sh.Cells[i, 2].Text,
                        ViewName = sh.Cells[i, 3].Text,
                        Sheet = sh.Cells[i, 4].Text,
                        Order = sh.Cells[i, 5].Text,
                    };

                    if (CaseNo != order.CaseNo)
                    {
                        model.ExcelData.OrderList.Add(order.CaseNo, new List<OrderItem>());
                    }
                    model.ExcelData.OrderList[order.CaseNo].Add(order);
                    LogUtility.WriteInfo($"EXE测试:[実行シナリオ]获取{order.CaseNo}-{order.Order}");

                    CaseNo = order.CaseNo;

                    if (!model.ExcelData.Sheets.Contains(order.Sheet))
                        model.ExcelData.Sheets.Add(order.Sheet);
                }
                LogUtility.WriteInfo($"EXE测试:[実行シナリオ]获取完了");

                model.ExcelData.TCList = new Dictionary<string, TCSheet>();
                model.ExcelData.DBList = new Dictionary<string, DBSheet>();
                model.ExcelData.Sheets.ForEach(shName =>
                {
                    ExcelWorksheet sh1 = excel.GetSheet(shName);
                    if (sh1 == null)
                    {
                        Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"[{shName}]不存在", "OK", delegate () { }));
                        LogUtility.WriteError($"EXE测试：[{shName}]Sheet不存在", null);
                        return;
                    }

                    switch (shName.Substring(0, 2))
                    {
                        case "TC":
                            TCSheet tc = new TCSheet();
                            tc.ControlItems = new Dictionary<string, ControlInfo>();
                            int maxRow = sh1.GetMaxRow();
                            for (int i = 4; i <= maxRow; i++)
                            {
                                if (string.IsNullOrWhiteSpace(sh1.Cells[i, 1].Text)) continue;
                                try
                                {
                                    ControlInfo ctl = new ControlInfo()
                                    {
                                        Id = sh1.Cells[i, 1].Text,
                                        Name = sh1.Cells[i, 2].Text,
                                        Class = sh1.Cells[i, 3].Text,
                                        X = int.Parse(sh1.Cells[i, 4].Text),
                                        Y = int.Parse(sh1.Cells[i, 5].Text),
                                    };
                                    tc.ControlItems.Add(ctl.Id, ctl);
                                    LogUtility.WriteInfo($"EXE测试：[{shName}]获取控件信息{ctl.Id}");
                                }
                                catch (Exception)
                                {
                                    Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"[{shName}]{i}行异常", "OK", delegate () { }));
                                    LogUtility.WriteError($"EXE测试：[{shName}]{i}行异常", null);
                                    return;
                                }
                            }

                            tc.CaseDatas = new Dictionary<string, Dictionary<string, string>>();
                            for (int i = 6; i <= sh1.GetMaxColumn(); i++)
                            {
                                if (string.IsNullOrWhiteSpace(sh1.Cells[1, i].Text)) continue;
                                Dictionary<string, string> dic = new Dictionary<string, string>();
                                try
                                {
                                    dic.Add("Sleep", sh1.Cells[2, i].Text);
                                    int sno;
                                    if (sh1.Cells[2, i].Text != "Wait" && !int.TryParse(sh1.Cells[2, i].Text, out sno))
                                    {
                                        Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"[{shName}]{i}列Sleep异常", "OK", delegate () { }));
                                        LogUtility.WriteError($"EXE测试：[{shName}]{i}列Sleep异常", null);
                                        return;
                                    }

                                    dic.Add("Pic", sh1.Cells[3, i].Text);

                                    for (int j = 4; j <= maxRow; j++)
                                    {
                                        if (!string.IsNullOrWhiteSpace(sh1.Cells[j, 1].Text) && !string.IsNullOrWhiteSpace(sh1.Cells[j, i].Text))
                                        {
                                            dic.Add(sh1.Cells[j, 1].Text, sh1.Cells[j, i].Text);
                                        }
                                    }

                                    tc.CaseDatas.Add(sh1.Cells[1, i].Text, dic);
                                    LogUtility.WriteInfo($"EXE测试：[{shName}]获取列Case信息{sh1.Cells[1, i].Text}");
                                }
                                catch (Exception)
                                {
                                    Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"[{shName}]{i}列异常", "OK", delegate () { }));
                                    LogUtility.WriteError($"EXE测试：[{shName}]{i}列异常", null);
                                    return;
                                }
                            }

                            model.ExcelData.TCList.Add(shName, tc);
                            break;
                        case "DB":
                            DBSheet db = new DBSheet();
                            db.CaseDatas = new Dictionary<string, DBSheetInfo>();
                            for (int i = 2; i <= sh1.GetMaxRow(); i++)
                            {
                                if (string.IsNullOrWhiteSpace(sh1.Cells[i, 1].Text)) continue;
                                try
                                {
                                    DBSheetInfo info = new DBSheetInfo()
                                    {
                                        Id = sh1.Cells[i, 1].Text,
                                        Server = sh1.Cells[i, 2].Text,
                                        DataBase = sh1.Cells[i, 3].Text,
                                        User = sh1.Cells[i, 4].Text,
                                        PassWord = sh1.Cells[i, 5].Text,
                                        Table = sh1.Cells[i, 6].Text,
                                        IsDown = sh1.Cells[i, 7].Text,
                                        FileName = sh1.Cells[i, 8].Text,
                                        Sql = sh1.Cells[i, 9].Text,
                                        Sleep = string.IsNullOrWhiteSpace(sh1.Cells[i, 10].Text) ? 0 : int.Parse(sh1.Cells[i, 10].Text),
                                    };

                                    db.CaseDatas.Add(info.Id, info);
                                    LogUtility.WriteInfo($"EXE测试:[{shName}]获取行{info.Id}信息");
                                }
                                catch (Exception)
                                {
                                    Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"[{shName}]{i}行异常", "OK", delegate () { }));
                                    LogUtility.WriteError($"EXE测试：[{shName}]{i}行异常", null);
                                    return;
                                }
                            }

                            model.ExcelData.DBList.Add(shName, db);
                            break;
                        default:
                            Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"[{shName}]请以TC,DB命名", "OK", delegate () { }));
                            LogUtility.WriteError($"EXE测试：[{shName}]命名不正确", null);
                            return;
                    }
                    LogUtility.WriteInfo($"EXE测试：[{shName}]获取完了");
                });
            }

            model.FlgStart = true;
            Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"读取完毕", null, null, null, false, true, TimeSpan.FromMilliseconds(500)));
            LogUtility.WriteInfo($"EXE测试：获取完了");
        }

        #region 自动测试

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        public void Start(P102ViewModel model)
        {
            model.PicNum = 1;
            // 创建路径
            CreatePath(model);
            StartTest(model);

            Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"测试完了", "OK", delegate () { }));
            LogUtility.WriteInfo($"EXE测试：测试完了");
        }

        public void Continue(P102ViewModel model)
        {
            StartTest(model);

            Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"测试完了", "OK", delegate () { }));
            LogUtility.WriteInfo($"EXE测试：测试完了");
        }

        /// <summary>
        /// 创建路径
        /// </summary>
        /// <param name="model"></param>
        private void CreatePath(P102ViewModel model)
        {
            // Create Result path
            model.OutResultPath = Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss");
            Directory.CreateDirectory(model.OutResultPath);
            LogUtility.WriteInfo($"EXE测试:路径创建[{model.OutResultPath}]");

            if (model.FlgCodeOld)
            {
                model.OutResultPath = model.OutResultPath + @"\" + SNEW;
            }
            else
            {
                model.OutResultPath = model.OutResultPath + @"\" + SOLD;
            }
            Directory.CreateDirectory(model.OutResultPath);
            LogUtility.WriteInfo($"EXE测试:路径创建[{model.OutResultPath}]");
        }

        private void StartTest(P102ViewModel model)
        {
            foreach (string key in model.ExcelData.OrderList.Keys)
            {
                // 获取Case00X
                int no;
                string caseNo;
                if (int.TryParse(key, out no))
                {
                    caseNo = "Case" + key.PadLeft(3, '0');
                }
                else
                {
                    caseNo = key;
                }

                // 循环Order
                foreach (OrderItem order in model.ExcelData.OrderList[key])
                {
                    if (order.Sheet.StartsWith("TC"))
                        DoTCOrder(model, caseNo, order);
                    else if (order.Sheet.StartsWith("DB"))
                        DoDBOrder(model, caseNo, order);
                }
            }
        }

        private string _currentTCCasePath;
        Dictionary<string, IntPtr> _currHwnds;
        HwndModel _topHwndModel;

        /// <summary>
        /// TC
        /// </summary>
        /// <param name="model"></param>
        /// <param name="caseNo"></param>
        /// <param name="order"></param>
        private void DoTCOrder(P102ViewModel model, string caseNo, OrderItem order)
        {
            Dictionary<string, string> caseData = model.ExcelData.TCList[order.Sheet].CaseDatas[order.Order];

            if (model.ContinueOrder != null && model.ContinueOrder != order)
                return;
            else if (model.ContinueOrder != null)
                model.ContinueOrder = null;
            else if (caseData["Sleep"] == "Wait")
            {
                model.ContinueOrder = order;
                model.FlgContinue = true;
                throw new Exception("【停止】");
            }
            else
            {
                Thread.Sleep(int.Parse(caseData["Sleep"]));
            }

            // Case Path
            _currentTCCasePath = model.OutResultPath + @"\Picture\" + caseNo;
            if (!Directory.Exists(_currentTCCasePath))
                Directory.CreateDirectory(_currentTCCasePath);

            Dictionary<string, ControlInfo> caseCtrl = model.ExcelData.TCList[order.Sheet].ControlItems;

            // 截图
            IntPtr hwnd = User32Utility.FindWindow(
                string.IsNullOrWhiteSpace(caseCtrl["1"].Class) ? null : caseCtrl["1"].Class,
                string.IsNullOrWhiteSpace(caseCtrl["1"].Name) ? null : caseCtrl["1"].Name
                );
            if (hwnd == IntPtr.Zero)
            {
                Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"Exe不存在[{caseCtrl["1"].Name}-{caseCtrl["1"].Class}]", "OK", delegate () { }));
                LogUtility.WriteError($"EXE测试：Exe不存在[{caseCtrl["1"].Name}-{caseCtrl["1"].Class}]", null);
                return;
            }
            _topHwndModel = HwndUtility.GetHwndModel(hwnd);
            if (caseData["Pic"] == "●") GetPicture(model, hwnd, order);

            List<HwndModel> hwnds = HwndUtility.GetAllModels(hwnd);
            // 真实存在的Hwnds
            _currHwnds = new Dictionary<string, IntPtr>();
            foreach (string key in caseCtrl.Keys)
            {
                if (caseCtrl[key].Class == "$CustomControl") continue;

                foreach (HwndModel item in hwnds)
                {
                    if (caseCtrl[key].Class == item.Class && caseCtrl[key].X == item.ExeX && caseCtrl[key].Y == item.ExeY)
                    {
                        _currHwnds.Add(key, item.HwndId);
                        break;
                    }
                }
            }

            // 
            foreach (string key in caseData.Keys)
            {
                if (key == "Sleep" || key == "Pic") continue;
                if (caseData[key].StartsWith("Click:"))
                {
                    DoClick(model, caseCtrl, key, caseData[key]);
                }
                else if (caseData[key] == "Click")
                {
                    DoClick(model, caseCtrl, key);
                }
                else if (caseData[key] == "Clear")
                {
                    DoInput(model, caseCtrl, key, "");
                }
                else if (caseData[key].StartsWith("Index:"))
                {
                    DoComBoBox(model, key, caseData[key]);
                }
                else if (caseData[key].StartsWith("Tab:"))
                {
                    DoTab(model, key, caseData[key].Split(':')[1]);
                }
                else if (caseData[key].StartsWith("Keys:"))
                {
                    DoInputKeys(model, caseCtrl, key, caseData[key].Split(':')[1]);
                }
                else
                {
                    DoInput(model, caseCtrl, key, caseData[key]);
                }
            }
        }

        /// <summary>
        /// 截图
        /// </summary>
        /// <param name="model"></param>
        /// <param name="hwnd"></param>
        /// <param name="order"></param>
        private void GetPicture(P102ViewModel model, IntPtr hwnd, OrderItem order)
        {
            // Focuse
            User32Utility.SetForegroundWindow(hwnd);
            Thread.Sleep(300);

            // Alt + PS
            User32Utility.keybd_event((byte)Keys.Menu, 0, 0x0, IntPtr.Zero);
            User32Utility.keybd_event((byte)0x2c, 0, 0x0, IntPtr.Zero);
            User32Utility.keybd_event((byte)0x2c, 0, 0x2, IntPtr.Zero);
            User32Utility.keybd_event((byte)Keys.Menu, 0, 0x2, IntPtr.Zero);
            Thread.Sleep(100);

            User32Utility.keybd_event((byte)Keys.Menu, 0, 0x0, IntPtr.Zero);
            User32Utility.keybd_event((byte)0x2c, 0, 0x0, IntPtr.Zero);
            User32Utility.keybd_event((byte)0x2c, 0, 0x2, IntPtr.Zero);
            User32Utility.keybd_event((byte)Keys.Menu, 0, 0x2, IntPtr.Zero);
            Thread.Sleep(100);

            // Pic
            //IDataObject newObject = null;
            Bitmap newBitmap = null;
            //newObject = Clipboard.GetDataObject();
            if (Clipboard.ContainsImage())
            {
                newBitmap = (Bitmap)(Clipboard.GetImage().Clone());
                newBitmap.Save(_currentTCCasePath + @"\" + model.PicNum.ToString().PadLeft(3, '0') + "_" + order.Sheet + "_" + order.Order + ".png", ImageFormat.Png);
                model.PicNum = model.PicNum + 1;
            }
            LogUtility.WriteInfo($"EXE测试：截图[{_currentTCCasePath + @"\" + model.PicNum.ToString().PadLeft(3, '0') + "_" + order.Sheet + "_" + order.Order + ".png"}");
        }

        /// <summary>
        /// 左击
        /// </summary>
        /// <param name="model"></param>
        /// <param name="id"></param>
        /// <param name="value"></param>
        private void DoClick(P102ViewModel model, Dictionary<string, ControlInfo> caseCtrl, string id, string value = null)
        {
            int nums = 1;
            try
            {
                if (!string.IsNullOrWhiteSpace(value))
                    nums = int.Parse(value.Split(':')[1]);
            }
            catch (Exception ex)
            {
                Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"Click连击不明[{id}]", "OK", delegate () { }));
                LogUtility.WriteError($"EXE测试：Click连击不明[{id}]", ex);
                throw new Exception("Stop");
            }

            if (_currHwnds.ContainsKey(id))
            {
                HwndModel temp = HwndUtility.GetHwndModel(_currHwnds[id]);
                // Focuse
                User32Utility.SetForegroundWindow(_currHwnds[id]);
                Thread.Sleep(100);

                // Mouse Move
                User32Utility.SetCursorPos((int)(temp.DeskX + temp.Width / 2), (int)(temp.DeskY + temp.Height / 2));
            }
            else
            {
                User32Utility.SetCursorPos(_topHwndModel.DeskX + caseCtrl[id].X, _topHwndModel.DeskY + caseCtrl[id].Y);
            }
            Thread.Sleep(100);

            for (int i = 1; i <= nums; i++)
            {
                User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                Thread.Sleep(100);
                User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                Thread.Sleep(100);
            }
            LogUtility.WriteInfo($"EXE测试：[{id}]Click点击{nums}次");
        }

        /// <summary>
        /// 输入
        /// </summary>
        /// <param name="model"></param>
        /// <param name="caseCtrl"></param>
        /// <param name="id"></param>
        /// <param name="value"></param>
        private void DoInput(P102ViewModel model, Dictionary<string, ControlInfo> caseCtrl, string id, string value)
        {
            if (_currHwnds.ContainsKey(id))
            {
                // HwndModel
                HwndModel temp = HwndUtility.GetHwndModel(_currHwnds[id]);

                // Focuse
                User32Utility.SetForegroundWindow(_currHwnds[id]);
                Thread.Sleep(100);

                // Mouse Move
                User32Utility.SetCursorPos((int)(temp.ExeX + temp.Width / 2), (int)(temp.ExeY + temp.Height / 2));
                Thread.Sleep(100);

                // Mouse Click
                //SendKeys.SendWait(str);
                User32Utility.SendMessage(_currHwnds[id], User32Utility.WM_SETTEXT, 0, new StringBuilder(value));
                Thread.Sleep(100);
            }
            else
            {
                User32Utility.SetCursorPos(_topHwndModel.DeskX + caseCtrl[id].X, _topHwndModel.DeskY + caseCtrl[id].Y);
                User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                Thread.Sleep(100);
                User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                Thread.Sleep(100);

                SendKeys.SendWait(value);
            }
        }

        /// <summary>
        /// 输入
        /// </summary>
        /// <param name="model"></param>
        /// <param name="caseCtrl"></param>
        /// <param name="id"></param>
        /// <param name="value"></param>
        private void DoInputKeys(P102ViewModel model, Dictionary<string, ControlInfo> caseCtrl, string id, string value)
        {
            if (_currHwnds.ContainsKey(id))
            {
                // HwndModel
                HwndModel temp = HwndUtility.GetHwndModel(_currHwnds[id]);

                // Focuse
                User32Utility.SetForegroundWindow(_currHwnds[id]);
                Thread.Sleep(100);

                // Mouse Move
                User32Utility.SetCursorPos((int)(temp.ExeX + temp.Width / 2), (int)(temp.ExeY + temp.Height / 2));
                Thread.Sleep(100);
            }
            else
            {
                User32Utility.SetCursorPos(_topHwndModel.DeskX + caseCtrl[id].X, _topHwndModel.DeskY + caseCtrl[id].Y);
                User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                Thread.Sleep(100);
                User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                Thread.Sleep(100);
            }
            SendKeys.SendWait(value);
        }

        /// <summary>
        /// Tab
        /// </summary>
        /// <param name="model"></param>
        /// <param name="caseCtrl"></param>
        /// <param name="id"></param>
        /// <param name="value"></param>
        private void DoTab(P102ViewModel model, string id, string value)
        {
            int nums;
            if (int.TryParse(value, out nums))
            {
                if (!_currHwnds.ContainsKey(id))
                {
                    Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"控件[{id}]不存在", "OK", delegate () { }));
                    LogUtility.WriteError($"EXE测试：控件[{id}]不存在", null);
                    throw new Exception("Stop");
                }

                _ = User32Utility.SendMessage(_currHwnds[id], User32Utility.TCM_GETCURSEL, nums - 1, 0);
                User32Utility.SendMessage(_currHwnds[id], User32Utility.TCM_SETCURFOCUS, nums - 1, 0);
                User32Utility.SendMessage(_currHwnds[id], User32Utility.TCM_SETCURSEL, nums - 1, 0);
                Thread.Sleep(100);
            }
            else
            {
                Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"Tab的Index不明[{id}-{value}]", "OK", delegate () { }));
                LogUtility.WriteError($"EXE测试：Tab的Index不明[{id}-{value}]", null);
                throw new Exception("Stop");
            }
        }

        /// <summary>
        /// ComBoBox
        /// </summary>
        /// <param name="model"></param>
        /// <param name="caseCtrl"></param>
        /// <param name="id"></param>
        /// <param name="value"></param>
        private void DoComBoBox(P102ViewModel model, string id, string value)
        {
            if (!_currHwnds.ContainsKey(id))
            {
                Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"ComBobox控件[{id}]不存在", "OK", delegate () { }));
                LogUtility.WriteError($"EXE测试：ComBobox控件[{id}]不存在", null);
                throw new Exception("Stop");
            }

            string[] arr = value.Split(':');
            int num;
            if (arr.Length == 2)
            {
                if (int.TryParse(arr[1], out num))
                {
                    User32Utility.SendMessage(_currHwnds[id], User32Utility.CB_SETCURSEL, num - 1, 0);
                }
                else
                {
                    Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"[{id}]Index不明[{arr[1]}]", "OK", delegate () { }));
                    LogUtility.WriteError($"EXE测试：[{id}]Index不明[{arr[1]}]", null);
                    throw new Exception("Stop");
                }
            }
            else
            {
                if (int.TryParse(arr[2], out num))
                {
                    string idCbx = arr[1];
                    if (!_currHwnds.ContainsKey(idCbx))
                    {
                        Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"ComBobox对应控件[{idCbx}]不存在", "OK", delegate () { }));
                        LogUtility.WriteError($"EXE测试：ComBobox对应控件[{idCbx}]不存在", null);
                        throw new Exception("Stop");
                    }

                    // Click
                    HwndModel temp = HwndUtility.GetHwndModel(_currHwnds[id]);
                    User32Utility.SetForegroundWindow(_currHwnds[id]);
                    Thread.Sleep(100);
                    User32Utility.SetCursorPos((int)(temp.DeskX + temp.Width / 2), (int)(temp.DeskY + temp.Height / 2));
                    User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                    Thread.Sleep(100);
                    User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                    Thread.Sleep(100);

                    // Hwnd Model
                    HwndModel temp2 = HwndUtility.GetHwndModel(_currHwnds[idCbx]);
                    User32Utility.SetForegroundWindow(_currHwnds[idCbx]);
                    Thread.Sleep(100);
                    User32Utility.SetCursorPos((int)(temp2.DeskX + temp2.Width / 2), (int)(temp2.DeskY + model.CbxRowHeight * num - model.CbxRowHeight / 2));
                    Thread.Sleep(100);

                    // Mouse Click
                    User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                    Thread.Sleep(100);
                    User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                    Thread.Sleep(100);
                }
                else
                {
                    Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"[{id}]Index不明[{arr[2]}]", "OK", delegate () { }));
                    LogUtility.WriteError($"EXE测试：[{id}]Index不明[{arr[2]}]", null);
                    throw new Exception("Stop");
                }
            }
        }

        private string _currentDBCasePath;

        /// <summary>
        /// DB
        /// </summary>
        /// <param name="model"></param>
        /// <param name="caseNo"></param>
        /// <param name="order"></param>
        private void DoDBOrder(P102ViewModel model, string caseNo, OrderItem order)
        {
            if (model.ContinueOrder != null) return;

            _currentDBCasePath = model.OutResultPath + @"\DB\" + caseNo;
            if (!Directory.Exists(_currentTCCasePath))
                Directory.CreateDirectory(_currentTCCasePath);

            DBSheetInfo caseData = model.ExcelData.DBList[order.Sheet].CaseDatas[order.Order];

            Thread.Sleep(caseData.Sleep);

            SqlServerModel sqlModel = new SqlServerModel()
            {
                Server = caseData.Server,
                Db = caseData.DataBase,
                User = caseData.User,
                PassWord = caseData.PassWord,
            };

            using (SqlServerUtiltiy db = new SqlServerUtiltiy(sqlModel))
            {
                if (caseData.IsDown == "D")
                {
                    string file = _currentDBCasePath + @"\" + (string.IsNullOrWhiteSpace(caseData.FileName) ? caseData.Table + ".csv" : caseData.FileName);
                    string sql = string.IsNullOrEmpty(caseData.Sql) ? $"SELECT * FROM {caseData.Table}" : caseData.Sql;
                    DataTable dt = db.ExecuteTable(sql);

                    FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write);
                    StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default);
                    string data = "";

                    //write colData
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        data += dt.Columns[i].ColumnName.ToString();
                        if (i < dt.Columns.Count - 1)
                        {
                            data += ",";
                        }
                    }
                    sw.WriteLine(data);

                    //weite RowData
                    if (dt.Rows.Count > 0)
                    {
                        Console.WriteLine("start write rows data");
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            data = "";
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                data += dt.Rows[i][j].ToString() == "" ? "null" : dt.Rows[i][j].ToString();
                                if (j < dt.Columns.Count - 1)
                                {
                                    data += ",";
                                }
                            }
                            sw.WriteLine(data);
                        }
                    }
                    else
                    {
                        Console.WriteLine("data empty");
                    }

                    sw.Close();
                    fs.Close();
                }
                else
                {
                    db.ExecuteNonQuery(caseData.Sql);
                }
            }

        }

        #endregion

        #region 比较

        Dictionary<string, List<string>> _oldPicFiles;
        Dictionary<string, List<string>> _newPicFiles;
        Dictionary<string, List<string>> _oldDbFiles;
        Dictionary<string, List<string>> _newDbFiles;
        List<ResultRecord> _results;

        public void Compare(P102ViewModel model)
        {
            if (!Directory.Exists(model.ComparePath))
            {
                Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"比较路径不存在", "OK", delegate () { }));
                LogUtility.WarnInfo($"EXE测试：比较路径不存在{model.ComparePath}");
                return;
            }

            if (File.Exists(model.ComparePath + @"\比較結果.xls"))
            {
                try
                {
                    File.Delete(model.ComparePath + @"\比較結果.xls");
                }
                catch
                {
                    Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"删除失败[比較結果.xls]", "OK", delegate () { }));
                    LogUtility.WarnInfo($"EXE测试：删除失败[比較結果.xls]");
                    return;
                }
            }

            GetCompareFiles(model);

            using (ExcelUtility excel = new ExcelUtility(model.ComparePath + @"\比較結果.xls"))
            {
                ExcelWorksheet sh = excel.AddSheet("比較結果");
                sh.Cells.Style.Font.SetFromFont(new Font("ＭＳ Ｐゴシック", 11));
                sh.DefaultRowHeight = 18.75;

                List<string> cases = new List<string>();
                cases.AddRange(_oldPicFiles.Keys);
                cases.AddRange(_newPicFiles.Keys);
                cases.AddRange(_oldDbFiles.Keys);
                cases.AddRange(_newDbFiles.Keys);
                cases = cases.Distinct().ToList();
                cases.Sort();

                _results = new List<ResultRecord>();
                foreach (string caseNo in cases)
                {
                    if (_oldPicFiles.ContainsKey(caseNo) || _newPicFiles.ContainsKey(caseNo))
                        CompareTC(model, excel, caseNo);
                    if (_oldDbFiles.ContainsKey(caseNo) || _newDbFiles.ContainsKey(caseNo))
                        CompareDB(model, excel, caseNo);
                }

                SetResultSheet(model, sh);

                excel.Save();
            }

            Task.Factory.StartNew(() => App.MessageQueue.Enqueue($"比较完了", "OK", delegate () { }));
            LogUtility.WriteInfo($"EXE测试：比较完了");
        }

        private void GetCompareFiles(P102ViewModel model)
        {
            _oldPicFiles = new Dictionary<string, List<string>>();
            _newPicFiles = new Dictionary<string, List<string>>();
            _oldDbFiles = new Dictionary<string, List<string>>();
            _newDbFiles = new Dictionary<string, List<string>>();

            string currentPath = model.ComparePath + @"\" + SOLD + @"\Picture";
            if (Directory.Exists(currentPath))
                AddFiles(model, currentPath, _oldPicFiles);
            currentPath = model.ComparePath + @"\" + SOLD + @"\DB";
            if (Directory.Exists(currentPath))
                AddFiles(model, currentPath, _oldDbFiles);
            currentPath = model.ComparePath + @"\" + SNEW + @"\Picture";
            if (Directory.Exists(currentPath))
                AddFiles(model, currentPath, _newPicFiles);
            currentPath = model.ComparePath + @"\" + SNEW + @"\DB";
            if (Directory.Exists(currentPath))
                AddFiles(model, currentPath, _newDbFiles);
        }

        private void AddFiles(P102ViewModel model, string path, Dictionary<string, List<string>> list)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            foreach (DirectoryInfo caseDir in dir.GetDirectories())
            {
                if (caseDir.Name.StartsWith("Case"))
                {
                    List<string> files = new List<string>();
                    foreach (FileInfo file in caseDir.GetFiles())
                    {
                        files.Add(caseDir.Name + @"\" + file.Name);
                    }
                    list.Add(caseDir.Name, files);
                }
            }
        }

        private void CompareTC(P102ViewModel model, ExcelUtility excel, string caseNo)
        {
            ExcelWorksheet sh = excel.AddSheet("No" + int.Parse(caseNo.Replace("Case", "")));
            sh.Cells.Style.Font.SetFromFont(new Font("ＭＳ Ｐゴシック", 11));
            sh.DefaultRowHeight = 18.75;
            sh.Cells[1, 2].Value = SOLD;
            sh.Cells[1, 18].Value = SNEW;

            int row = 1;
            if (!_oldPicFiles.ContainsKey(caseNo))
            {
                foreach (string file in _newPicFiles[caseNo])
                {
                    row = AddPic(model, sh, "", file, row);
                }
            }
            else if (!_newPicFiles.ContainsKey(caseNo))
            {
                foreach (string file in _oldPicFiles[caseNo])
                {
                    row = AddPic(model, sh, file, "", row);
                }
            }
            else
            {
                List<string> fls = new List<string>();
                fls.AddRange(_oldPicFiles[caseNo]);
                fls.AddRange(_newPicFiles[caseNo]);
                fls = fls.Distinct().ToList();
                fls.Sort();
                foreach (string f in fls)
                {
                    if (_oldPicFiles[caseNo].Contains(f) && _newPicFiles[caseNo].Contains(f))
                        row = AddPic(model, sh, f, f, row);
                    else if (_oldPicFiles.ContainsKey(f))
                        row = AddPic(model, sh, f, "", row);
                    else
                        row = AddPic(model, sh, "", f, row);
                }
            }
        }

        private int AddPic(P102ViewModel model, ExcelWorksheet sh, string picOld, string picNew, int row)
        {
            int picHeight;
            int rowHeight = sh.GetHeightPix();

            ResultRecord item = new ResultRecord()
            {
                Sold = picOld,
                Snew = picNew,
                Link = $"#'{sh.Name}'!" + sh.Cells[row, 1].Address,
            };
            if (string.IsNullOrEmpty(picOld) || string.IsNullOrEmpty(picNew))
            {
                if (!string.IsNullOrEmpty(picOld))
                {
                    ExcelPicture picO = sh.Drawings.AddPicture(SOLD + picOld, new FileInfo(model.ComparePath + @"\" + SOLD + @"\Picture\" + picOld));
                    picO.SetPosition(row, 0, 2, 0);
                    picHeight = picO.Image.Height;
                }
                else
                {
                    ExcelPicture picN = sh.Drawings.AddPicture(SNEW + picNew, new FileInfo(model.ComparePath + @"\" + SNEW + @"\Picture\" + picNew));
                    picN.SetPosition(row, 0, 16, 0);
                    picHeight = picN.Image.Height;
                }

                item.Sresult = "NG";
            }
            else
            {
                int colWidth = sh.GetWidthPix();
                string fileOld = model.ComparePath + @"\" + SOLD + @"\Picture\" + picOld;
                string fileNew = model.ComparePath + @"\" + SNEW + @"\Picture\" + picNew;
                ExcelPicture picO = sh.Drawings.AddPicture(SOLD + picOld, new FileInfo(fileOld));
                ExcelPicture picN = sh.Drawings.AddPicture(SNEW + picNew, new FileInfo(fileNew));

                picO.SetPosition(row, 0, 1, 0);
                picN.SetPosition(row, 0, 16, 0);
                picHeight = picO.Image.Height > picN.Image.Height ? picO.Image.Height : picN.Image.Height;

                List<Rectangle> points = PictureUtility.Compare(fileOld, fileNew, new Margins(model.PicLeft, model.PicRight, model.PicTop, model.PicButtom));
                if (points != null)
                {
                    int spsCnt = 1;
                    foreach (Rectangle rag in points)
                    {
                        ExcelShape spsOld = sh.AddShape((spsCnt++) + picOld);
                        spsOld.SetSize(rag.Width + 10, rag.Height + 10);
                        spsOld.SetPosition((int)(picO.From.Row * rowHeight + rag.Top - 5), (int)(picO.From.Column * colWidth + rag.Left - 5));

                        ExcelShape spsNew = sh.AddShape((spsCnt++) + picNew);
                        spsNew.SetSize(rag.Width + 10, rag.Height + 10);
                        spsNew.SetPosition((int)(picN.From.Row * rowHeight + rag.Top - 5), (int)(picN.From.Column * colWidth + rag.Left - 5));
                    }

                    item.Sresult = "NG";
                }
                else
                {
                    item.Sresult = "OK";
                }
            }
            _results.Add(item);

            int cnt = 0;
            do
            {
            } while (rowHeight * cnt++ < picHeight);
            row += cnt;

            return row;
        }

        private void CompareDB(P102ViewModel model, ExcelUtility excel, string caseNo)
        {
            ExcelWorksheet sh = excel.AddSheet("DB" + int.Parse(caseNo.Replace("Case", "")));
            sh.Cells.Style.Font.SetFromFont(new Font("ＭＳ Ｐゴシック", 11));
            sh.DefaultRowHeight = 18.75;
            sh.Cells.Style.Numberformat.Format = "@";

            int row = 1;
            if (!_oldDbFiles.ContainsKey(caseNo))
            {
                foreach (string file in _oldDbFiles[caseNo])
                {
                    row = AddFile(model, sh, "", file, row);
                }
            }
            else if (!_newDbFiles.ContainsKey(caseNo))
            {
                foreach (string file in _newDbFiles[caseNo])
                {
                    row = AddFile(model, sh, file, "", row);
                }
            }
            else
            {
                List<string> fls = new List<string>();
                fls.AddRange(_oldDbFiles[caseNo]);
                fls.AddRange(_newDbFiles[caseNo]);
                fls.Distinct().ToList().Sort();

                foreach (string file in fls)
                {
                    if (_oldDbFiles[caseNo].Contains(file) && _newDbFiles[caseNo].Contains(file))
                        row = AddFile(model, sh, file, file, row);
                    else if (_oldDbFiles.ContainsKey(file))
                        row = AddFile(model, sh, file, file, row);
                    else
                        row = AddFile(model, sh, file, file, row);
                }
            }
        }

        private int AddFile(P102ViewModel model, ExcelWorksheet sh, string filOld, string filNew, int row)
        {
            ResultRecord record = new ResultRecord() { Sold = filOld, Snew = filNew, Link = $"#'{sh.Name}'!" + sh.Cells[row, 1].Address, };

            if (string.IsNullOrWhiteSpace(filOld))
            {
                sh.Cells[row, 1].Value = filNew;
                sh.Row(row++).SetColor(Color.Yellow);
                sh.LoadData(model.ComparePath + @"\" + SNEW + @"\DB\" + filNew, row, 1);

                record.Sresult = "NG";
            }
            else if (string.IsNullOrWhiteSpace(filNew))
            {
                sh.Cells[row, 1].Value = filOld;
                sh.Row(row++).SetColor(Color.Yellow);
                sh.LoadData(model.ComparePath + @"\" + SOLD + @"\DB\" + filOld, row, 1);

                record.Sresult = "NG";
            }
            else
            {
                // 文件名
                sh.Cells[row, 1].Value = filNew;
                sh.Row(row++).SetColor(Color.Yellow);
                // 旧数据
                int strOld = row;
                int cntOld = sh.LoadData(model.ComparePath + @"\" + SOLD + @"\DB\" + filOld, row++, 1);
                row = sh.GetMaxRow() + 1;
                // 新数据
                int strNew = row;
                int cntNew = sh.LoadData(model.ComparePath + @"\" + SNEW + @"\DB\" + filNew, row++, 1);
                row = sh.GetMaxRow() + 1;

                int colCnt = sh.GetMaxColumn(strOld);

                // 比较结果
                // Head
                sh.Cells[strOld, 1, strOld, colCnt].Copy(sh.Cells[row, 1, row, colCnt]);
                row++;

                sh.Cells[row, 1].Formula = $"=IF(A{strOld + 1}=A{strNew + 1},TRUE,FALSE)";
                for (int i = 2; i <= colCnt; i++)
                    sh.Cells[row, 1].Copy(sh.Cells[row, i]);
                for (int i = row + 1; i < row + (cntOld > cntNew ? cntOld : cntNew) - 1; i++)
                    sh.Cells[row, 1, row, colCnt].Copy(sh.Cells[i, 1, i, colCnt]);

                var cond = sh.ConditionalFormatting.AddEqual(new ExcelAddress(sh.Cells[row, 1, row + cntOld - 2, colCnt].Address));
                cond.Formula = "FALSE";
                cond.Style.Font.Color.Color = Color.Red;
                cond.Style.Fill.BackgroundColor.Color = Color.LightPink;

                record.Sresult = "OK";
                for (int i = row; i < row + (cntOld > cntNew ? cntOld : cntNew); i++)
                {
                    for (int j = 1; j <= colCnt; j++)
                        if (sh.Cells[i, j].Text == "FALSE")
                        {
                            record.Sresult = "NG";
                            break;
                        }
                    if (record.Sresult == "NG") break;
                }
            }

            _results.Add(record);
            return sh.GetMaxRow() + 2;
        }

        private void SetResultSheet(P102ViewModel model, ExcelWorksheet sh)
        {
            sh.Cells["B2"].Formula = $"=\"差異総行数：\" & COUNTIF(F6:F65536,\"NG\")";
            sh.Cells["B3"].Formula = $"=\"削除行数：\" & (COUNTA(F6:F65536)-COUNTA(E6:E65536))";
            sh.Cells["B4"].Formula = $"=\"追加行数：\" & (COUNTA(F6:F65536)-COUNTA(C6:C65536))";

            sh.Cells["B5:G5"].SetRangeColor(Color.FromArgb(91, 155, 213));
            sh.Cells["B5:G5"].SetFontColor(Color.White);
            sh.Cells["B5:G5"].Style.Font.Bold = true;

            sh.Cells["C5"].Value = SOLD;
            sh.Cells["E5"].Value = SNEW;
            sh.Cells["F5"].Value = "判定";
            sh.Cells["G5"].Value = "備考";

            sh.Cells[5, 2, 5 + _results.Count, 7].SetRangeBorder();
            sh.Cells[5, 2, 5 + _results.Count, 7].Style.Border.BorderAround(ExcelBorderStyle.Medium);

            if (_results.Count > 0)
            {
                sh.Cells[6, 6, 5 + _results.Count, 6].Style.Font.UnderLine = true;
                sh.Cells[6, 6, 5 + _results.Count, 6].SetFontColor(Color.FromArgb(91, 155, 213));
                int row = 6;
                foreach (ResultRecord item in _results)
                {
                    sh.Cells[row, 2].Formula = $"=ROW()-5";
                    sh.Cells[row, 3].Value = item.Sold;
                    sh.Cells[row, 4].Formula = $"=ROW()-5";
                    sh.Cells[row, 5].Value = item.Snew;
                    sh.Cells[row, 6].Value = item.Sresult;
                    sh.Cells[row++, 6].Hyperlink = new Uri(item.Link, UriKind.Relative);
                }

                var cond = sh.ConditionalFormatting.AddEqual(new ExcelAddress(sh.Cells[6, 6, 5 + _results.Count, 6].Address));
                cond.Formula = "\"NG\"";
                cond.Style.Font.Color.Color = Color.Red;
                cond.Style.Fill.BackgroundColor.Color = Color.LightPink;

                sh.Cells[6, 2, 5 + _results.Count, 7].Style.Font.SetFromFont(new Font("ＭＳ Ｐゴシック", 11));
            }

            sh.Column(1).Width = 2.88;
            sh.Column(2).Width = 5.0;
            sh.Column(3).Width = 50.0;
            sh.Column(4).Width = 5.0;
            sh.Column(5).Width = 50.0;
            sh.Column(6).Width = 8.0;
            sh.Column(7).Width = 60.0;
        }

        #endregion
    }

    public struct ResultRecord
    {
        public string Sold;
        public string Snew;
        public string Sresult;
        public string Link;
    }
}
