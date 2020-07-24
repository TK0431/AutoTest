using AutoTest.ViewModels;
using FrameWork.Consts;
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
using System.Runtime.InteropServices;
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
            if (!Directory.Exists(Environment.CurrentDirectory + @"\ExcelScript\"))
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + @"\ExcelScript");
                Directory.CreateDirectory(Environment.CurrentDirectory + @"\ExcelScript\DB");
                LogUtility.WriteInfo($"Create directory [{Environment.CurrentDirectory + @"\ExcelScript\"}]");
            }

            DirectoryInfo di = new DirectoryInfo(Environment.CurrentDirectory + @"\ExcelScript\");
            di.GetFiles().ToList().ForEach(x => model.Files.Add(x));

            LogUtility.WriteInfo($"初始化完了");
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
            model.ContinueOrder = null;

            if (model.SelectedFile == null)
            {
                App.ShowMessage("文件未选择", "Check");
                return;
            }

            using (ExcelUtility excel = new ExcelUtility(model.SelectedFile.FullName))
            {
                ExcelWorksheet sh = excel.GetSheet("実行シナリオ");
                if (sh == null)
                {
                    App.ShowMessage("[実行シナリオ]不存在", "Error", EnumMessageType.Error);
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
                        App.ShowMessage($"Sheet[{shName}]不存在", "Error", EnumMessageType.Error);
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
                                    App.ShowMessage($"[{shName}]{i}行异常", "Check", EnumMessageType.Error);
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
                                        App.ShowMessage($"[{shName}]{i}列Sleep异常", "Check", EnumMessageType.Error);
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
                                    App.ShowMessage($"[{shName}]{i}列异常", "Check", EnumMessageType.Error);
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
                                    App.ShowMessage($"[{shName}]{i}行异常", "Check", EnumMessageType.Error);
                                    return;
                                }
                            }

                            model.ExcelData.DBList.Add(shName, db);
                            break;
                        default:
                            //App.ShowMessage($"[{shName}]请以TC,DB命名", "Check");
                            break;
                    }
                    LogUtility.WriteInfo($"EXE测试：[{shName}]获取完了");
                });
            }

            model.FlgStart = true;
            App.ShowMessage($"读取完毕");
        }

        #region 自动测试

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        public void Start(P102ViewModel model)
        {
            // 创建路径
            CreatePath(model);

            model.FlgNewDir = false;

            StartTest(model);

            App.ShowMessage($"测试完了");
        }

        public void Continue(P102ViewModel model)
        {
            StartTest(model);

            model.FlgNewDir = false;

            App.ShowMessage($"测试完了");
        }

        /// <summary>
        /// 创建路径
        /// </summary>
        /// <param name="model"></param>
        private void CreatePath(P102ViewModel model)
        {
            model.OutResultPath = Environment.CurrentDirectory + @"\Result";

            if (model.FlgNewDir && Directory.Exists(model.OutResultPath))
            {
                try
                {
                    Directory.Move(model.OutResultPath, model.OutResultPath + DateTime.Now.ToString("yyyyMMddHHmmss"));
                }
                catch
                {
                    App.ShowMessage($"Result文件夹重命名失败", "Error");
                    throw new Exception("Stop");
                }
            }

            if (!Directory.Exists(model.OutResultPath))
            {
                // Create Result path
                Directory.CreateDirectory(model.OutResultPath);
                LogUtility.WriteInfo($"EXE测试:路径创建[{model.OutResultPath}]");
            }

            if (model.FlgCodeOld)
            {
                model.OutResultPath = model.OutResultPath + @"\" + SOLD;
            }
            else
            {
                model.OutResultPath = model.OutResultPath + @"\" + SNEW;
            }
            Directory.CreateDirectory(model.OutResultPath);
            LogUtility.WriteInfo($"EXE测试:路径创建[{model.OutResultPath}]");
        }

        private void StartTest(P102ViewModel model)
        {
            string oldCaseNo = null;
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

                if (model.ContinueOrder == null && oldCaseNo != caseNo)
                    model.PicNum = 1;

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
                model.FlgStart = false;
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
                if (caseCtrl["1"].Class.StartsWith("Afx:"))
                {
                    List<HwndModel> list = HwndUtility.GetDeskHwndModels();
                    foreach (HwndModel hm in list)
                        if (hm.Class.StartsWith("Afx:"))
                        {
                            hwnd = hm.HwndId;
                            break;
                        }
                }

                if (hwnd == IntPtr.Zero)
                {
                    App.ShowMessage($"Exe不存在[{caseCtrl["1"].Name}-{caseCtrl["1"].Class}]", "Error", EnumMessageType.Error);
                    model.ContinueOrder = order;
                    model.FlgStart = false;
                    model.FlgContinue = true;
                    throw new Exception("Stop");
                }
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
                //else if (caseData[key].StartsWith("Copy:"))
                //{
                //    DoCopy(model, caseCtrl, key, caseData[key].Split(':')[1]);
                //}
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
                else if (caseData[key].StartsWith("Key:"))
                {
                    DoInputKeys(model, caseCtrl, key, caseData[key].Split(':')[1]);
                }
                else if (caseData[key] == "Enter" || caseData[key] == "Up" || caseData[key] == "Down" || caseData[key] == "Left" || caseData[key] == "Right" || caseData[key] == "PgUp" || caseData[key] == "PgDown")
                {
                    DoInputKey(model, caseData[key]);
                }
                //else if (caseData[key].StartsWith("Move"))
                //{
                //    DoMove(model, caseCtrl[key]);
                //}
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
            //Clipboard.Clear();

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
            IDataObject newObject = null;
            Bitmap newBitmap = null;
            newObject = Clipboard.GetDataObject();
            if (Clipboard.ContainsImage())
            {
                newBitmap = (Bitmap)(Clipboard.GetImage().Clone());
                newBitmap.Save(_currentTCCasePath + @"\" + model.PicNum.ToString().PadLeft(3, '0') + "_" + order.Sheet + "_" + order.Order + ".png", ImageFormat.Png);
                LogUtility.WriteInfo($"EXE测试：截图[{_currentTCCasePath + @"\" + model.PicNum++.ToString().PadLeft(3, '0') + "_" + order.Sheet + "_" + order.Order + ".png"}");
            }
            else
            {
                LogUtility.WriteError($"EXE测试：截图失败[{_currentTCCasePath + @"\" + model.PicNum++.ToString().PadLeft(3, '0') + "_" + order.Sheet + "_" + order.Order + ".png"}", null);
            }

            //Clipboard.Clear();
        }

        //private void DoMove(P102ViewModel model, ControlInfo ci)
        //{
        //    User32Utility.mouse_event(User32Utility.MOUSEEVENTF_MOVE, ci.X, ci.Y, 0, 0);
        //}

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
            catch
            {
                App.ShowMessage($"Click连击不明[{id}]", "Error", EnumMessageType.Error);
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
                LogUtility.WriteInfo($"EXE测试：[{id}]Click点击");
            }
        }

        //private void DoCopy(P102ViewModel model, Dictionary<string, ControlInfo> caseCtrl, string key, string value)
        //{
        //    //Clipboard.Clear();

        //    Clipboard.SetDataObject(value);

        //    User32Utility.SetCursorPos(_topHwndModel.DeskX + caseCtrl[key].X, _topHwndModel.DeskY + caseCtrl[key].Y);
        //    Thread.Sleep(100);

        //    User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
        //    Thread.Sleep(100);
        //    User32Utility.mouse_event(User32Utility.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
        //    Thread.Sleep(100);

        //    SendKeys.SendWait("^{v}");

        //    //Clipboard.Clear();

        //    LogUtility.WriteInfo($"EXE测试：[{key}]Copy{value}");
        //}

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

                foreach (char c in value)
                {
                    SendKeys.SendWait(c.ToString());
                    Thread.Sleep(model.KeySleep);
                }

            }

            LogUtility.WriteInfo($"EXE测试：[{id}]DoInput[{value}]");
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
            foreach (char c in value)
            {
                SendKeys.SendWait(c.ToString());
                Thread.Sleep(model.KeySleep);
            }

            LogUtility.WriteInfo($"EXE测试：[{id}]DoInputKeys[{value}]");
        }

        private void DoInputKey(P102ViewModel model, string value)
        {
            switch (value)
            {
                case "Enter":
                    SendKeys.SendWait("{ENTER}");
                    break;
                case "Up":
                    SendKeys.SendWait("{UP}");
                    break;
                case "Down":
                    SendKeys.SendWait("{DOWN}");
                    break;
                case "Left":
                    SendKeys.SendWait("{LEFT}");
                    break;
                case "Right":
                    SendKeys.SendWait("{RIGHT}");
                    break;
                case "PgUp":
                    SendKeys.SendWait("{PGUP}");
                    break;
                case "PgDown":
                    SendKeys.SendWait("{PGDN}");
                    break;
            }
            Thread.Sleep(100);

            LogUtility.WriteInfo($"EXE测试：DoInputKey[{value}]");
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
                    App.ShowMessage($"控件[{id}-{value}]不存在", "Error", EnumMessageType.Error);
                    throw new Exception("Stop");
                }

                _ = User32Utility.SendMessage(_currHwnds[id], User32Utility.TCM_GETCURSEL, nums - 1, 0);
                User32Utility.SendMessage(_currHwnds[id], User32Utility.TCM_SETCURFOCUS, nums - 1, 0);
                User32Utility.SendMessage(_currHwnds[id], User32Utility.TCM_SETCURSEL, nums - 1, 0);
                Thread.Sleep(100);
            }
            else
            {
                App.ShowMessage($"Tab的Index不明[{id}-{value}]", "Error", EnumMessageType.Error);
                throw new Exception("Stop");
            }

            LogUtility.WriteInfo($"EXE测试：[{id}]DoTab[{value}]");
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
                App.ShowMessage($"ComBobox控件[{id}-{value}]不存在", "Error", EnumMessageType.Error);
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
                    App.ShowMessage($"ComBobox[{id}]Index不明[{arr[1]}]", "Error", EnumMessageType.Error);
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
                        App.ShowMessage($"ComBobox对应控件[{id}-{value}]不存在", "Error", EnumMessageType.Error);
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
                    App.ShowMessage($"ComBobox[{id}]Index不明[{arr[2]}]", "Error", EnumMessageType.Error);
                    throw new Exception("Stop");
                }
            }

            LogUtility.WriteInfo($"EXE测试：[{id}]DoComBoBox[{value}]");
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
            if (!Directory.Exists(_currentDBCasePath))
                Directory.CreateDirectory(_currentDBCasePath);

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
                    List<string> data = new List<string>();

                    //write colData
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        data.Add("\"" + dt.Columns[i].ColumnName.ToString() + "\"");
                    }

                    sw.WriteLine(string.Join(",", data));

                    //weite RowData
                    if (dt.Rows.Count > 0)
                    {
                        Console.WriteLine("start write rows data");
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            data = new List<string>();
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                data.Add(dt.Rows[i][j].ToString() == "" ? "\"null\"" : "\"" + dt.Rows[i][j].ToString() + "\"");
                            }
                            sw.WriteLine(string.Join(",", data));
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
            LogUtility.WriteInfo($"EXE测试：Case[{caseNo}-{order.Order}]完了");
        }

        #endregion

        #region 比较

        Dictionary<string, List<string>> _oldPicFiles;
        Dictionary<string, List<string>> _newPicFiles;
        Dictionary<string, List<string>> _oldDbFiles;
        Dictionary<string, List<string>> _newDbFiles;
        Dictionary<string, List<string>> _oldFiles;
        Dictionary<string, List<string>> _newFiles;
        List<ResultRecord> _results;

        public void Compare(P102ViewModel model)
        {
            if (!Directory.Exists(model.ComparePath))
            {
                App.ShowMessage($"比较路径不存在", "Error", EnumMessageType.Error);
                return;
            }

            if (File.Exists(model.ComparePath + @"\比較結果.xlsx"))
            {
                try
                {
                    File.Delete(model.ComparePath + @"\比較結果.xlsx");
                }
                catch
                {
                    App.ShowMessage($"删除失败[比較結果.xls]", "Error", EnumMessageType.Error);
                    return;
                }
            }

            GetCompareFiles(model);

            using (ExcelUtility excel = new ExcelUtility(model.ComparePath + @"\比較結果.xlsx"))
            {
                ExcelWorksheet sh = excel.AddSheet("比較結果");
                sh.Cells.Style.Font.SetFromFont(new Font("ＭＳ Ｐゴシック", 11));
                sh.DefaultRowHeight = 13.5;
                sh.DefaultColWidth = 9.57;

                List<string> cases = new List<string>();
                cases.AddRange(_oldPicFiles.Keys);
                cases.AddRange(_newPicFiles.Keys);
                cases.AddRange(_oldDbFiles.Keys);
                cases.AddRange(_newDbFiles.Keys);
                cases.AddRange(_oldFiles.Keys);
                cases.AddRange(_newFiles.Keys);
                cases = cases.Distinct().ToList();
                cases.Sort();

                _results = new List<ResultRecord>();
                foreach (string caseNo in cases)
                {
                    if (_oldPicFiles.ContainsKey(caseNo) || _newPicFiles.ContainsKey(caseNo))
                        CompareTC(model, excel, caseNo);
                    if (_oldDbFiles.ContainsKey(caseNo) || _newDbFiles.ContainsKey(caseNo))
                        CompareDB(model, excel, caseNo);
                    if (_oldFiles.ContainsKey(caseNo) || _newFiles.ContainsKey(caseNo))
                        CompareFL(model, excel, caseNo);
                    GC.Collect();
                }

                SetResultSheet(model, sh);

                excel.Save();
            }

            App.ShowMessage($"比较完了");
        }

        private void GetCompareFiles(P102ViewModel model)
        {
            _oldPicFiles = new Dictionary<string, List<string>>();
            _newPicFiles = new Dictionary<string, List<string>>();
            _oldDbFiles = new Dictionary<string, List<string>>();
            _newDbFiles = new Dictionary<string, List<string>>();
            _oldFiles = new Dictionary<string, List<string>>();
            _newFiles = new Dictionary<string, List<string>>();

            string currentPath = model.ComparePath + @"\" + SOLD + @"\Picture";
            if (Directory.Exists(currentPath))
                AddFiles(model, currentPath, _oldPicFiles);
            currentPath = model.ComparePath + @"\" + SOLD + @"\DB";
            if (Directory.Exists(currentPath))
                AddFiles(model, currentPath, _oldDbFiles);
            currentPath = model.ComparePath + @"\" + SOLD + @"\File";
            if (Directory.Exists(currentPath))
                AddFiles(model, currentPath, _oldFiles);
            currentPath = model.ComparePath + @"\" + SNEW + @"\Picture";
            if (Directory.Exists(currentPath))
                AddFiles(model, currentPath, _newPicFiles);
            currentPath = model.ComparePath + @"\" + SNEW + @"\DB";
            if (Directory.Exists(currentPath))
                AddFiles(model, currentPath, _newDbFiles);
            currentPath = model.ComparePath + @"\" + SNEW + @"\File";
            if (Directory.Exists(currentPath))
                AddFiles(model, currentPath, _newFiles);
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
                        if (file.Extension == ".db") continue;
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
            sh.DefaultColWidth = 20;
            sh.DefaultRowHeight = 13.5;
            sh.Cells[1, 2].Value = SOLD;
            sh.Cells[1, 18].Value = SNEW;

            Enumerable.Range(1, 18).ToList().ForEach(c => sh.Column(c).Width = 10.3);

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
                    else if (_oldPicFiles[caseNo].Contains(f))
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

            LogUtility.WriteInfo($"EXE测试：PicCompare:[{picOld}-{picNew}]完了");

            return row;
        }

        private void CompareDB(P102ViewModel model, ExcelUtility excel, string caseNo)
        {
            ExcelWorksheet sh = excel.AddSheet("DB" + int.Parse(caseNo.Replace("Case", "")));
            sh.Cells.Style.Font.SetFromFont(new Font("ＭＳ Ｐゴシック", 11));
            sh.DefaultRowHeight = 13.5;
            sh.DefaultColWidth = 9.57;
            sh.Cells.Style.Numberformat.Format = "@";

            if (!_oldDbFiles.ContainsKey(caseNo))
            {
                foreach (string file in _oldDbFiles[caseNo])
                {
                    AddFile(model, sh, "", file);
                }
            }
            else if (!_newDbFiles.ContainsKey(caseNo))
            {
                foreach (string file in _newDbFiles[caseNo])
                {
                    AddFile(model, sh, file, "");
                }
            }
            else
            {
                List<string> fls = new List<string>();
                fls.AddRange(_oldDbFiles[caseNo]);
                fls.AddRange(_newDbFiles[caseNo]);
                fls = fls.Distinct().ToList();
                fls.Sort();

                foreach (string file in fls)
                {
                    if (_oldDbFiles[caseNo].Contains(file) && _newDbFiles[caseNo].Contains(file))
                        AddFile(model, sh, file, file);
                    else if (_oldDbFiles[caseNo].Contains(file))
                        AddFile(model, sh, file, "");
                    else
                        AddFile(model, sh, "", file);
                }
            }
        }

        private void CompareFL(P102ViewModel model, ExcelUtility excel, string caseNo)
        {
            ExcelWorksheet sh = excel.AddSheet("FL" + int.Parse(caseNo.Replace("Case", "")));
            sh.Cells.Style.Font.SetFromFont(new Font("ＭＳ Ｐゴシック", 11));
            sh.DefaultRowHeight = 13.5;
            sh.DefaultColWidth = 9.57;
            sh.Column(1).Width = 95.7;
            sh.Cells.Style.Numberformat.Format = "@";

            if (!_oldFiles.ContainsKey(caseNo))
            {
                foreach (string file in _oldFiles[caseNo])
                {
                    AddFileFL(model, sh, "", file);
                }
            }
            else if (!_newFiles.ContainsKey(caseNo))
            {
                foreach (string file in _newFiles[caseNo])
                {
                    AddFileFL(model, sh, file, "");
                }
            }
            else
            {
                List<string> fls = new List<string>();
                fls.AddRange(_oldFiles[caseNo]);
                fls.AddRange(_newFiles[caseNo]);
                fls = fls.Distinct().ToList();
                fls.Sort();

                foreach (string file in fls)
                {
                    if (_oldFiles[caseNo].Contains(file) && _newFiles[caseNo].Contains(file))
                        AddFileFL(model, sh, file, file);
                    else if (_oldFiles[caseNo].Contains(file))
                        AddFileFL(model, sh, file, "");
                    else
                        AddFileFL(model, sh, "", file);
                }
            }
        }

        private void AddFileFL(P102ViewModel model, ExcelWorksheet sh, string filOld, string filNew)
        {
            int row = sh.GetMaxRow(1);
            if (row > 1) row += 2;

            ResultRecord record = new ResultRecord() { Sold = filOld, Snew = filNew, Link = $"#'{sh.Name}'!" + sh.Cells[row, 1].Address, };

            if (string.IsNullOrWhiteSpace(filOld))
            {
                sh.Cells[row, 1].Value = filNew;
                sh.Row(row++).SetColor(Color.Yellow);
                sh.Cells[row, 1].Value = SOLD;
                sh.Row(row++).SetColor(Color.LightGreen);
                sh.Cells[row, 1].Value = "ファイル存在なし";
                sh.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sh.Cells[row++, 1].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                row++;
                sh.Cells[row, 1].Value = SNEW;
                sh.Row(row++).SetColor(Color.LightGreen);
                sh.LoadDataText(model.ComparePath + @"\" + SNEW + @"\File\" + filNew, row, 1);

                record.Sresult = "NG";
            }
            else if (string.IsNullOrWhiteSpace(filNew))
            {
                sh.Cells[row, 1].Value = filOld;
                sh.Row(row++).SetColor(Color.Yellow);
                sh.Cells[row, 1].Value = SOLD;
                sh.Row(row++).SetColor(Color.LightGreen);
                sh.LoadDataText(model.ComparePath + @"\" + SOLD + @"\File\" + filOld, row++, 1);
                row = sh.GetMaxRow(1) + 2;
                sh.Cells[row, 1].Value = SNEW;
                sh.Row(row++).SetColor(Color.LightGreen);
                sh.Cells[row, 1].Value = "ファイル存在なし";
                sh.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sh.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.LightPink);

                record.Sresult = "NG";
            }
            else
            {
                // 文件名
                sh.Cells[row, 1].Value = filNew;
                sh.Row(row++).SetColor(Color.Yellow);
                // 旧数据
                sh.Cells[row, 1].Value = SOLD;
                sh.Row(row++).SetColor(Color.LightGreen);
                int strOld = row;
                int cntOld = sh.LoadDataText(model.ComparePath + @"\" + SOLD + @"\File\" + filOld, row++, 1);
                if (cntOld == 0)
                {
                    sh.Cells[row - 1, 1].Value ="空ファイル";
                    sh.Cells[row - 1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                }
                row = sh.GetMaxRow(1) + 2;
                // 新数据
                sh.Cells[row, 1].Value = SNEW;
                sh.Row(row++).SetColor(Color.LightGreen);
                int strNew = row;
                int cntNew = sh.LoadDataText(model.ComparePath + @"\" + SNEW + @"\File\" + filNew, row++, 1);
                if (cntNew == 0)
                {
                    sh.Cells[row - 1, 1].Value = "空ファイル";
                    sh.Cells[row - 1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                }

                if(cntOld == 0 && cntNew != 0)
                    sh.Cells[cntOld, 1].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                if (cntOld != 0 && cntNew == 0)
                    sh.Cells[cntNew, 1].Style.Fill.BackgroundColor.SetColor(Color.LightPink);

                bool flg = true;

                int cnt = cntOld < cntNew ? cntOld : cntNew;
                if (cntOld != cntNew) flg = false;
                for (int i = 0; i < cnt; i++)
                {
                    if (sh.Cells[strOld + i, 1].Text != sh.Cells[strNew + i, 1].Text)
                    {
                        sh.Row(strOld + i).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sh.Row(strOld + i).Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                        sh.Row(strNew + i).Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sh.Row(strNew + i).Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                        flg = false;
                    }
                }
                for (int i = cnt; i < cntOld; i++)
                {
                    sh.Row(strOld + i).Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sh.Row(strOld + i).Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                }
                for (int i = cnt; i < cntNew; i++)
                {
                    sh.Row(strNew + i).Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sh.Row(strNew + i).Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                }

                record.Sresult = flg ? "OK" : "NG";
            }

            _results.Add(record);

            LogUtility.WriteInfo($"EXE测试：DBCompare:[{filOld}-{filNew}]完了");
        }

        private void AddFile(P102ViewModel model, ExcelWorksheet sh, string filOld, string filNew)
        {
            int row = sh.GetMaxRow(1);
            if (row > 1) row += 2;

            ResultRecord record = new ResultRecord() { Sold = filOld, Snew = filNew, Link = $"#'{sh.Name}'!" + sh.Cells[row, 1].Address, };

            if (string.IsNullOrWhiteSpace(filOld))
            {
                sh.Cells[row, 1].Value = filNew;
                sh.Row(row++).SetColor(Color.Yellow);
                sh.Cells[row, 1].Value = SOLD + " ファイル存在なし";
                sh.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sh.Cells[row++, 1].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                row++;
                sh.LoadData(model.ComparePath + @"\" + SNEW + @"\DB\" + filNew, row++, 1);

                record.Sresult = "NG";
            }
            else if (string.IsNullOrWhiteSpace(filNew))
            {
                sh.Cells[row, 1].Value = filOld;
                sh.Row(row++).SetColor(Color.Yellow);
                sh.LoadData(model.ComparePath + @"\" + SOLD + @"\DB\" + filOld, row, 1);
                row = sh.GetMaxRow(1) + 2;
                sh.Cells[row, 1].Value = SNEW + "  ファイル存在なし";
                sh.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sh.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.LightPink);

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
                if (cntOld == 0)
                {
                    sh.Cells[row - 1, 1].Value = SOLD + " 空ファイル";
                    sh.Cells[row - 1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sh.Cells[row - 1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                }
                row = sh.GetMaxRow(1) + 2;
                // 新数据
                int strNew = row;
                int cntNew = sh.LoadData(model.ComparePath + @"\" + SNEW + @"\DB\" + filNew, row++, 1);
                if (cntNew == 0)
                {
                    sh.Cells[row - 1, 1].Value = SNEW + " 空ファイル";
                    sh.Cells[row - 1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sh.Cells[row - 1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightPink);
                }
                row = sh.GetMaxRow(1) + 2;

                int colCnt = sh.GetMaxColumn(strOld);

                int cnt = cntOld > cntNew ? cntOld : cntNew;
                if (cnt == 1) cnt = 2;

                if (cntOld > 0 && cntNew > 0)
                {
                    // 比较结果
                    // Head
                    sh.Cells[strOld, 1, strOld, colCnt].Copy(sh.Cells[row, 1, row, colCnt]);
                    row++;

                    sh.Cells[row, 1].Formula = $"=IF(A{strOld + 1}=A{strNew + 1},TRUE,FALSE)";
                    for (int i = 2; i <= colCnt; i++)
                        sh.Cells[row, 1].Copy(sh.Cells[row, i]);
                    for (int i = row + 1; i < row + cnt - 1; i++)
                        sh.Cells[row, 1, row, colCnt].Copy(sh.Cells[i, 1, i, colCnt]);

                    var cond = sh.ConditionalFormatting.AddEqual(new ExcelAddress(sh.Cells[row, 1, row + cnt - 2, colCnt].Address));
                    cond.Formula = "FALSE";
                    cond.Style.Font.Color.Color = Color.Red;
                    cond.Style.Fill.BackgroundColor.Color = Color.LightPink;
                }
                record.Sresult = FileUtility.isValidFileContent(model.ComparePath + @"\" + SOLD + @"\DB\" + filOld, model.ComparePath + @"\" + SNEW + @"\DB\" + filNew) ? "OK" : "NG";
            }

            _results.Add(record);

            LogUtility.WriteInfo($"EXE测试：DBCompare:[{filOld}-{filNew}]完了");
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

                var cond = sh.ConditionalFormatting.AddExpression(new ExcelAddress(sh.Cells[6, 2, 5 + _results.Count, 7].Address));
                //cond.Formula = "\"NG\"";
                cond.Formula = "=$F6=\"NG\"";
                //cond.Style.Font.Color.Color = Color.Red;
                cond.Style.Fill.BackgroundColor.Color = Color.FromArgb(255, 255, 204);

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
