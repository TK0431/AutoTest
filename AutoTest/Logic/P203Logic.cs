using AutoTest.ViewModels;
using FrameWork.Consts;
using FrameWork.Models;
using FrameWork.Utility;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

namespace AutoTest.Logic
{
    public class P203Logic
    {
        public void Init(P203ViewModel model)
        {
            model.Types = new System.Collections.ObjectModel.ObservableCollection<string> {
                "$id","$name","$link","$linkpart","$tag","$cssselect","$xpath",
                "$frame","$frameout"
            };
        }

        public void StartButton(P203ViewModel model)
        {
            if (string.IsNullOrWhiteSpace(model.Url))
                App.ShowMessage("URL未入力", "OK");
            else
            {
                if (model.Su == null) model.Su = new SeleniumUtility(new SeleniumScriptModel());

                model.Su.DoUrl(new SeleniumEvent() { Value = model.Url });
            }
        }

        public void AddButton(P203ViewModel model)
        {
            int max = model.Items.Count() > 0 ? model.Items.Max(x => x.Id) : 0;
            P203ItemViewModel item = new P203ItemViewModel()
            {
                Id = max + 1
            };

            model.Items.Add(item);
        }

        public void ShowButton(P203ViewModel model)
        {
            if (model.Su == null) return;

            string pic = @"D:\test_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".jpg";
            model.Su.SavePic(pic);

            List<PicRectModel> list = new List<PicRectModel>();
            foreach (P203ItemViewModel item in model.Items)
            {
                if (string.IsNullOrWhiteSpace(item.Value)) continue;

                IWebElement element = null;
                try
                {
                    element = model.Su.GetElement(item.Type, item.Value);
                }
                catch (NoSuchElementException)
                { }

                if (element == null) continue;

                PicRectModel rect = new PicRectModel()
                {
                    X = element.Location.X - 5,
                    Y = element.Location.Y - 5,
                    Width = element.Size.Width + 10,
                    Height = element.Size.Height + 10,
                    Color = "#FF0000",
                    Thickness = 2,
                };
                list.Add(rect);
            }
            OpenCvUtility.AddRects(pic, list);

            model.Image = new BitmapImage(new Uri(pic));

            App.Win_top();
        }
    }
}
