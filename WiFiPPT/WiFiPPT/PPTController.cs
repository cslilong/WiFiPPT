using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

namespace WiFiPPT
{
    public class PPTController
    {
        // 定义PowerPoint应用程序对象
        static PPt.Application pptApplication;
        // 定义演示文稿对象
        static PPt.Presentation presentation;
        // 定义幻灯片集合对象
        static PPt.Slides slides;
        // 定义单个幻灯片对象
        static PPt.Slide slide;

        // 幻灯片的数量
        static int slidescount;
        // 幻灯片的索引
        static int slideIndex;

        // 获取ppt的句柄
        public static bool Init()
        {
            // 必须先运行幻灯片，下面才能获得PowerPoint应用程序，否则会出现异常
            // 获得正在运行的PowerPoint应用程序
            try
            {
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;
            }
            catch
            {
                return false;
            }
            if (pptApplication != null)
            {
                //获得演示文稿对象
                presentation = pptApplication.ActivePresentation;
                // 获得幻灯片对象集合
                slides = presentation.Slides;
                // 获得幻灯片的数量
                slidescount = slides.Count;
                // 获得当前选中的幻灯片
                try
                {
                    // 在普通视图下这种方式可以获得当前选中的幻灯片对象
                    // 然而在阅读模式下，这种方式会出现异常
                    slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                }
                catch
                {
                    // 在阅读模式下出现异常时，通过下面的方式来获得当前选中的幻灯片对象
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
            return true;
        }
        // First
        public static void First()
        {
            try
            {
                // 在普通视图中调用Select方法来选中第一张幻灯片
                slides[1].Select();
                slide = slides[1];
            }
            catch
            {
                // 在阅读模式下使用下面的方式来切换到第一张幻灯片
                pptApplication.SlideShowWindows[1].View.First();
                slide = pptApplication.SlideShowWindows[1].View.Slide;
            }
        }
        // Last
        public static void Last()
        {
            try
            {
                slides[slidescount].Select();
                slide = slides[slidescount];
            }
            catch
            {
                // 在阅读模式下使用下面的方式来切换到最后幻灯片
                pptApplication.SlideShowWindows[1].View.Last();
                slide = pptApplication.SlideShowWindows[1].View.Slide;
            }
        }
        // Next
        public static void Next()
        {
            slideIndex = slide.SlideIndex + 1;
            if (slideIndex > slidescount)
            {
                //MessageBox.Show("已经是最后一页了");
            }
            else
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    // 在阅读模式下使用下面的方式来切换到下一张幻灯片
                    pptApplication.SlideShowWindows[1].View.Next();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
        }
        // Prev
        public static void Prev()
        {
            slideIndex = slide.SlideIndex - 1;
            if (slideIndex >= 1)
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    // 在阅读模式下使用下面的方式来切换到上一张幻灯片
                    pptApplication.SlideShowWindows[1].View.Previous();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
            else
            {
                //MessageBox.Show("已经是第一页了");
            }
        }
    }
}
