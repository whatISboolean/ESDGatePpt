using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System.Drawing;
using Microsoft.Office.Interop.PowerPoint;
using FluentDate;
using PuppeteerSharp;

namespace ESDGatePpt
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            {
                //var imageFolder = @"C:\Users\user\Desktop\esd_gate_img";
                //capture and download esd gate image
                var imageFolder = @"C:\esd_gate_pc_Shared\esd_gate_img";

                await GenerateESDGateGraph(imageFolder);
            //close running powerpoint
                //string filepath = $@"C:\Users\user\Desktop\ESDGate_PowerPoint.pptx";
                string filepath = $@"C:\esd_gate_pc_Shared\ESDGate_PowerPoint.pptx";

                Microsoft.Office.Interop.PowerPoint.Application pptOpen = new Microsoft.Office.Interop.PowerPoint.Application();
                pptOpen.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                pptOpen.Activate();
                pptOpen.Quit();
            //start process replace image of esd gate
                Spire.Presentation.Presentation ppt = new Spire.Presentation.Presentation();
                ppt.LoadFromFile(filepath);
                //DWB
                string imgpath1 = $@"{imageFolder}\1_ESDChart.PNG";
                ISlide slide = ppt.Slides[0];
                IImageData image = ppt.Images.Append(Image.FromFile(imgpath1));
                foreach (IShape shape in slide.Shapes)
                {
                    if (shape is SlidePicture)
                    {
                        //if (shape.AlternativeTitle == "image1")
                        //{
                            (shape as SlidePicture).PictureFill.Picture.EmbedImage = image;
                        //}
                    }
                }
                //MIDDLE
                string imgpath2 = $@"{imageFolder}\2_ESDChart.PNG";
                ISlide slide2 = ppt.Slides[1];
                IImageData image2 = ppt.Images.Append(Image.FromFile(imgpath2));
                foreach (IShape shape in slide2.Shapes)
                {
                    if (shape is SlidePicture)
                    {
                        (shape as SlidePicture).PictureFill.Picture.EmbedImage = image2;
                    }
                }
                //FINAL
                string imgpath3 = $@"{imageFolder}\3_ESDChart.PNG";
                ISlide slide3 = ppt.Slides[2];
                IImageData image3 = ppt.Images.Append(Image.FromFile(imgpath3));
                foreach (IShape shape in slide3.Shapes)
                {
                    if (shape is SlidePicture)
                    {
                        (shape as SlidePicture).PictureFill.Picture.EmbedImage = image3;
                    }
                }
                //INSPECTION
                string imgpath4 = $@"{imageFolder}\4_ESDChart.PNG";
                ISlide slide4 = ppt.Slides[3];
                IImageData image4 = ppt.Images.Append(Image.FromFile(imgpath4));
                foreach (IShape shape in slide4.Shapes)
                {
                    if (shape is SlidePicture)
                    {
                        (shape as SlidePicture).PictureFill.Picture.EmbedImage = image4;
                    }
                }

                ppt.SaveToFile(filepath, FileFormat.Pptx2013);




            //Open powerpoint in slideshow mode 
                Microsoft.Office.Interop.PowerPoint.Application pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
                pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                pptApp.Activate();
                Microsoft.Office.Interop.PowerPoint.Presentations ps = pptApp.Presentations;
                Microsoft.Office.Interop.PowerPoint.Presentation p = ps.Open(filepath,
                            Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue);

                Presentations ppPresens = pptApp.Presentations;
                Slides objSlides = p.Slides;

                SlideShowWindows objSSWs;
                SlideShowSettings objSSS;
                //Run the Slide show                                
                objSSS = p.SlideShowSettings;
                objSSS.Run();
                objSSWs = pptApp.SlideShowWindows;
                
            }

        }

        //Generate ESD Gate IN/OUT for process monitoring every day
        private static async Task GenerateESDGateGraph(string imageFolder)
        {
            //var imageFolder = @"C:\Users\SM11405\Desktop";
            var processList = new List<string> { "SML%20DWB", "SML%20MOLD", "SML%20FINAL", "SML%20FINAL%20INSPECTION" };
            var today = DateTime.Today;
            var graphWaitOption = new WaitForSelectorOptions { Timeout = 120000 };



            //Use PuppeterSharp to generate graph image----------------------------------------------------------
            var browserFetcher = new BrowserFetcher(new BrowserFetcherOptions { Path = AppDomain.CurrentDomain.BaseDirectory + @".local-chromium" });
            await browserFetcher.DownloadAsync(BrowserFetcher.DefaultChromiumRevision);
            var browser = await Puppeteer.LaunchAsync(new LaunchOptions
            {
                Headless = true,
                ExecutablePath = AppDomain.CurrentDomain.BaseDirectory + @"\.local-chromium\Win64-970485\chrome-win\chrome.exe",
            });



            var page = await browser.NewPageAsync();



            await page.SetViewportAsync(new ViewPortOptions { Height = 1200, Width = 2000 });



            //If first time/no cookies -> esd gate redirect to => http://10.30.1.171/esdgate/Division . So redirect to sml page first to create cookues
            await page.GoToAsync("http://10.30.1.171/esdgate/Home/SML");

            await page.GoToAsync("http://10.30.1.171/esdgate/Process");



            //Change date to yesterday
            await page.EvaluateExpressionAsync($"$('#seldate').val('{today.AddDays(-1).ToString("yyyy-MM-dd")}');");



            for (int i = 0; i < processList.Count; i++)
            {
                await page.SelectAsync("#selprocess", processList[i]);



                await page.WaitForTimeoutAsync(2000);



                await page.ClickAsync("button[type='submit']");



                await page.WaitForTimeoutAsync(2000);



                await page.WaitForSelectorAsync("#myChart", graphWaitOption);



                await page.WaitForTimeoutAsync(5000);



                var script =
                       " () => {"
                      + "var canvas = document.querySelector('#myChart');"
                      + $"var link = document.createElement('a'); link.setAttribute('download', '{(i + 1)}_ESDChart_{today.ToString("dd_MM_yyyy")}.png'); "
                      + "link.setAttribute('href', canvas.toDataURL('image/png').replace('image/png', 'image/octet-stream')); link.click(); "
                      + "return link.attributes.href.value;"
                      + "}";



                var imgStreamStr = await page.EvaluateFunctionAsync(script);
                byte[] byteArray = Convert.FromBase64String(imgStreamStr.ToString().Replace("data:image/octet-stream;base64,", ""));
                System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
                Image img = System.Drawing.Image.FromStream(stream);
                var imgSavePath = $@"{imageFolder}\{(i + 1)}_ESDChart.png";  //$@"{rootPath}\{folderName}\Graph\summaryChartSok_{today.ToString("dd_MM_yyyy")}.png";
                //var imgSavePath = $@"{imageFolder}\{(i + 1)}_ESDChart_{today.ToString("dd_MM_yyyy")}.png";  
                img.Save(imgSavePath);
            }
            await browser.CloseAsync();
        }


    }
}
