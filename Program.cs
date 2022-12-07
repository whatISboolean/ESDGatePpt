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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Office2016.Drawing.Charts;
using System.IO;

namespace ESDGatePpt
{
    internal class Program
    {
        static uint uniqueId;
        static async Task Main(string[] args)
        {
            {
                //capture and download esd gate image
                var imageFolder = @"C:\esd_gate_pc_Shared\esd_gate_img";
                //var imageFolder = @"C:\Users\user\Desktop\esd_gate_img";

                await GenerateESDGateGraph(imageFolder);
                //close running powerpoint
                string filepath = $@"C:\esd_gate_pc_Shared\ESDGate_PowerPoint.pptx";
               // string filepath = $@"C:\Users\user\Desktop\ESDGate_PowerPoint.pptx";

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

                //COMMON
                string imgpath5 = $@"{imageFolder}\5_ESDChart.PNG";
                ISlide slide5 = ppt.Slides[4];
                IImageData image5 = ppt.Images.Append(Image.FromFile(imgpath5));
                foreach (IShape shape in slide5.Shapes)
                {
                    if (shape is SlidePicture)
                    {
                        (shape as SlidePicture).PictureFill.Picture.EmbedImage = image5;
                    }
                }

                //PC
                string imgpath6 = $@"{imageFolder}\6_ESDChart.PNG";
                ISlide slide6 = ppt.Slides[5];
                IImageData image6 = ppt.Images.Append(Image.FromFile(imgpath6));
                foreach (IShape shape in slide6.Shapes)
                {
                    if (shape is SlidePicture)
                    {
                        (shape as SlidePicture).PictureFill.Picture.EmbedImage = image6;
                    }
                }

                //QC
                string imgpath7 = $@"{imageFolder}\7_ESDChart.PNG";
                ISlide slide7 = ppt.Slides[6];
                IImageData image7 = ppt.Images.Append(Image.FromFile(imgpath7));
                foreach (IShape shape in slide7.Shapes)
                {
                    if (shape is SlidePicture)
                    {
                        (shape as SlidePicture).PictureFill.Picture.EmbedImage = image7;
                    }
                }

                //INSTRUMENT
                string imgpath8 = $@"{imageFolder}\8_ESDChart.PNG";
                ISlide slide8 = ppt.Slides[7];
                IImageData image8 = ppt.Images.Append(Image.FromFile(imgpath8));
                foreach (IShape shape in slide8.Shapes)
                {
                    if (shape is SlidePicture)
                    {
                        (shape as SlidePicture).PictureFill.Picture.EmbedImage = image8;
                    }
                }

                //DT
                string imgpath9 = $@"{imageFolder}\9_ESDChart.PNG";
                ISlide slide9 = ppt.Slides[8];
                IImageData image9 = ppt.Images.Append(Image.FromFile(imgpath9));
                foreach (IShape shape in slide9.Shapes)
                {
                    if (shape is SlidePicture)
                    {
                        (shape as SlidePicture).PictureFill.Picture.EmbedImage = image9;
                    }
                }

                //ET
                string imgpath10 = $@"{imageFolder}\10_ESDChart.PNG";
                ISlide slide10 = ppt.Slides[9];
                IImageData image10 = ppt.Images.Append(Image.FromFile(imgpath10));
                foreach (IShape shape in slide10.Shapes)
                {
                    if (shape is SlidePicture)
                    {
                        (shape as SlidePicture).PictureFill.Picture.EmbedImage = image10;
                    }
                }

                ppt.SaveToFile(filepath, FileFormat.Pptx2013);



                string mergedPresentation = "Display_Powerpoint.pptx";
                string presentationTemplate = "Display_Template.pptx";
                string[] sourcePresentations = new string[]
                { "ESDGate_PowerPoint.pptx","QC_Slide.pptx"};
                string presentationFolder = $@"C:\esd_gate_pc_Shared\";
                //string presentationFolder = $@"C:\Users\user\Desktop\";

                // Make a copy of the template presentation. This will throw an
                // exception if the template presentation does not exist.
                File.Copy(presentationFolder + presentationTemplate,
                  presentationFolder + mergedPresentation, true);

                // Loop through each source presentation and merge the slides 
                // into the merged presentation.
                foreach (string sourcePresentation in sourcePresentations)
                    MergeSlides(presentationFolder, sourcePresentation,
                      mergedPresentation);

                // Validate the merged presentation.
                OpenXmlValidator validator = new OpenXmlValidator();

                //var errors =
                //  validator.Validate(presentationFolder + mergedPresentation);

                var mergePptPath = presentationFolder + mergedPresentation;




            //Open powerpoint in slideshow mode 
            Microsoft.Office.Interop.PowerPoint.Application pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
                pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                pptApp.Activate();
                Microsoft.Office.Interop.PowerPoint.Presentations ps = pptApp.Presentations;
                Microsoft.Office.Interop.PowerPoint.Presentation p = ps.Open(mergePptPath,
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
            var processList = new List<string> { "SML%20DWB", "SML%20MOLD", "SML%20FINAL", "SML%20FINAL%20INSPECTION", "SML%20COMMON", "SML%20PC", "SML%20QC", "SML%20INSTRUMENT", "SML%20DEV.%20TECH", "SML%20EQUIP.%20TECH" };
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


        static void MergeSlides(string presentationFolder,string sourcePresentation, string destPresentation)
        {
            int id = 0;

            // Open the destination presentation.
            using (PresentationDocument myDestDeck =
              PresentationDocument.Open(presentationFolder + destPresentation,
              true))
            {
                PresentationPart destPresPart = myDestDeck.PresentationPart;

                // If the merged presentation does not have a SlideIdList 
                // element yet, add it.
                if (destPresPart.Presentation.SlideIdList == null)
                    destPresPart.Presentation.SlideIdList = new SlideIdList();

                // Open the source presentation. This will throw an exception if
                // the source presentation does not exist.
                using (PresentationDocument mySourceDeck =
                  PresentationDocument.Open(
                    presentationFolder + sourcePresentation, false))
                {
                    PresentationPart sourcePresPart =
                      mySourceDeck.PresentationPart;

                    // Get unique ids for the slide master and slide lists
                    // for use later.
                    uniqueId =
                      GetMaxSlideMasterId(
                        destPresPart.Presentation.SlideMasterIdList);

                    uint maxSlideId =
                      GetMaxSlideId(destPresPart.Presentation.SlideIdList);

                    // Copy each slide in the source presentation, in order, to 
                    // the destination presentation.
                    foreach (SlideId slideId in
                      sourcePresPart.Presentation.SlideIdList)
                    {
                        SlidePart sp;
                        SlidePart destSp;
                        SlideMasterPart destMasterPart;
                        string relId;
                        SlideMasterId newSlideMasterId;
                        SlideId newSlideId;

                        // Create a unique relationship id.
                        id++;
                        sp =
                          (SlidePart)sourcePresPart.GetPartById(
                            slideId.RelationshipId);

                        relId =
                          sourcePresentation.Remove(
                            sourcePresentation.IndexOf('.')) + id;

                        // Add the slide part to the destination presentation.
                        destSp = destPresPart.AddPart<SlidePart>(sp, relId);

                        // The slide master part was added. Make sure the
                        // relationship between the main presentation part and
                        // the slide master part is in place.
                        destMasterPart = destSp.SlideLayoutPart.SlideMasterPart;
                        destPresPart.AddPart(destMasterPart);

                        // Add the slide master id to the slide master id list.
                        uniqueId++;
                        newSlideMasterId = new SlideMasterId();
                        newSlideMasterId.RelationshipId =
                          destPresPart.GetIdOfPart(destMasterPart);
                        newSlideMasterId.Id = uniqueId;

                        destPresPart.Presentation.SlideMasterIdList.Append(
                          newSlideMasterId);

                        // Add the slide id to the slide id list.
                        maxSlideId++;
                        newSlideId = new SlideId();
                        newSlideId.RelationshipId = relId;
                        newSlideId.Id = maxSlideId;

                        destPresPart.Presentation.SlideIdList.Append(newSlideId);
                    }

                    // Make sure that all slide layout ids are unique.
                    FixSlideLayoutIds(destPresPart);
                }

                // Save the changes to the destination deck.
                destPresPart.Presentation.Save();
            }
        }

        static void FixSlideLayoutIds(PresentationPart presPart)
        {
            // Make sure that all slide layouts have unique ids.
            foreach (SlideMasterPart slideMasterPart in
              presPart.SlideMasterParts)
            {
                foreach (SlideLayoutId slideLayoutId in
                  slideMasterPart.SlideMaster.SlideLayoutIdList)
                {
                    uniqueId++;
                    slideLayoutId.Id = (uint)uniqueId;
                }

                slideMasterPart.SlideMaster.Save();
            }
        }

        static uint GetMaxSlideId(SlideIdList slideIdList)
        {
            // Slide identifiers have a minimum value of greater than or
            // equal to 256 and a maximum value of less than 2147483648. 
            uint max = 256;

            if (slideIdList != null)
                // Get the maximum id value from the current set of children.
                foreach (SlideId child in slideIdList.Elements<SlideId>())
                {
                    uint id = child.Id;

                    if (id > max)
                        max = id;
                }

            return max;
        }

        static uint GetMaxSlideMasterId(SlideMasterIdList slideMasterIdList)
        {
            // Slide master identifiers have a minimum value of greater than
            // or equal to 2147483648. 
            uint max = 2147483648;

            if (slideMasterIdList != null)
                // Get the maximum id value from the current set of children.
                foreach (SlideMasterId child in
                  slideMasterIdList.Elements<SlideMasterId>())
                {
                    uint id = child.Id;

                    if (id > max)
                        max = id;
                }

            return max;
        }

        static void DisplayValidationErrors(IEnumerable<ValidationErrorInfo> errors)
        {
            int errorIndex = 1;

            foreach (ValidationErrorInfo errorInfo in errors)
            {
                Console.WriteLine(errorInfo.Description);
                Console.WriteLine(errorInfo.Path.XPath);

                if (++errorIndex <= errors.Count())
                    Console.WriteLine("================");
            }

        }

    }
}
