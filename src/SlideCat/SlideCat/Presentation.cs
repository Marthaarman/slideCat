using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Presentation = Microsoft.Office.Interop.PowerPoint.Presentation;
using Slide = Microsoft.Office.Interop.PowerPoint.Slide;

namespace SlideCat
{
    public class SlideCatPresentation
    {
        private static uint uniqueId;

        private Application _application;
        private int _currentSlideIndex;

        private int _intervalCounter = 0;

        private readonly string _pptPath = "";
        private Presentation _presentation;

        private readonly string _slideCatPath = Path.GetTempPath() + "slidecat\\";

        public SlideCatPresentation()
        {
            Console.WriteLine(_slideCatPath);
            if (!Directory.Exists(_slideCatPath)) Directory.CreateDirectory(_slideCatPath);
            _pptPath = _slideCatPath + new Random().Next() + "\\";

            if (!Directory.Exists(_pptPath)) Directory.CreateDirectory(_pptPath);
        }

        public string slideNotes { get; } = string.Empty;

        public string slideNotesNext { get; } = string.Empty;

        public bool IsPlaying { get; private set; }

        public bool stopping { get; private set; }

        private void _emptyPresentationDirectory()
        {
            if (!IsPlaying)
            {
                DirectoryInfo di = new DirectoryInfo(_pptPath);

                foreach (FileInfo file in di.GetFiles())
                    try
                    {
                        file.Delete();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                foreach (DirectoryInfo dir in di.GetDirectories())
                    try
                    {
                        dir.Delete(true);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
            }
        }

        public void createPresentation(MediaItems mediaItems, ref BackgroundWorker worker)
        {
            //  the destination powerpoint
            string _destinationPowerPoint = "destinationPowerPoint.pptx";

            //  clear the folder to which temporary files are stored
            _emptyPresentationDirectory();

            //  initiate new application and main presentation

            Application powerPointApplication = new Application();
            Presentation powerPointPresentation = powerPointApplication.Presentations.Add(MsoTriState.msoFalse);

            //  insert first slide, make black
            CustomLayout customLayout = powerPointPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];
            Slide slide = powerPointPresentation.Slides.AddSlide(1, customLayout);
            Color slideBGColor = Color.Black;
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.ForeColor.RGB = slideBGColor.ToArgb();
            powerPointPresentation.SaveCopyAs(_pptPath + "destinationPowerPoint", PpSaveAsFileType.ppSaveAsDefault,
                MsoTriState.msoTrue);
            powerPointPresentation.Close();

            //  save each presentation as powerpoint presentation into the tmp folder
            //  add each temporary powerpoint into the main powerpoint
            mediaItems.sort();
            int i = 0;
            int nrItems = mediaItems.mediaItems.Count;
            foreach (MediaItem mediaItem in mediaItems.mediaItems)
            {
                //  indicate the nth powerpoint to be converting
                i++;

                //  report progress to the backgroundworker
                double percentageDouble = i * 100 / (nrItems + 1);
                percentageDouble = Math.Round(percentageDouble);
                int percentageInt = (int)percentageDouble;
                worker.ReportProgress(percentageInt);

                //  store each powerpoint to a file
                Presentation pres = mediaItem.presentation;
                pres.SaveCopyAs(_pptPath + "temporary_pptx_" + i, PpSaveAsFileType.ppSaveAsDefault,
                    MsoTriState.msoTrue);
                pres.Close();

                //  merge into the destination powerpoint
                MergeSlides(_pptPath, "temporary_pptx_" + i + ".pptx", _destinationPowerPoint);

                //  delete temporary powerpoint
                File.Delete(_pptPath + "temporary_pptx_" + i + ".pptx");
            }

            _application = new Application();
            _presentation = _application.Presentations.Open2007(_pptPath + _destinationPowerPoint, MsoTriState.msoTrue,
                MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
        }

        public void playPresentation()
        {
            IsPlaying = false;

            //  must have at least one slide to continue
            if (!(_presentation.Slides.Count > 0)) return;

            //  load the powerpoint in preview mode and run
            SlideShowSettings settings = _presentation.SlideShowSettings;
            settings.ShowType = (PpSlideShowType)1;
            settings.ShowPresenterView = MsoTriState.msoTrue;
            SlideShowWindow sw = settings.Run();

            //  go to first slide
            _presentation.SlideShowWindow.View.GotoSlide(1);
            _presentation.SlideShowWindow.View.FirstAnimationIsAutomatic();
            stopping = false;

            //  add handler for when stopping
            _application.PresentationBeforeClose += delegate
            {
                stopping = true;
                IsPlaying = false;
            };

            //  set status to playing
            IsPlaying = true;
        }

        public void stopPresentation()
        {
            if (_presentationPlaying())
            {
                IsPlaying = false;
                _stopPresentation(_presentation);
            }
        }

        private void _stopPresentation(Presentation _presentation)
        {
            if (_presentationPlaying())
                try
                {
                    _presentation.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("LOG - Presentation.cs - _stopPresentation() - catch");
                    Console.WriteLine(ex.Message);
                }
        }


        public void stop()
        {
            if (_presentationPlaying())
            {
                stopPresentation();
                _emptyPresentationDirectory();
            }
        }

        private bool _presentationPlaying()
        {
            try
            {
                int tmp = _presentation.Slides.Count;
                return IsPlaying;
            }
            catch
            {
                return false;
            }
        }


        private static void MergeSlides(string presentationFolder, string sourcePresentation, string destPresentation)
        {
            int id = 0;
            Console.WriteLine(presentationFolder);
            Console.WriteLine(sourcePresentation);
            Console.WriteLine(destPresentation);
            // Open the destination presentation.
            using (PresentationDocument myDestDeck =
                   PresentationDocument.Open(presentationFolder + destPresentation, true))
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
                        GetMaxSlideMasterId(destPresPart.Presentation.SlideMasterIdList);

                    uint maxSlideId = GetMaxSlideId(destPresPart.Presentation.SlideIdList);

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
                        destSp = destPresPart.AddPart(sp, relId);

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

                        destPresPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);

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

        private static void FixSlideLayoutIds(PresentationPart presPart)
        {
            // Make sure that all slide layouts have unique ids.
            foreach (SlideMasterPart slideMasterPart in
                     presPart.SlideMasterParts)
            {
                foreach (SlideLayoutId slideLayoutId in
                         slideMasterPart.SlideMaster.SlideLayoutIdList)
                {
                    uniqueId++;
                    slideLayoutId.Id = uniqueId;
                }

                slideMasterPart.SlideMaster.Save();
            }
        }

        private static uint GetMaxSlideId(SlideIdList slideIdList)
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

        private static uint GetMaxSlideMasterId(SlideMasterIdList slideMasterIdList)
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

        private static void DisplayValidationErrors(
            IEnumerable<ValidationErrorInfo> errors)
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