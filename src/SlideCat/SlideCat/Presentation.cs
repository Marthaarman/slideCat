using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.Drawing;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;

namespace SlideCat
{
    public class SlideCatPresentation
    {
        private int _currentSlideIndex;

        static uint uniqueId;

        private bool _isPlaying = false;

        private PowerPoint.Application _application;
        private PowerPoint.Presentation _presentation;

        private int _intervalCounter = 0;
        private String _slideNotes = String.Empty;
        private String _slideNotesNext = String.Empty;

        public String slideNotes {  get { return this._slideNotes; } }
        public String slideNotesNext { get { return this._slideNotesNext; } }

        public bool IsPlaying { get { return _isPlaying; } }

        private string _pptPath = System.IO.Path.GetTempPath() + "slidecat\\";

        private bool _stopping = false;
        public bool stopping { get { return _stopping; } }

        public SlideCatPresentation()
        {
            if(!Directory.Exists(this._pptPath))
            {
                Directory.CreateDirectory(this._pptPath);
            }
            this._pptPath += new Random().Next() + "/";

            if(!Directory.Exists(_pptPath))
            {
                Directory.CreateDirectory(this._pptPath);
            }
        }

        private void _emptyPresentationDirectory()
        {
            if(!this._isPlaying)
            {

                System.IO.DirectoryInfo di = new DirectoryInfo(this._pptPath);

                foreach (FileInfo file in di.GetFiles())
                {
                    try
                    {
                        file.Delete();
                    }
                    catch (Exception ex)
                    { 
                        Console.WriteLine(ex.Message);
                    } 
                }
                foreach (DirectoryInfo dir in di.GetDirectories())
                {
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
            
        }

        public void createPresentation(MediaItems mediaItems, ref BackgroundWorker worker)
        {
            //  the destination powerpoint
            String _destinationPowerPoint = "destinationPowerPoint.pptx";

            //  clear the folder to which temporary files are stored
            this._emptyPresentationDirectory();

            //  initiate new application and main presentation

            PowerPoint.Application powerPointApplication = new PowerPoint.Application();
            PowerPoint.Presentation powerPointPresentation = powerPointApplication.Presentations.Add(MsoTriState.msoFalse);

            //  insert first slide, make black
            PowerPoint.CustomLayout customLayout = powerPointPresentation.SlideMaster.CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutTitle];
            PowerPoint.Slide slide = powerPointPresentation.Slides.AddSlide(1, customLayout);
            Color slideBGColor = Color.Black;
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.ForeColor.RGB = slideBGColor.ToArgb();
            powerPointPresentation.SaveCopyAs(this._pptPath + "destinationPowerPoint", PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
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
                double percentageDouble = (i * 100) / (nrItems + 1);
                percentageDouble = Math.Round(percentageDouble);
                int percentageInt = (int)percentageDouble;
                worker.ReportProgress(percentageInt);

                //  store each powerpoint to a file
                PowerPoint.Presentation pres = mediaItem.presentation;
                pres.SaveCopyAs(this._pptPath + "temporary_pptx_" + i, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                pres.Close();

                //  merge into the destination powerpoint
                MergeSlides(this._pptPath, "temporary_pptx_" + i + ".pptx", _destinationPowerPoint);

                //  delete temporary powerpoint
                File.Delete(this._pptPath + "temporary_pptx_" + i + ".pptx");
            }

            this._application = new PowerPoint.Application();
            this._presentation = _application.Presentations.Open2007(this._pptPath + _destinationPowerPoint, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue);
        }

        public void playPresentation()
        {
            this._isPlaying = false;

            //  must have at least one slide to continue
            if (!(this._presentation.Slides.Count > 0))
            {
                return;
            }

            //  load the powerpoint in preview mode and run
            PowerPoint.SlideShowSettings settings = this._presentation.SlideShowSettings;
            settings.ShowType = (PowerPoint.PpSlideShowType)1;
            settings.ShowPresenterView = MsoTriState.msoTrue;
            PowerPoint.SlideShowWindow sw = settings.Run();

            //  go to first slide
            this._presentation.SlideShowWindow.View.GotoSlide(1);
            this._presentation.SlideShowWindow.View.FirstAnimationIsAutomatic();
            this._stopping = false;

            //  add handler for when stopping
            this._application.PresentationBeforeClose += delegate
            {
                this._stopping = true;
                this._isPlaying = false;
            };

            //  set status to playing
            this._isPlaying = true;
        }

        public void stopPresentation()
        {
            if (this._presentationPlaying())
            {
                this._isPlaying = false;
                this._stopPresentation(this._presentation);
            }
        }

        private void _stopPresentation(PowerPoint.Presentation _presentation)
        {
            if (this._presentationPlaying())
            {
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
        }
        
       
        
       public void stop()
        {
            if(this._presentationPlaying())
            {
                this.stopPresentation();
                this._emptyPresentationDirectory();
            }
        }      

        private bool _presentationPlaying()
        {
            try
            {
                int tmp = this._presentation.Slides.Count;
                return this._isPlaying;
                
            }catch
            {
                return false;
            }
        }



        static void MergeSlides(string presentationFolder, string sourcePresentation, string destPresentation)
        {
            int id = 0;
            Console.WriteLine(presentationFolder);
            Console.WriteLine(sourcePresentation);
            Console.WriteLine(destPresentation);
            // Open the destination presentation.
            using (PresentationDocument myDestDeck = PresentationDocument.Open(presentationFolder + destPresentation, true))
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

        static void DisplayValidationErrors(
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
