using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
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
        private static uint _mUniqueId;

        private readonly string _mPptPath = string.Empty;

        private readonly string _mSlideCatPath = Path.GetTempPath() + "slidecat\\";

        private Application _mApplication;
        private Presentation _mPresentation;

        public SlideCatPresentation()
        {
            Console.WriteLine(_mSlideCatPath);
            if (!Directory.Exists(_mSlideCatPath)) Directory.CreateDirectory(_mSlideCatPath);
            _mPptPath = _mSlideCatPath + new Random().Next() + "\\";

            if (!Directory.Exists(_mPptPath)) Directory.CreateDirectory(_mPptPath);
        }

        public bool mIsPlaying { get; private set; }

        public bool mStopping { get; private set; }

        private void _EmptyPresentationDirectory(bool deleteFolder = false)
        {
            
            if (mIsPlaying) return;
            DirectoryInfo directoryInfo = new DirectoryInfo(_mPptPath);

            foreach (FileInfo file in directoryInfo.GetFiles())
            {
                
                try
                {
                    file.Delete();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception deleting file " + file.Name + " - " + ex.Message);
                }
            }

            foreach (DirectoryInfo dir in directoryInfo.GetDirectories())
            {
                Console.WriteLine(dir.FullName);
                try
                {
                    dir.Delete(true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception deleting directory " + dir.FullName + " - " + ex.Message);
                }
            }

            if (deleteFolder)
            {
                try { Directory.Delete(_mPptPath, true); }
                catch (Exception ex) { Console.WriteLine(ex.Message); }
            }
        }

        public void CreatePresentation(ItemManager mediaItems, ref BackgroundWorker worker)
        {
            //  the destination powerpoint
            string destinationPowerPointName = "destinationPowerPoint";
            string destinationPowerPointExtension = "pptx";
            string temporaryPowerPointName = "temporary_pptx_";
            string temporaryPowerPointExtension = "pptx";

            //  clear the folder to which temporary files are stored
            _EmptyPresentationDirectory();

            //  initiate new application and main presentation

            Application powerPointApplication = new Application();
            Presentation powerPointPresentation = powerPointApplication.Presentations.Add(MsoTriState.msoFalse);

            //  insert first slide, make black
            CustomLayout customLayout = powerPointPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle];
            Slide slide = powerPointPresentation.Slides.AddSlide(1, customLayout);
            Color slideBgColor = Color.Black;
            slide.FollowMasterBackground = MsoTriState.msoFalse;
            slide.Background.Fill.ForeColor.RGB = slideBgColor.ToArgb();
            powerPointPresentation.SaveCopyAs(_mPptPath + destinationPowerPointName, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            powerPointPresentation.Close();

            //  save each presentation as powerpoint presentation into the tmp folder
            //  add each temporary powerpoint into the main powerpoint
            mediaItems.Sort();
            int i = 0;
            int nrItems = mediaItems.mMediaItems.Count;
            foreach (MediaItem mediaItem in mediaItems.mMediaItems)
            {
                //  indicate the nth powerpoint to be converting
                i++;

                //  report progress to the background-worker
                double percentageDouble = i * (100 / (nrItems + 1));
                percentageDouble = Math.Round(percentageDouble);
                int percentageInt = (int)percentageDouble;
                worker.ReportProgress(percentageInt);

                //  store each powerpoint to a file
                Presentation pres = mediaItem.presentation;
                pres.SaveCopyAs(_mPptPath + temporaryPowerPointName + i, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                pres.Close();

                //  merge into the destination powerpoint
                MergeSlides(_mPptPath, temporaryPowerPointName + i + "." + temporaryPowerPointExtension, destinationPowerPointName+"."+destinationPowerPointExtension);

                //  delete temporary powerpoint
                File.Delete(_mPptPath + temporaryPowerPointName + i + "." + temporaryPowerPointExtension);
            }

            _mApplication = new Application();
            _mPresentation = _mApplication.Presentations.Open2007(
                _mPptPath + destinationPowerPointName + "." + destinationPowerPointExtension, 
                MsoTriState.msoTrue,
                MsoTriState.msoFalse, 
                MsoTriState.msoFalse, 
                MsoTriState.msoTrue
                );
        }

        public void PlayPresentation()
        {
            mIsPlaying = false;

            //  must have at least one slide to continue
            if (!(_mPresentation.Slides.Count > 0)) return;

            //  load the powerpoint in preview mode and run
            SlideShowSettings settings = _mPresentation.SlideShowSettings;
            settings.ShowType = (PpSlideShowType)1;
            settings.ShowPresenterView = MsoTriState.msoTrue;
            SlideShowWindow sw = settings.Run();

            //  go to first slide
            _mPresentation.SlideShowWindow.View.GotoSlide(1);
            _mPresentation.SlideShowWindow.View.FirstAnimationIsAutomatic();
            
            mStopping = false;

            //  add handler for when stopping
            _mApplication.PresentationBeforeClose += delegate
            {
                mStopping = true;
                mIsPlaying = false;

                // force close without save
                // https://social.msdn.microsoft.com/Forums/vstudio/en-US/1390d2ff-aa94-490d-a689-569a573bd0b4/how-to-close-powerpoint-in-wpf-without-save-or-dont-save-alerts-?forum=wpf
                Process[] processes = Process.GetProcesses();
                for (int i = 0; i < processes.Count(); i++)
                {
                    if (processes[i].ProcessName.ToLower().Contains("powerpnt"))
                    {
                        processes[i].Kill();
                    }
                }
                
            };

            //  set status to playing
            mIsPlaying = true;
        }

        public void StopPresentation()
        {
            if (!_PresentationPlaying()) return;
            mIsPlaying = false;
            _StopPresentation();
        }

        private void _StopPresentation()
        {
            if (!_PresentationPlaying()) return;
            try
            {
                _mPresentation.Save();
                _mPresentation.Close();
                _mPresentation = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("LOG - Presentation.cs - _stopPresentation() - catch");
                Console.WriteLine(ex.Message);
            }
        }


        public void Stop()
        {
            if (_PresentationPlaying())
            {
                StopPresentation();
            }

            _EmptyPresentationDirectory(true);
            
        }

        private bool _PresentationPlaying()
        {
            try
            {
                if (_mPresentation == null) return false;
                int tmp = _mPresentation.Slides.Count;
                return mIsPlaying;
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
            using (PresentationDocument myDestDoc = PresentationDocument.Open(presentationFolder + destPresentation, true))
            {
                PresentationPart destPresPart = myDestDoc.PresentationPart;
                if (destPresPart == null) return;

                // If the merged presentation does not have a SlideIdList 
                // element yet, add it.
                if (destPresPart.Presentation.SlideIdList == null)
                {
                    destPresPart.Presentation.SlideIdList = new SlideIdList();
                }
                    
                // Open the source presentation. This will throw an exception if
                // the source presentation does not exist.
                using (PresentationDocument mySourceDoc = PresentationDocument.Open(presentationFolder + sourcePresentation, false))
                {
                    PresentationPart sourcePresPart = mySourceDoc.PresentationPart;
                    if (sourcePresPart == null) return;
                    if (sourcePresPart.Presentation.SlideIdList == null) return;


                    // Get unique ids for the slide master and slide lists
                    // for use later.
                    _mUniqueId = GetMaxSlideMasterId(destPresPart.Presentation.SlideMasterIdList);

                    uint maxSlideId = GetMaxSlideId(destPresPart.Presentation.SlideIdList);

                    // Copy each slide in the source presentation, in order, to 
                    // the destination presentation.
                    foreach (SlideId slideId in sourcePresPart.Presentation.SlideIdList)
                    {
                        // Create a unique relationship id.
                        id++;
                        if (slideId.RelationshipId == null) continue;
                        SlidePart slidePart = (SlidePart)sourcePresPart.GetPartById(slideId.RelationshipId);

                        string relId = sourcePresentation.Remove(sourcePresentation.IndexOf('.')) + id;

                        // Add the slide part to the destination presentation.
                        SlidePart destSlidePart = destPresPart.AddPart(slidePart, relId);

                        // The slide master part was added. Make sure the
                        // relationship between the main presentation part and
                        // the slide master part is in place.
                        if (destSlidePart?.SlideLayoutPart == null) continue;
                        SlideMasterPart destMasterPart = destSlidePart.SlideLayoutPart.SlideMasterPart;

                        if (destMasterPart == null) continue;
                        destPresPart.AddPart(destMasterPart);

                        // Add the slide master id to the slide master id list.
                        _mUniqueId++;
                        SlideMasterId newSlideMasterId = new SlideMasterId()
                        {
                            RelationshipId = destPresPart.GetIdOfPart(destMasterPart),
                            Id = _mUniqueId
                        };

                        if (destPresPart.Presentation.SlideMasterIdList == null) continue;
                        destPresPart.Presentation.SlideMasterIdList.Append(newSlideMasterId);

                        // Add the slide id to the slide id list.
                        maxSlideId++;
                        SlideId newSlideId = new SlideId()
                        {
                            RelationshipId = relId,
                            Id = maxSlideId
                        };

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
            foreach (SlideMasterPart slideMasterPart in presPart.SlideMasterParts)
            {
                if (slideMasterPart.SlideMaster.SlideLayoutIdList == null) continue;
                foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList)
                {
                    _mUniqueId++;
                    slideLayoutId.Id = _mUniqueId;
                }
                slideMasterPart.SlideMaster.Save();
            }
        }

        private static uint GetMaxSlideId(SlideIdList slideIdList)
        {
            // Slide identifiers have a minimum value of greater than or
            // equal to 256 and a maximum value of less than 2147483648. 
            uint max = 256;

            if (slideIdList == null) return max;
            
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

            if (slideMasterIdList == null) return max;
            // Get the maximum id value from the current set of children.
            foreach (SlideMasterId child in slideMasterIdList.Elements<SlideMasterId>())
            {
                uint id = child.Id;

                if (id > max) max = id;
            }

            return max;
        }

        private static void DisplayValidationErrors(IEnumerable<ValidationErrorInfo> errors)
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