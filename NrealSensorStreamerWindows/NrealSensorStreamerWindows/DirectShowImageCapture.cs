namespace NrealSensorStreamerWindows
{
    using System;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Threading;
    using DirectShowLib;

    public class DirectShowImageCapture : ISampleGrabberCB
    {        
        private Guid mediaType;
        private int width;
        private int height;
        private int minimumFrameRate;

        private DsDevice cameraDevice;
        private IFilterGraph2 filterGraph = null;
        public IMediaControl mediaCtrl = null;
        public IBaseFilter captureFilter = null;
     
        private Mutex mutex = new Mutex();

        private DirectShowImageCapture(DsDevice device, Guid mediaType, int width, int height, int minimumFrameRate = 30)
        {
            this.cameraDevice = device;
            this.mediaType = mediaType;
            this.width = width;
            this.height = height;
            this.minimumFrameRate = minimumFrameRate;
            this.filterGraph = new FilterGraph() as IFilterGraph2;
            this.SetupGraph();
        }

        public event EventHandler<byte[]> FrameReady;

        public static bool TryCreateImageCapture(string deviceId, Guid mediaType, int width, int height, int minimumFrameRate, out DirectShowImageCapture imageCapture)
        {
            // Get the collection of video devices
            var captureDevices = DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice);

            var cameraDevice = captureDevices.Where(x => x.DevicePath.Contains(deviceId)).FirstOrDefault();
            if (cameraDevice == null)
            {
                imageCapture = null;
                return false;
            }

            imageCapture = new DirectShowImageCapture(cameraDevice, mediaType, width, height, minimumFrameRate);
            return true;
        }

        public void Start()
        {
            // Start the graph
            this.mediaCtrl = this.filterGraph as IMediaControl;
            this.mediaCtrl.Run();
        }

        public void Stop()
        {
            if (this.filterGraph != null)
            {
                IMediaControl mediaCtrl = this.filterGraph as IMediaControl;

                // Stop the graph
                mediaCtrl.Stop();
            }
        }

        private void SetupGraph()
        {
            int hr;
            try
            {
                // add the video input device
                hr = this.filterGraph.AddSourceFilterForMoniker(this.cameraDevice.Mon, null, this.cameraDevice.Name, out this.captureFilter);

                DsError.ThrowExceptionForHR(hr);
                IPin pCaptureOut = DsFindPin.ByDirection(this.captureFilter, PinDirection.Output, 0);

                SetConfigParms(pCaptureOut);

                ISampleGrabber sampleGrabber = new SampleGrabber() as ISampleGrabber;
                ConfigureSampleGrabber(sampleGrabber);

                // Get the default video renderer
                IBaseFilter pRenderer = new NullRenderer() as IBaseFilter;
                hr = this.filterGraph.AddFilter(pRenderer, string.Format("NullRenderer {0}", 0));
                DsError.ThrowExceptionForHR(hr);

                SmartTee smartTee = null;

                // Add the sample grabber to the graph
                hr = this.filterGraph.AddFilter((IBaseFilter)sampleGrabber, string.Format("Ds.NET Grabber {0}", 0));
                DsError.ThrowExceptionForHR(hr);

                // Connect the Capture pin to the sample grabber                    
                hr = this.filterGraph.Connect(pCaptureOut, DsFindPin.ByDirection((IBaseFilter)sampleGrabber, PinDirection.Input, 0));
                hr = this.filterGraph.Connect(DsFindPin.ByDirection((IBaseFilter)sampleGrabber, PinDirection.Output, 0), DsFindPin.ByDirection(pRenderer, PinDirection.Input, 0));

                if (smartTee != null)
                {
                    Marshal.ReleaseComObject(smartTee);
                }

                if (pRenderer != null)
                {
                    Marshal.ReleaseComObject(pRenderer);
                }

                if (sampleGrabber != null)
                {
                    Marshal.ReleaseComObject(sampleGrabber);
                }

                DsError.ThrowExceptionForHR(hr);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
            }
        }

        private void SetConfigParms(IPin pStill)
        {
            int hr;

            IAMStreamConfig videoStreamConfig = pStill as IAMStreamConfig;
            if (null == videoStreamConfig)
                throw new Exception("Cannot obtain IAMStreamConfig");

            // Get number of possible combinations of format and size
            int capsCount, capSize;
            hr = videoStreamConfig.GetNumberOfCapabilities(out capsCount, out capSize);
            if (hr != 0)
            {
                DsError.ThrowExceptionForHR(hr);
            }

            VideoInfoHeader vih = new VideoInfoHeader();
            VideoStreamConfigCaps vsc = new VideoStreamConfigCaps();
            IntPtr pSC = Marshal.AllocHGlobal(capSize);
            AMMediaType mt = null;
            int videoFormatIndex = -1;

            for (int i = 0; i < capsCount; ++i)
            {
                // the video format is described in AMMediaType and VideoStreamConfigCaps
                hr = videoStreamConfig.GetStreamCaps(i, out mt, pSC);
                DsError.ThrowExceptionForHR(hr);

                // copy the unmanaged structures to managed in order to check the format
                Marshal.PtrToStructure(mt.formatPtr, vih);
                Marshal.PtrToStructure(pSC, vsc);

                // check colorspace
                if (mt.subType == this.mediaType)
                {
                    // the video format has a range of supported frame rates (min-max)
                    // check the required frame rate and frame size
                    if ((vih.BmiHeader.Width == this.width) &&
                        (vih.BmiHeader.Height == this.height) &&
                        10000000 / vsc.MaxFrameInterval >= this.minimumFrameRate)
                    {
                        // remember the index of the video format that we’ll use
                        videoFormatIndex = i;
                        break;
                    }
                }

                DsUtils.FreeAMMediaType(mt);
            }

            // didn't find what we wanted
            if (videoFormatIndex < 0)
            {
                throw (new Exception(string.Format("Unable to find acceptable format for {0}, {1} x {2}", this.mediaType, this.width, this.height)));
            }

            try
            {                
                hr = videoStreamConfig.GetStreamCaps(videoFormatIndex, out mt, pSC);
                DsError.ThrowExceptionForHR(hr);
                // explicitly set the framerate since the default may not what we want
                Marshal.PtrToStructure(mt.formatPtr, vih);

                Marshal.StructureToPtr(vih, mt.formatPtr, false);
                hr = videoStreamConfig.SetFormat(mt);
                DsError.ThrowExceptionForHR(hr);
            }
            catch (Exception e)
            {
                throw (e);
            }
            finally
            {
                DsUtils.FreeAMMediaType(mt);
                Marshal.FreeHGlobal(pSC);
                mt = null;
            }
        }

        private void ConfigureSampleGrabber(ISampleGrabber sampGrabber)
        {
            int hr;
            AMMediaType media = new AMMediaType();

            // Set the media type to Video/NV12
            media.majorType = MediaType.Video;
            media.subType = this.mediaType;
            media.formatType = FormatType.VideoInfo;
            hr = sampGrabber.SetMediaType(media);
            DsError.ThrowExceptionForHR(hr);

            DsUtils.FreeAMMediaType(media);
            media = null;

            hr = sampGrabber.SetBufferSamples(true);
            hr = sampGrabber.SetOneShot(false);

            hr = sampGrabber.SetCallback(this, 1);
            DsError.ThrowExceptionForHR(hr);
        }

        int ISampleGrabberCB.SampleCB(double SampleTime, IMediaSample pSample)
        {
            Marshal.ReleaseComObject(pSample);
            return 0;
        }

        int ISampleGrabberCB.BufferCB(double SampleTime, IntPtr pBuffer, int BufferLen)
        {
            mutex.WaitOne();
            var data = new byte[BufferLen];
            Marshal.Copy(pBuffer, data, 0, BufferLen);
            this.FrameReady?.Invoke(this, data);
            mutex.ReleaseMutex();
            return 0;
        }
    }
}
