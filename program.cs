using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ConvertSldToGif
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Count() < 2)
            {
                Console.WriteLine("Please specify the file name to convert and the target GIF file name.");
                //Console.ReadKey();
                return;
            }
            if (!File.Exists(args[0]))
            {
                Console.WriteLine(string.Format("[{0}] doesn't exist", args[0]));
                //Console.ReadKey();
                return;
            }
            swDocumentTypes_e sldworksFileType = swDocumentTypes_e.swDocNONE;
            string lowerExt = new FileInfo(args[0]).Extension.ToLower();
            if (string.Compare(lowerExt, ".sldprt") == 0)
            {
                sldworksFileType = swDocumentTypes_e.swDocPART;
            }
            else if (string.Compare(lowerExt, ".sldasm") == 0)
            {
                sldworksFileType = swDocumentTypes_e.swDocASSEMBLY;
            }
            else
            {
                Console.WriteLine("Currently only .sldprt and .sldasm extension types are supported");
                //Console.ReadKey();
                return;
            }
            Console.WriteLine("Now talking to SolidWorks to generate the PNG from {0}", args[0]);

            FileInfo destinationGifFileInfo = new FileInfo(args[1]);
            if (destinationGifFileInfo.Extension.ToLower() != ".gif")
            {
                Console.WriteLine("Sorry, the second file name must be of .gif");
                //Console.ReadKey();
                return;
            }
            if (!Directory.Exists(destinationGifFileInfo.DirectoryName))
            { // If the destination folder doesn't exist, create it first
                Directory.CreateDirectory(destinationGifFileInfo.DirectoryName);
            }
            object processSW = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
            if (processSW == null)
            {
                Console.WriteLine("Couldn't create an instance of SolidWorks.");
                //Console.ReadKey();
                return;
            }
            SldWorks sldWorks = (SldWorks)processSW;
            int error = 0;
            int warning = 0;
            ModelDoc2 modelDoc = sldWorks.OpenDoc6(
                args[0],
                (int)sldworksFileType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                string.Empty,
                ref error,
                ref warning);
            if (modelDoc == null)
            {
                sldWorks.ExitApp();
                sldWorks = null;
                Console.WriteLine("Failed to open the document, press any key to quit");
                return;
            }
            if (!GenerateGif(modelDoc, destinationGifFileInfo.DirectoryName, destinationGifFileInfo.FullName))
            {
                Console.WriteLine("Failed to create GIF from the document [{0}]", args[0]);
            }
            else
            {
                Console.WriteLine("Successfully created [{0}]", destinationGifFileInfo.FullName);
            }
           
            modelDoc = null;
            sldWorks.ExitApp();
            sldWorks = null;
            //Console.ReadKey();
        }

        /// <summary>
        /// Generate some PNGs from a SolidWorks document
        /// </summary>
        /// <param name="modelDoc">The model doc of specified part</param>
        /// <param name="directory">The name of the directory to save PNG files</param>
        /// <param name="destinationGifFile">The destination GIF file name</param>
        /// <returns>true - succeeded, false - failed</returns>
        private static bool GenerateGif(ModelDoc2 modelDoc, string directory, string destinationGifFile)
        {
            if (modelDoc == null)
            {
                return false;
            }
            GifBitmapEncoder pngBitmapEncoder = new GifBitmapEncoder();
            modelDoc.ViewZoomtofit2();
            Guid guid = Guid.NewGuid();
            string path = Path.Combine(directory, string.Format("{0}_1.PNG", guid));
            if (modelDoc.SaveAsSilent(path, true) != 0)
            {
                return false;
            }
            Dictionary<string, FileStream> allBitmaps = new Dictionary<string, FileStream>();
            FileStream bitmapStream1 = new FileStream(path, FileMode.Open); // don't close before the encoder.Save
            pngBitmapEncoder.Frames.Add(BitmapFrame.Create(bitmapStream1));
            allBitmaps[path] = bitmapStream1;

            modelDoc.ViewRotate();
            modelDoc.ViewRotateminusx();
            path = Path.Combine(directory, string.Format("{0}_2.PNG", guid));
            if (modelDoc.SaveAsSilent(path, true) != 0)
            {
                return false;
            }
            FileStream bitmapStream2 = new FileStream(path, FileMode.Open);
            pngBitmapEncoder.Frames.Add(BitmapFrame.Create(bitmapStream2));
            allBitmaps[path] = bitmapStream2;

            modelDoc.ViewRotateminusy();
            path = Path.Combine(directory, string.Format("{0}_3.PNG", guid));
            if (modelDoc.SaveAsSilent(path, true) != 0)
            {
                return false;
            }
            FileStream bitmapStream3 = new FileStream(path, FileMode.Open);
            pngBitmapEncoder.Frames.Add(BitmapFrame.Create(bitmapStream3));
            allBitmaps[path] = bitmapStream3;

            modelDoc.ViewRotateminusz();
            path = Path.Combine(directory, string.Format("{0}_4.PNG", guid));
            if (modelDoc.SaveAsSilent(path, true) != 0)
            {
                return false;
            }
            FileStream bitmapStream4 = new FileStream(path, FileMode.Open);
            pngBitmapEncoder.Frames.Add(BitmapFrame.Create(bitmapStream4));
            allBitmaps[path] = bitmapStream4;

            
            modelDoc.ViewRotXMinusNinety();
            path = Path.Combine(directory, string.Format("{0}_8.PNG", guid));
            if (modelDoc.SaveAsSilent(path, true) != 0)
            {
                return false;
            }
            FileStream bitmapStream8 = new FileStream(path, FileMode.Open);
            pngBitmapEncoder.Frames.Add(BitmapFrame.Create(bitmapStream8));
            allBitmaps[path] = bitmapStream8;

            modelDoc.ViewRotYMinusNinety();
            path = Path.Combine(directory, string.Format("{0}_9.PNG", guid));
            if (modelDoc.SaveAsSilent(path, true) != 0)
            {
                return false;
            }
            FileStream bitmapStream9 = new FileStream(path, FileMode.Open);
            pngBitmapEncoder.Frames.Add(BitmapFrame.Create(bitmapStream9));
            allBitmaps[path] = bitmapStream9;

            modelDoc.ViewRotXPlusNinety();
            path = Path.Combine(directory, string.Format("{0}_10.PNG", guid));
            if (modelDoc.SaveAsSilent(path, true) != 0)
            {
                return false;
            }
            FileStream bitmapStream10 = new FileStream(path, FileMode.Open);
            pngBitmapEncoder.Frames.Add(BitmapFrame.Create(bitmapStream10));
            allBitmaps[path] = bitmapStream10;
            
            modelDoc.ViewRotYPlusNinety();
            path = Path.Combine(directory, string.Format("{0}_11.PNG", guid));
            if (modelDoc.SaveAsSilent(path, true) != 0)
            {
                return false;
            }
            FileStream bitmapStream11 = new FileStream(path, FileMode.Open);
            pngBitmapEncoder.Frames.Add(BitmapFrame.Create(bitmapStream11));
            allBitmaps[path] = bitmapStream11;

            modelDoc.ViewRotateplusx();
            path = Path.Combine(directory, string.Format("{0}_5.PNG", guid));
            if (modelDoc.SaveAsSilent(path, true) != 0)
            {
                return false;
            }
            FileStream bitmapStream5 = new FileStream(path, FileMode.Open);
            pngBitmapEncoder.Frames.Add(BitmapFrame.Create(bitmapStream5));
            allBitmaps[path] = bitmapStream5;

            modelDoc.ViewRotateplusy();
            path = Path.Combine(directory, string.Format("{0}_6.PNG", guid));
            if (modelDoc.SaveAsSilent(path, true) != 0)
            {
                return false;
            }
            FileStream bitmapStream6 = new FileStream(path, FileMode.Open);
            pngBitmapEncoder.Frames.Add(BitmapFrame.Create(bitmapStream6));
            allBitmaps[path] = bitmapStream6;

            modelDoc.ViewRotateplusz();
            path = Path.Combine(directory, string.Format("{0}_7.PNG", guid));
            if (modelDoc.SaveAsSilent(path, true) != 0)
            {
                return false;
            }
            FileStream bitmapStream7 = new FileStream(path, FileMode.Open);
            pngBitmapEncoder.Frames.Add(BitmapFrame.Create(bitmapStream7));
            allBitmaps[path] = bitmapStream7;

            using (FileStream stream = new FileStream(destinationGifFile, FileMode.Create))
            {
                pngBitmapEncoder.Save(stream);
            }

            foreach (string pngFileToDelete in allBitmaps.Keys)
            {
                allBitmaps[pngFileToDelete].Close(); // we can only close after encoder.Save
                File.Delete(pngFileToDelete);
            }

            return true;
        }

    }
}
