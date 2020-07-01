using Pulse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SetColors
{
    class Program
    {
        static IThreadProperty FindThreadByCode (IThreadTables ThreadDatabase, string manufacturer, string code )
        {
            IThreadTable SelectedThradChart = ThreadDatabase[manufacturer];
            int i = 0;
            bool found = false;
            IThreadProperty aThreead = SelectedThradChart[i];
            while (i < SelectedThradChart.Count & !found)
            {
                
                if (aThreead.Code == code)
                {
                    found = true;
                }
                else
                {
                    i++;
                    aThreead = SelectedThradChart[i];

                }    
              
            }
            return aThreead;
        }

        static void PrintThreadChartNames(IThreadTables ThreadDatabase)

        {
            Console.WriteLine("Thread Chart Names");
            Console.WriteLine("------------------------------");


            for (int i = 0; i < ThreadDatabase.Count; i++)
            {
                IThreadTable SelectedThradChart = ThreadDatabase[i];
                Console.WriteLine(SelectedThradChart.Name);
             

            }

            Console.WriteLine();
        }
        

            static void UpdateThreadColor(IThreadTables ThreadDatabase, IEmbDesign myDesign, int threadIndex, string manufacturer, string code)
        {
            INeedleSequence designNeedleSequence = myDesign.NeedleSequence;
            IThreadPalette designThreadPalette = myDesign.ThreadPalette;
            int needle = designNeedleSequence[threadIndex];
            IThreadProperty newThread = FindThreadByCode(ThreadDatabase,manufacturer,code);
            designThreadPalette[needle].Code = newThread.Code;
            designThreadPalette[needle].Manufacturer = newThread.Manufacturer;
            designThreadPalette[needle].Name = newThread.Name;
            designThreadPalette[needle].Red = newThread.Red;
            designThreadPalette[needle].Green = newThread.Green;
            designThreadPalette[needle].Blue = newThread.Blue;
        }

         

        static void Main(string[] args)
        {
            IApplication PulseID = new Pulse.Application();
            try
            {
                IEmbDesign myDesign = PulseID.OpenDesign("C:\\Temp\\Eagle2.pxf", FileTypes.ftAuto, OpenTypes.otDefault, "Tajima");
                try
                {


                    IBitmapImage myImage = PulseID.NewImage(300, 300);
                    try
                    {
                            myDesign.Render(myImage, 0, 0, 300, 300);
                            myImage.Save("C:\\Temp\\myImage.png", ImageTypes.itAuto);

                            PrintThreadChartNames(PulseID.ThreadCharts);

                        INeedleSequence designNeedleSequence = myDesign.NeedleSequence;
                        try
                        {
                            IThreadPalette designThreadPalette = myDesign.ThreadPalette;
                            // Get the thread palette information for the design.
                            try
                            {
                                // find the needle that is used for each color in the design.  Needle sequence is zero based 0=needle 1
                                for (int i = 0; i < designNeedleSequence.Count; i++)
                                {
                                    int needle = designNeedleSequence[i];
                                    IThreadProperty threadInfo = designThreadPalette[needle];

                                    Console.WriteLine("Manufacturer= {0} Name= {1} Code= {2} R= {3} G= {4} B= {5}", threadInfo.Manufacturer, threadInfo.Name, threadInfo.Code, threadInfo.Red, threadInfo.Green, threadInfo.Blue);

                                    
                                }
                                UpdateThreadColor(PulseID.ThreadCharts, myDesign, 0, "Madeira Classic Rayon 40", "1002");
                                UpdateThreadColor(PulseID.ThreadCharts, myDesign, 1, "Madeira Classic Rayon 40", "1000");


                                Console.WriteLine();
                                Console.WriteLine("New Thread Sequence");
                                Console.WriteLine("------------------------------" +
                                    "");

                                // Get the thread palette information for the design.

                                // find the needle that is used for each color in the design.  Needle sequence is zero based 0=needle 1
                                for (int i = 0; i < designNeedleSequence.Count; i++)
                                        {
                                            int needle = designNeedleSequence[i];
                                            IThreadProperty threadInfo = designThreadPalette[needle];

                                            Console.WriteLine("Manufacturer= {0} Name= {1} Code= {2} R= {3} G= {4} B= {5}", threadInfo.Manufacturer, threadInfo.Name, threadInfo.Code, threadInfo.Red, threadInfo.Green, threadInfo.Blue);


                                        }

                                        Console.WriteLine("Done");

                                ;
                           
                                Console.ReadLine();

                                myDesign.Render(myImage, 0, 0, 300, 300);
                                myImage.Save("C:\\Temp\\myNewImage.png", ImageTypes.itAuto);
                                myDesign.Save("C:\\Temp\\NewDesign.pxf", FileTypes.ftAuto);
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(designThreadPalette);
                            }
                            
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(designNeedleSequence);
                        }


                       
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(myImage);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(myDesign);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(PulseID);
            }
        }
    }
}
