using Microsoft.Extensions.DependencyInjection;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.IO;

namespace PowerPointToPlaces
{
    class Program
    {
        private static readonly Application ppt = new Application();
        private static ServiceProvider serviceProvider;
        static void Main(string[] args)
        {
            Console.WriteLine("Connecting to PowerPoint...");

            ppt.SlideShowNextSlide += ShowNextSlide;

            var sc = new ServiceCollection();
            sc.AddSingleton<ISendToPlace, SlackPlace>();

            serviceProvider = sc.BuildServiceProvider();

            Console.WriteLine("Connected...");

            Console.ReadLine();
        }

        private static void ShowNextSlide(SlideShowWindow Wn)
        {
            if (Wn != null)
            {
                Console.WriteLine($"Moved to Slide Number {Wn.View.Slide.SlideNumber}");

                var note = string.Empty;
                try { note = Wn.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text; }
                catch { /*no notes*/ }

                using var reader = new StringReader(note);

                string line;
                var places = serviceProvider.GetServices<ISendToPlace>();

                while ((line = reader.ReadLine()) != null)
                {
                    foreach (var place in places)
                    {
                        if (place.CanSend(line))
                            place.SendAsync(line).Wait();
                    }
                }
            }
        }
    }
}
