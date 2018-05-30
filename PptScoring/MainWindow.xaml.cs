using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace PptScoring
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        PowerPoint.Application ppt_app;
        PowerPoint.Presentation objPres;
        string[] pptfiles;// = { @"C:/Users/Alien1/Downloads/1.pptx", @"C:/Users/Alien1/Downloads/2.pptx", @"C:/Users/Alien1/Downloads/4.pptx" };
        string path = @"C:\Users\Alien1\Documents\diapositivas_estudiantes\1";
        int curr = -1;
        int currentSlide = 1;
        int []size_scores;
        int []cantidad_scores;
        StreamWriter writeText;
        public MainWindow()
        {
            InitializeComponent();

            pptfiles = Directory.GetFiles(path, "*.pptx");
      
            foreach (var item in pptfiles)
            {
                Console.WriteLine(item);
            }
            this.writeText = new StreamWriter(@"C:/Users/Alien1/Documents/scores.csv", true);
            writeText.WriteLine("ppt_file, slide_id, size, quantity");
            this.ppt_app = new PowerPoint.Application();
            ppt_app.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            ppt_app.SlideShowBegin += Ppt_app_SlideShowBegin;
            ppt_app.SlideShowEnd += Ppt_app_SlideShowEnd;
            ppt_app.PresentationNewSlide += Ppt_app_PresentationNewSlide;

        }

        private void Ppt_app_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            Console.WriteLine("New slide");
        }

        public bool Open_PPT(string ppt_filename)
        {
            try
            {
                var presentation = ppt_app.Presentations;
                
                this.objPres = presentation.Open(ppt_filename, Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);
                currentSlide = 1;
                
                //var objSSS = objPres.SlideShowSettings;
                //objSSS.Run();
                //objPres.SlideShowWindow.Activate();

                //presentation = null;
                //objSSS = null;
                return true;
            }catch (Exception e)
            {
                Console.WriteLine("Error startting ppt");
                Console.WriteLine(e);
                return false;
            }
        }

        private void Ppt_app_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            Console.WriteLine("SlideShow Began");
        }
        private void Ppt_app_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            Console.WriteLine("END");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.OpenNewPresentation();
        }
        private void OpenNewPresentation()
        {
            if (this.size_scores != null && this.cantidad_scores != null)
                this.WriteToCsv();
            if (++curr == pptfiles.Length)
                curr = 0;
            if (this.objPres != null)
                this.objPres.Close();
            this.label_pptfile.Content = pptfiles[curr];
            
            Open_PPT(pptfiles[curr]);
            this.size_scores = new int[this.objPres.Slides.Count];
            this.cantidad_scores = new int[this.objPres.Slides.Count];
            this.showWindow();
        }
        private void NextSlide()
        {
            if (this.objPres != null && this.currentSlide != this.objPres.Slides.Count)
                this.objPres.Slides[++currentSlide].Select();
        }

        private void PrevSlide()
        {
            if (this.objPres != null && this.currentSlide != 1)
                this.objPres.Slides[--currentSlide].Select();
        }

        private void Button_Click_Prev(object sender, RoutedEventArgs e)
        {
            this.PrevSlide();
        }

        private void Button_Click_Next(object sender, RoutedEventArgs e)
        {
            this.NextSlide();  
        }

        private void On_Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (this.objPres != null)
                this.objPres.Close();
            this.ppt_app.Quit();
            this.writeText.Close();
        }
        private void printArray(int[] array, string type)
        {
            Console.Write(type);
            foreach (var item in array)
            {
                Console.Write(item + " ");
            }
            Console.WriteLine("");
        }

        private void showWindow()
        {
            if (!this.IsVisible)
                this.Show();
            
            this.Activate();
            this.Topmost = true;
            this.Topmost = false;
            this.Focus();
        }

        private void WriteToCsv()
        {
            
            for (int i = 0; i < this.size_scores.Length; i++)
            {                   
                    writeText.WriteLine(System.IO.Path.GetFileNameWithoutExtension(this.pptfiles[curr]) + "," + (i+1) + "," + this.size_scores[i] + "," + this.cantidad_scores[i]);
                    
            }
            
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Right:
                    Console.WriteLine("Right key");
                    this.NextSlide();
                    break;
                case Key.Left:
                    Console.WriteLine("Left key");
                    this.PrevSlide();
                    break;
                case Key.Q:
                    Console.WriteLine("cantidad bad");
                    this.label_cantidad.Content = "Bad";
                    this.cantidad_scores[currentSlide - 1] = 1;
                    break;
                case Key.W:
                    Console.WriteLine("cantidad regular");
                    this.label_cantidad.Content = "Regular";
                    this.cantidad_scores[currentSlide - 1] = 2;
                    break;
                case Key.E:
                    Console.WriteLine("cantidad good");
                    this.label_cantidad.Content = "Good";
                    this.cantidad_scores[currentSlide - 1] = 3;
                    break;
                case Key.A:
                    Console.WriteLine("size bad");
                    this.label_size.Content = "Bad";
                    this.size_scores[currentSlide - 1] = 1;
                    break;
                case Key.S:
                    Console.WriteLine("size regular");
                    this.label_size.Content = "Regular";
                    this.size_scores[currentSlide - 1] = 2;
                    break;
                case Key.D:
                    Console.WriteLine("size good");
                    this.label_size.Content = "Good";
                    this.size_scores[currentSlide - 1] = 3;
                    break;
                case Key.P:
                    this.printArray(this.cantidad_scores, "Cantidad texto: ");
                    this.printArray(this.size_scores, "Tamaño texto: ");
                    break;
                case Key.N:
                    this.OpenNewPresentation();
                    break;
                default:
                    Console.WriteLine("Another Key");
                    break;
            }
            

        }
    }
}
