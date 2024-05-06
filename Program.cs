using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace PDFRenderer
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // user inputs
            if (args.Length != 6)
                throw new ArgumentOutOfRangeException();

            string _gsPath = "C:\\Program Files\\gs\\gs10.02.0\\bin\\gswin64c.exe";
            string inputPath = args[0];
            string outputPath = args[1];
            string[] pages = args[2].Split(',');
            int resolution = int.Parse(args[3]);
            bool _overprintPreview = args[4] == "1";
            char _fileType = args[5][0];

            PDFRenderHelper PDFRenderer = new PDFRenderHelper(_gsPath);
            foreach (string page in pages)
            {
                PDFRenderer.Render_(_gsPath, inputPath, outputPath, int.Parse(page), resolution, _overprintPreview, _fileType == 'm');
            }
        }
    }

    public class PDFRenderHelper
    {
        private string _script_page = "-dSAFER -dBATCH -dNOPROMPT -dQUIET -dNOPAUSE{0} -sDEVICE=png16m -dTextAlphaBits=4 -dGraphicsAlphaBits=4 -dOverprint={1} -r{4} -sOutputFile=\"{2}.png\" \"{3}\"";
        private string _script_separation = "-dSAFER -dBATCH -dNOPROMPT -dQUIET -dNOPAUSE{0} -sDEVICE=tiffsep -dTextAlphaBits=4 -dGraphicsAlphaBits=4 -dOverprint={1} -r{4} -sOutputFile=\"{2}.tiff\" \"{3}\"";
        private string _script_ai2pdf = "-dSAFER -dBATCH -dNOPROMPT -dQUIET -dNOPAUSE -sDEVICE=pdfwrite -sOutputFile=\"{0}\" \"{1}\"";
        private string _script_help = "-h";

        private string _path;
        public string path 
        { 
            get 
            { 
                return _path; 
            }
        }

        public PDFRenderHelper(string gsPath)
        {
            this._path = gsPath;
        }

        // Convert to Adobe Illustrator
        public bool ConvertoAIToPDF(string inputPath, string outputPath)
        {
            return callGS(this._path, string.Format(_script_ai2pdf, outputPath, inputPath));
        }

        // Render specific page
        public bool RenderPage(string inputPath, string outputPath, int page, int resolution, bool overprint = false, bool is_master = true)
        {
            return Render_(this._path, inputPath, outputPath, page , resolution, overprint, is_master);
        }

        internal bool Render_(string gsPath, string inputPath, string outputPath, int page, int resolution, bool overprint, bool is_master)
        {
            // default configs
            int pageNumber = page; // Render all pages: -1
            string overprintText = "/disable";
            string fileType = is_master ? "m" : "s";
            string pageCmd = string.Empty;


            // treat configs
            if (overprint)
            {
                overprintText = "/simulate";
                outputPath = Path.Combine(outputPath, "OP");
            }
            else
            {
                outputPath = Path.Combine(outputPath, "NOP");
            }

            if (pageNumber != -1)
            {
                pageCmd = string.Format(" -dFirstPage={0} -dLastPage={0}", pageNumber + 1);
            }
            string fileName = fileType + page.ToString();
            outputPath = Path.Combine(outputPath, fileName);

            return callGS(gsPath, string.Format(_script_page, pageCmd, overprintText, outputPath, inputPath, resolution));
        }

        internal bool callGS(string gsPath, string arguments)
        {
            try
            {
                // call GS
                System.Diagnostics.Process ghostRender = new System.Diagnostics.Process();
                ghostRender.StartInfo.FileName = gsPath;
                ghostRender.StartInfo.UseShellExecute = false;
                ghostRender.StartInfo.RedirectStandardOutput = true;
                ghostRender.StartInfo.RedirectStandardError = true;
                ghostRender.StartInfo.CreateNoWindow = true;
                ghostRender.StartInfo.Arguments = arguments;
                ghostRender.Start();
                ghostRender.BeginErrorReadLine();
                string processOutput = ghostRender.StandardOutput.ReadToEnd();
                ghostRender.WaitForExit();

                if (!string.IsNullOrWhiteSpace(processOutput))
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }
    }
}
