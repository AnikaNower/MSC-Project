using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace QCAutomationFramework.Utils
{
    public class PresentationUtils
    {
        public string PresentationName { get; set; }
        public string PresentationPath { get; set; }

        public PowerPoint.Application PPApplication { get; set; }
        public PowerPoint.Presentations oPPTPres { get; set; }
        public PowerPoint.Presentation oPres { get; set; }
        public PowerPoint.Shapes Shapes { get; set; }

        public PresentationUtils(string filePath)
        {
            PPApplication = new PowerPoint.Application();
            oPPTPres = PPApplication.Presentations;
            this.PresentationPath = filePath;
            oPres = oPPTPres.Open(filePath, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse);
        }

        public bool CheckSlideCount(PresentationUtils template)
        {
            if (this.oPres.Slides.Count == template.oPres.Slides.Count)
                return true;
            else
                return false;
        }

        public IList<string> CheckShapeCount(PowerPoint.Slides templateSlides)
        {
            IList<string> shapeCountResult = new List<string>();
            int slideNo = 0;
            foreach (PowerPoint.Slide templateSlide in templateSlides)
            {
                slideNo++;
                if (this.oPres.Slides[slideNo].Shapes.Count != templateSlide.Shapes.Count)
                {
                    shapeCountResult.Add("Shape Count Mismatch for Slide: " + slideNo + "--> Original:" + templateSlide.Shapes.Count + "  & Generated:" + this.oPres.Slides[slideNo].Shapes.Count);
                    foreach (PowerPoint.Shape shape in this.oPres.Slides[slideNo].Shapes)
                    {
                        System.Windows.Forms.MessageBox.Show(shape.Name + ": " + shape.Type.ToString());
                    }
                }
            }
            if (shapeCountResult.Count == 0)
                shapeCountResult.Add("Shape Count Match for all Slides of " + this.PresentationPath);
            return shapeCountResult;
        }

        public void CloseApplication()
        {
            PPApplication.Quit();
        }
    }
}
