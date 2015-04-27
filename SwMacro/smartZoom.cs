using System;
using System.Collections.Generic;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace smartZoom.csproj
{
    class smartZoom
    {
        private double ulCornX;
        private double ulCornY;
        private double lrCornX;
        private double lrCornY;
        private double conversionFactor = .0254;
        private double[] vSheetProps;

        public smartZoom(SldWorks swAppIn)
        {
            this.swApp = swAppIn;
            this.BorderSize = 0.1;
        }

        public smartZoom(SldWorks swAppIn, double bs)
        {
            this.swApp = swAppIn;
            this.BorderSize = bs;
        }

        private void GetSheetProperties()
        {
            //System.Diagnostics.Debug.Print("GetSheetProperties()");
            swPart = (ModelDoc2)this.swApp.ActiveDoc;

            if (swPart != null && swPart.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
            {
                swDraw = (DrawingDoc)this.swPart;
                swSheet = (Sheet)swDraw.GetCurrentSheet();
                vSheetProps = (double[])swSheet.GetProperties();
            }
        }

        private void CalculateZoom()
        {
            ulCornX = (0 - BorderSize) * conversionFactor;
            ulCornY = vSheetProps[6] + (BorderSize * conversionFactor);
            lrCornX = vSheetProps[5] + (BorderSize * conversionFactor);
            lrCornY = (0 - BorderSize) * conversionFactor;
            //System.Diagnostics.Debug.Print(string.Format("CalculateZoom() = {0} -> {1}, {2} -> {3}"
            //    , vSheetProps[6].ToString()
            //    , ulCornY.ToString()
            //    , vSheetProps[5].ToString()
            //    , lrCornX.ToString()));
        }

        public void Zoom()
        {
            GetSheetProperties();
            //System.Diagnostics.Debug.Print("Zoom()");
            if (swPart != null && swPart.GetType() == (int)swDocumentTypes_e.swDocDRAWING)
            {
                //System.Windows.Forms.MessageBox.Show(string.Format("Drawing: {0}", swPart.GetType().ToString()));
                CalculateZoom();
                swPart.ViewZoomTo2(ulCornX, ulCornY, 0, lrCornX, lrCornY, 0);
            }
            else
            {
                //System.Windows.Forms.MessageBox.Show(string.Format("Model: {0}", swPart.GetType().ToString()));
                swPart.ViewZoomtofit2();
            }
        }

        private double _bs;

        public double BorderSize
        {
            get { return _bs; }
            set { _bs = value; }
        }

        private ModelDoc2 _swPart;

        public ModelDoc2 swPart
        {
            get { return _swPart; }
            set { _swPart = value; }
        }

        private DrawingDoc _swDraw;

        public DrawingDoc swDraw
        {
            get { return _swDraw; }
            set { _swDraw = value; }
        }

        private Sheet _swSheet;

        public Sheet swSheet
        {
            get { return _swSheet; }
            set { _swSheet = value; }
        }
	
	
        private SldWorks _swApp;

        public SldWorks swApp
        {
            get { return _swApp; }
            set { _swApp = value; }
        }

    }
}
