using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;

namespace smartZoom.csproj
{
    public partial class SolidWorksMacro
    {


        public void Main()
        {
            //System.Diagnostics.Debug.Print(DateTime.Now.Ticks.ToString());
            smartZoom sz = new smartZoom(swApp, 0.1);
            sz.Zoom();
        }

        /// <summary>
        ///  The SldWorks swApp variable is pre-assigned for you.
        /// </summary>
        public SldWorks swApp;
    }
}


