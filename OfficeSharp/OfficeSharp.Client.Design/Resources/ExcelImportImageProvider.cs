using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Globalization;
using System.Windows.Media.Imaging;
using Microsoft.LightSwitch.BaseServices.ResourceService;

namespace OfficeSharp.Resources
{
    namespace Resources
    {
        [Export(typeof(IResourceProvider))]
        [ResourceProvider("OfficeIntegration.ExcelImport")]
        public class ExcelImportImageProvider : IResourceProvider
        {
            #region "IResourceProvider Members"
            public object GetResource(string resourceId, CultureInfo cultureInfo)
            {
                return new BitmapImage(new Uri("/OfficeIntegration.Client.Design;component/Resources/ControlImages/ExcelImport.png", UriKind.Relative));
            }
            #endregion
        }
    }
}
