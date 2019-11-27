using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Controls.Wpf;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace CopyTextOfSelectedAnnotation
{
    public static class Extensions
    {
        #region SwissAcademic.Citavi.Shell.Controls.Preview.PreviewControl

        public static PdfViewControl GetPdfViewControl(this PreviewControl previewControl)
        {
            return previewControl
                    .GetType()
                    .GetProperty("PdfViewControl", BindingFlags.Instance | BindingFlags.NonPublic)?
                    .GetValue(previewControl) as PdfViewControl;
        }


        static List<AdornmentCanvas> GetSelectedAdornmentCanvas(this PreviewControl previewControl)
        {

            return previewControl
                    .GetPdfViewControl()?
                    .Tool
                    .GetType()
                    .GetField("SelectedAdornmentContainers", BindingFlags.Instance | BindingFlags.NonPublic)?
                    .GetValue(previewControl.GetPdfViewControl()?.Tool) as List<AdornmentCanvas>;
        }

        public static ICitaviEntity GetSelectedCitaviEntity(this PreviewControl previewControl)
        {
            return previewControl?
                   .GetSelectedAdornmentCanvas()?
                   .FirstOrDefault()?
                   .CitaviEntity;
        }

        public static bool IsSupportedCitaviEntity(this ICitaviEntity citaviEntity)
        {
            return citaviEntity is KnowledgeItem || citaviEntity is TaskItem;
        }

        #endregion
    }
}
