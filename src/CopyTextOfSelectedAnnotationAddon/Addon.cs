using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Citavi.Shell.Controls.Preview;
using SwissAcademic.Controls;
using SwissAcademic.Pdf.Analysis;
using System;
using System.Linq;
using System.Windows.Forms;

namespace CopyTextOfSelectedAnnotation
{
    public class Addon : CitaviAddOn<MainForm>
    {
        #region Constants

        const string Keys_Button_CopyTextOfSelectedAnnotation = "CopyTextOfSelectedAnnotation.Button.CopyTextOfSelectedAnnotation";

        #endregion

        #region Methods

        public override void OnHostingFormLoaded(MainForm mainForm)
        {
            var button = mainForm
                          .GetPreviewCommandbar(MainFormPreviewCommandbarId.Toolbar)
                          .GetCommandbarMenu(MainFormPreviewCommandbarMenuId.Tools)
                          .AddCommandbarButton(Keys_Button_CopyTextOfSelectedAnnotation, "CopyTextOfSelectedAnnotation");
            button.Shortcut = (Shortcut)(Keys.Shift | Keys.C);
            button.Visible = false;

            base.OnHostingFormLoaded(mainForm);
        }
        public override void OnBeforePerformingCommand(MainForm mainForm, BeforePerformingCommandEventArgs e)
        {
            if (e.Key.Equals(Keys_Button_CopyTextOfSelectedAnnotation, StringComparison.OrdinalIgnoreCase))
            {
                ExtractTextFromEntity(mainForm, mainForm.PreviewControl.GetSelectedCitaviEntity());
                e.Handled = true;
            }

            base.OnBeforePerformingCommand(mainForm, e);
        }

        public override void OnApplicationIdle(MainForm mainForm)
        {
            var button = mainForm
                         .GetPreviewCommandbar(MainFormPreviewCommandbarId.Toolbar)
                         .GetCommandbarMenu(MainFormPreviewCommandbarMenuId.Tools)
                         .GetCommandbarButton(Keys_Button_CopyTextOfSelectedAnnotation);
            if (button != null)
            {
                button.Tool.SharedProps.Enabled = mainForm.PreviewControl.ActivePreviewType == PreviewType.Pdf
                                                 && mainForm.PreviewControl.GetSelectedCitaviEntity().IsSupportedCitaviEntity();
            }
            base.OnApplicationIdle(mainForm);
        }

        void ExtractTextFromEntity(MainForm mainForm, ICitaviEntity citaviEntity)
        {
            if (citaviEntity.EntityLinks.FirstOrDefault(link => link.Indication.Equals("PdfKnowledgeItem", StringComparison.OrdinalIgnoreCase))?.Target is Annotation annotation)
            {
                var documentParser = new DocumentParser(mainForm.PreviewControl.GetPdfViewControl()?.Document)
                {
                    ParseType = ParseType.Text,
                    DetectParagraphAlignment = true,
                    ExtractIdentifier = false
                };

                var text = documentParser
                           .Run(annotation.Quads.Where(q => !q.IsContainer).ToList())?
                           .GetDocumentText()?
                           .ContentAsPlainText;
                if (!string.IsNullOrEmpty(text)) Clipboard.SetText(text);
            }
        }

        #endregion
    }
}