﻿using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using word = Microsoft.Office.Interop.Word;

namespace WordAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnPDF_Click(object sender, RibbonControlEventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = "*";
            dlg.DefaultExt = "pdf";
            dlg.ValidateNames = true;
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(dlg.FileName, word.WdExportFormat.wdExportFormatPDF, OpenAfterExport: true);
            }
        }

        private void btnImage_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.FileName = "*";
            dlg.DefaultExt = "bmp";
            dlg.ValidateNames = true;

            dlg.Filter = "Bitmap Image (.bmp)|*.bmp|Gif Image (.gif)|*.gif|JPEG Image (.jpeg)|*.jpeg|Png Image (.png)|*.png";
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Globals.ThisAddIn.Application.ActiveDocument.Shapes.AddPicture(dlg.FileName);
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0), 3, 4);

            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorSeaGreen;
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Range.Font.Size = 12;

            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Rows.Borders.Enable = 1;
        }
    }
}
