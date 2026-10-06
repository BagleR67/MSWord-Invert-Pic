using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace wordPicInvertPlugin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void invertButton_Click(object sender, RibbonControlEventArgs e)
        {
            // опреление приложения
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Selection selection = app.Selection;
            //проверка на выделение
            Word.InlineShape inlineShape = null;
            if (selection.InlineShapes.Count > 0)
            {
                inlineShape = selection.InlineShapes[1];
            }

            if (inlineShape == null)
            {
                MessageBox.Show("Пожалуйста, сначала выберите картинку в документе.");
                return;
            }

            try
            {
                inlineShape.Select();
                app.Selection.Copy();


                // если буфер обмена содержит картинку: сохраняем оригинальную картинку и использем метод
                // ApplyNegativeEffect на оригинальную картинку. Сохраняем инвертированную картинку в буфер обмена и вставляем в 
                // документ с помощью app.Selection.Paste

                if (Clipboard.ContainsImage())
                {
                    Image originalImage = Clipboard.GetImage();

                    Bitmap negativeBitmap = ApplyNegativeEffect(new Bitmap(originalImage));
                    Clipboard.SetImage(negativeBitmap);

                    app.Selection.Paste();
                    originalImage.Dispose();
                    negativeBitmap.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при обработке изображения: " + ex.Message);
            }
        }
        private Bitmap ApplyNegativeEffect(Bitmap sourceVisual) // собстевенно метод ниггативизации
        {
            Bitmap result = new Bitmap(sourceVisual.Width, sourceVisual.Height);

            using (Graphics g = Graphics.FromImage(result))
            {
                ColorMatrix colorMatrix = new ColorMatrix(new float[][] // колор матрица для цвета собственно че не понятного
                {
            new float[] {-1,  0,  0,  0,  0}, // Red
            new float[] { 0, -1,  0,  0,  0}, // Green
            new float[] { 0,  0, -1,  0,  0}, // Blue
            new float[] { 0,  0,  0,  1,  0}, // Alpha
            new float[] { 1,  1,  1,  0,  1}  // Смещение
                });

                using (ImageAttributes attributes = new ImageAttributes()) // перерисовывает фото с цветами из colorMatrix
                {
                    attributes.SetColorMatrix(colorMatrix);
                    g.DrawImage(sourceVisual, new System.Drawing.Rectangle(0, 0, sourceVisual.Width, sourceVisual.Height),
                                0, 0, sourceVisual.Width, sourceVisual.Height, GraphicsUnit.Pixel, attributes);
                }
            }
            return result;
        }
    }
}
