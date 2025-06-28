using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace MsgToPdfConverter.Services
{
    public class HierarchyImageService
    {
        private const int BOX_WIDTH = 150;
        private const int BOX_HEIGHT = 40;
        private const int VERTICAL_SPACING = 60;
        private const int HORIZONTAL_SPACING = 30;
        private const int MARGIN = 20;

        public string CreateHierarchyImage(List<string> hierarchyChain, string currentAttachment, string outputFolder)
        {
            try
            {
                if (hierarchyChain == null || hierarchyChain.Count == 0)
                    return null;

                // Calculate image dimensions
                int maxLevel = hierarchyChain.Count;
                int imageWidth = (maxLevel * (BOX_WIDTH + HORIZONTAL_SPACING)) + (2 * MARGIN);
                int imageHeight = BOX_HEIGHT + (2 * MARGIN);

                // Create bitmap and graphics
                using (var bitmap = new Bitmap(imageWidth, imageHeight))
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    // Set high quality rendering
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;

                    // Fill background
                    graphics.Clear(Color.White);

                    // Define fonts and brushes
                    using (var normalFont = new Font("Arial", 9, FontStyle.Regular))
                    using (var boldFont = new Font("Arial", 9, FontStyle.Bold))
                    using (var normalBrush = new SolidBrush(Color.Black))
                    using (var currentBrush = new SolidBrush(Color.White))
                    using (var normalPen = new Pen(Color.DarkBlue, 2))
                    using (var currentPen = new Pen(Color.Red, 3))
                    using (var normalBoxBrush = new SolidBrush(Color.LightBlue))
                    using (var currentBoxBrush = new SolidBrush(Color.Red))
                    using (var linePen = new Pen(Color.DarkGray, 2))
                    {
                        // Draw hierarchy boxes and connections
                        for (int i = 0; i < hierarchyChain.Count; i++)
                        {
                            string item = hierarchyChain[i];
                            bool isCurrent = item.Equals(currentAttachment, StringComparison.OrdinalIgnoreCase);

                            // Calculate box position
                            int x = MARGIN + (i * (BOX_WIDTH + HORIZONTAL_SPACING));
                            int y = MARGIN;

                            // Draw connection line to next box (if not last)
                            if (i < hierarchyChain.Count - 1)
                            {
                                int lineStartX = x + BOX_WIDTH;
                                int lineEndX = x + BOX_WIDTH + HORIZONTAL_SPACING;
                                int lineY = y + (BOX_HEIGHT / 2);
                                
                                graphics.DrawLine(linePen, lineStartX, lineY, lineEndX, lineY);
                                
                                // Draw arrow
                                Point[] arrowPoints = {
                                    new Point(lineEndX - 8, lineY - 4),
                                    new Point(lineEndX, lineY),
                                    new Point(lineEndX - 8, lineY + 4)
                                };
                                graphics.DrawLines(linePen, arrowPoints);
                            }

                            // Draw box
                            Rectangle boxRect = new Rectangle(x, y, BOX_WIDTH, BOX_HEIGHT);
                            
                            if (isCurrent)
                            {
                                graphics.FillRectangle(currentBoxBrush, boxRect);
                                graphics.DrawRectangle(currentPen, boxRect);
                            }
                            else
                            {
                                graphics.FillRectangle(normalBoxBrush, boxRect);
                                graphics.DrawRectangle(normalPen, boxRect);
                            }

                            // Draw text
                            string displayText = TruncateText(item, normalFont, graphics, BOX_WIDTH - 10);
                            var textBrush = isCurrent ? currentBrush : normalBrush;
                            var font = isCurrent ? boldFont : normalFont;
                            
                            var textRect = new Rectangle(x + 5, y + 5, BOX_WIDTH - 10, BOX_HEIGHT - 10);
                            var stringFormat = new StringFormat
                            {
                                Alignment = StringAlignment.Center,
                                LineAlignment = StringAlignment.Center,
                                Trimming = StringTrimming.EllipsisCharacter
                            };
                            
                            graphics.DrawString(displayText, font, textBrush, textRect, stringFormat);
                        }
                    }

                    // Save the image
                    string imagePath = Path.Combine(outputFolder, $"hierarchy_{Guid.NewGuid()}.png");
                    bitmap.Save(imagePath, ImageFormat.Png);
                    return imagePath;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating hierarchy image: {ex.Message}");
                return null;
            }
        }

        private string TruncateText(string text, Font font, Graphics graphics, int maxWidth)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            // Remove file extension for cleaner display
            string displayText = Path.GetFileNameWithoutExtension(text);
            if (string.IsNullOrEmpty(displayText))
                displayText = text;

            var textSize = graphics.MeasureString(displayText, font);
            if (textSize.Width <= maxWidth)
                return displayText;

            // Truncate with ellipsis
            for (int i = displayText.Length - 1; i > 0; i--)
            {
                string truncated = displayText.Substring(0, i) + "...";
                textSize = graphics.MeasureString(truncated, font);
                if (textSize.Width <= maxWidth)
                    return truncated;
            }

            return "...";
        }
    }
}
