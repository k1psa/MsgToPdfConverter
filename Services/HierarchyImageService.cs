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
        private const int MIN_BOX_WIDTH = 120;
        private const int BOX_HEIGHT = 50;
        private const int VERTICAL_SPACING = 70;
        private const int HORIZONTAL_SPACING = 30;
        private const int MARGIN = 20;
        private const int TEXT_PADDING = 10;

        public string CreateHierarchyImage(List<string> hierarchyChain, string currentAttachment, string outputFolder)
        {
            try
            {
                if (hierarchyChain == null || hierarchyChain.Count == 0)
                    return null;

                // Add proper file extensions
                var processedChain = new List<string>();
                for (int i = 0; i < hierarchyChain.Count; i++)
                {
                    string item = hierarchyChain[i];
                    if (i == 0) // First item is always the email
                    {
                        processedChain.Add(AddFileExtension(item, true));
                    }
                    else
                    {
                        processedChain.Add(AddFileExtension(item, false));
                    }
                }

                // Calculate box widths based on text content
                var boxWidths = new List<int>();
                using (var tempBitmap = new Bitmap(1, 1))
                using (var tempGraphics = Graphics.FromImage(tempBitmap))
                using (var font = new Font("Arial", 10, FontStyle.Regular))
                {
                    foreach (string item in processedChain)
                    {
                        var textSize = tempGraphics.MeasureString(item, font);
                        int width = Math.Max(MIN_BOX_WIDTH, (int)textSize.Width + (TEXT_PADDING * 2));
                        boxWidths.Add(width);
                    }
                }

                // Calculate total image dimensions
                int totalWidth = boxWidths.Sum() + (HORIZONTAL_SPACING * (boxWidths.Count - 1)) + (2 * MARGIN);
                int imageHeight = BOX_HEIGHT + (2 * MARGIN);

                // Create high-resolution bitmap for vector-like quality
                int scale = 4; // 4x scaling for high quality
                int scaledWidth = totalWidth * scale;
                int scaledHeight = imageHeight * scale;

                using (var bitmap = new Bitmap(scaledWidth, scaledHeight))
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    // Set highest quality rendering
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                    // Scale everything
                    graphics.ScaleTransform(scale, scale);

                    // Fill background
                    graphics.Clear(Color.White);

                    // Define fonts and brushes (scaled)
                    using (var normalFont = new Font("Arial", 10, FontStyle.Regular))
                    using (var boldFont = new Font("Arial", 10, FontStyle.Bold))
                    using (var normalBrush = new SolidBrush(Color.Black))
                    using (var currentBrush = new SolidBrush(Color.White))
                    using (var normalPen = new Pen(Color.DarkBlue, 2))
                    using (var currentPen = new Pen(Color.Red, 3))
                    using (var normalBoxBrush = new SolidBrush(Color.LightBlue))
                    using (var currentBoxBrush = new SolidBrush(Color.Red))
                    using (var linePen = new Pen(Color.DarkGray, 2))
                    {
                        // Draw hierarchy boxes and connections
                        int currentX = MARGIN;
                        for (int i = 0; i < processedChain.Count; i++)
                        {
                            string item = processedChain[i];
                            bool isCurrent = i == processedChain.Count - 1; // Last item is current
                            int boxWidth = boxWidths[i];

                            // Draw connection line to next box (if not last)
                            if (i < processedChain.Count - 1)
                            {
                                int lineStartX = currentX + boxWidth;
                                int lineEndX = lineStartX + HORIZONTAL_SPACING;
                                int lineY = MARGIN + (BOX_HEIGHT / 2);
                                
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
                            Rectangle boxRect = new Rectangle(currentX, MARGIN, boxWidth, BOX_HEIGHT);
                            
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

                            // Draw text with word wrapping if needed
                            var textBrush = isCurrent ? currentBrush : normalBrush;
                            var font = isCurrent ? boldFont : normalFont;
                            
                            var textRect = new Rectangle(currentX + TEXT_PADDING/2, MARGIN + TEXT_PADDING/2, 
                                                       boxWidth - TEXT_PADDING, BOX_HEIGHT - TEXT_PADDING);
                            var stringFormat = new StringFormat
                            {
                                Alignment = StringAlignment.Center,
                                LineAlignment = StringAlignment.Center,
                                Trimming = StringTrimming.EllipsisWord,
                                FormatFlags = StringFormatFlags.LineLimit
                            };
                            
                            graphics.DrawString(item, font, textBrush, textRect, stringFormat);

                            currentX += boxWidth + HORIZONTAL_SPACING;
                        }
                    }

                    // Save the high-quality image
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

        private string AddFileExtension(string fileName, bool isEmail)
        {
            if (string.IsNullOrEmpty(fileName))
                return "Unknown";

            // If it's an email, add .msg extension
            if (isEmail)
            {
                return fileName + ".msg";
            }

            // For attachments, ensure they have an extension
            if (!Path.HasExtension(fileName))
            {
                // Try to guess extension based on name patterns
                string lowerName = fileName.ToLower();
                if (lowerName.Contains("word") || lowerName.Contains("doc"))
                    return fileName + ".docx";
                else if (lowerName.Contains("excel") || lowerName.Contains("sheet"))
                    return fileName + ".xlsx";
                else if (lowerName.Contains("pdf"))
                    return fileName + ".pdf";
                else if (lowerName.Contains("zip") || lowerName.Contains("archive"))
                    return fileName + ".zip";
                else if (lowerName.Contains("image") || lowerName.Contains("picture"))
                    return fileName + ".png";
                else
                    return fileName + ".file"; // Generic extension
            }

            return fileName; // Already has extension
        }
    }
}
