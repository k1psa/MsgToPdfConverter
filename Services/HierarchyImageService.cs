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
        private const int MIN_BOX_WIDTH = 180;
        private const int BOX_HEIGHT = 50;
        private const int VERTICAL_SPACING = 60;
        private const int MARGIN = 20;
        private const int TEXT_PADDING = 10;

        public string CreateHierarchyImage(List<string> hierarchyChain, string currentAttachment, string outputFolder)
        {
            try
            {
                if (hierarchyChain == null || hierarchyChain.Count == 0)
                    return null;

                // The hierarchyChain already includes the currentAttachment as the last item
                // All items except the last are emails (.msg), but we need to check if the last item
                // is actually also an email (nested message) or a real attachment
                var processedChain = new List<string>();
                for (int i = 0; i < hierarchyChain.Count; i++)
                {
                    string item = hierarchyChain[i];
                    bool isEmail;
                    
                    if (i < hierarchyChain.Count - 1)
                    {
                        // All items except the last are definitely emails
                        isEmail = true;
                    }
                    else
                    {
                        // For the last item, check if it looks like an email subject
                        isEmail = IsLikelyEmailSubject(item);
                    }
                    
                    processedChain.Add(AddFileExtension(item, isEmail));
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

                // Calculate image dimensions for vertical layout
                int maxWidth = boxWidths.Max();
                int totalWidth = maxWidth + (2 * MARGIN);
                int totalHeight = (processedChain.Count * BOX_HEIGHT) + ((processedChain.Count - 1) * VERTICAL_SPACING) + (2 * MARGIN);

                // Create high-resolution bitmap for vector-like quality
                int scale = 4; // 4x scaling for high quality
                int scaledWidth = totalWidth * scale;
                int scaledHeight = totalHeight * scale;

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
                        // Draw hierarchy boxes and connections vertically
                        int currentY = MARGIN;
                        for (int i = 0; i < processedChain.Count; i++)
                        {
                            string item = processedChain[i];
                            bool isCurrent = i == processedChain.Count - 1; // Last item is current
                            int boxWidth = boxWidths[i];
                            
                            // Center the box horizontally
                            int boxX = (totalWidth - boxWidth) / 2;

                            // Draw connection line to next box (if not last)
                            if (i < processedChain.Count - 1)
                            {
                                int lineStartY = currentY + BOX_HEIGHT;
                                int lineEndY = lineStartY + VERTICAL_SPACING;
                                int lineX = totalWidth / 2; // Center line
                                
                                graphics.DrawLine(linePen, lineX, lineStartY, lineX, lineEndY);
                                
                                // Draw arrow pointing down
                                Point[] arrowPoints = {
                                    new Point(lineX - 4, lineEndY - 8),
                                    new Point(lineX, lineEndY),
                                    new Point(lineX + 4, lineEndY - 8)
                                };
                                graphics.DrawLines(linePen, arrowPoints);
                            }

                            // Draw box
                            Rectangle boxRect = new Rectangle(boxX, currentY, boxWidth, BOX_HEIGHT);
                            
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
                            
                            var textRect = new Rectangle(boxX + TEXT_PADDING/2, currentY + TEXT_PADDING/2, 
                                                       boxWidth - TEXT_PADDING, BOX_HEIGHT - TEXT_PADDING);
                            var stringFormat = new StringFormat
                            {
                                Alignment = StringAlignment.Center,
                                LineAlignment = StringAlignment.Center,
                                Trimming = StringTrimming.EllipsisWord,
                                FormatFlags = StringFormatFlags.LineLimit
                            };
                            
                            graphics.DrawString(item, font, textBrush, textRect, stringFormat);

                            currentY += BOX_HEIGHT + VERTICAL_SPACING;
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

        private bool IsLikelyEmailSubject(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            string lowerText = text.ToLower();
            
            // Check for common email subject patterns
            if (lowerText.StartsWith("re:") || lowerText.StartsWith("fw:") || lowerText.StartsWith("fwd:"))
                return true;
                
            // Check for common email-related words
            if (lowerText.Contains("request") || lowerText.Contains("approval") || 
                lowerText.Contains("meeting") || lowerText.Contains("notification") ||
                lowerText.Contains("response") || lowerText.Contains("inquiry") ||
                lowerText.Contains("follow") || lowerText.Contains("update"))
                return true;
                
            // Check if it contains typical email subject indicators (colons, dashes in business context)
            if ((lowerText.Contains(" - ") || lowerText.Contains(": ")) && 
                (lowerText.Length > 20)) // Longer subjects are more likely emails
                return true;
                
            // If it has a file extension, it's probably a real attachment
            if (Path.HasExtension(text))
                return false;
                
            return false; // Default to attachment
        }

        private string AddFileExtension(string fileName, bool isEmail)
        {
            if (string.IsNullOrEmpty(fileName))
                return "Unknown";

            // If it's an email, always add .msg extension
            if (isEmail)
            {
                // Remove any existing extension and add .msg
                string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                if (string.IsNullOrEmpty(nameWithoutExt))
                    nameWithoutExt = fileName;
                return nameWithoutExt + ".msg";
            }

            // For attachments, ensure they have an appropriate extension
            if (!Path.HasExtension(fileName))
            {
                // Try to guess extension based on name patterns or content hints
                string lowerName = fileName.ToLower();
                
                // Check for common document types
                if (lowerName.Contains("word") || lowerName.Contains("doc") || lowerName.EndsWith("docx"))
                    return fileName + ".docx";
                else if (lowerName.Contains("excel") || lowerName.Contains("sheet") || lowerName.Contains("xls"))
                    return fileName + ".xlsx";
                else if (lowerName.Contains("powerpoint") || lowerName.Contains("ppt"))
                    return fileName + ".pptx";
                else if (lowerName.Contains("pdf"))
                    return fileName + ".pdf";
                    
                // Check for archive/compression
                else if (lowerName.Contains("zip") || lowerName.Contains("archive") || lowerName.Contains("rar"))
                    return fileName + ".zip";
                    
                // Check for images
                else if (lowerName.Contains("image") || lowerName.Contains("picture") || lowerName.Contains("photo"))
                    return fileName + ".png";
                    
                // Check for text files
                else if (lowerName.Contains("text") || lowerName.Contains("note"))
                    return fileName + ".txt";
                    
                // Check for data/email files by common patterns
                else if (lowerName.Contains("data") || lowerName.Contains("export") || lowerName.Contains("backup"))
                    return fileName + ".dat";
                    
                // Check for email-like patterns (RE:, FW:, etc.)
                else if (lowerName.StartsWith("re:") || lowerName.StartsWith("fw:") || lowerName.StartsWith("fwd:") || 
                         lowerName.Contains("request") || lowerName.Contains("approval") || lowerName.Contains("email"))
                    return fileName + ".msg";
                    
                // If it looks like a filename with numbers/letters pattern, it might be an email
                else if (System.Text.RegularExpressions.Regex.IsMatch(fileName, @"^[A-Z0-9]+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    return fileName + ".msg"; // Likely an email identifier
                else
                    return fileName + ".file"; // Generic fallback
            }

            return fileName; // Already has extension
        }
    }
}
