using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MsgReader.Outlook;

namespace MsgToPdfConverter.Services
{
    public class EmailConverterService
    {
        // Helper method to generate DejaVu Sans font style with base64 embedding
        private string GetDejaVuSansFontStyle()
        {
            try
            {
                string fontPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "fonts", "DejaVuSans.ttf");
                if (!File.Exists(fontPath))
                {
                    Console.WriteLine($"[FONT] DejaVu Sans font not found at: {fontPath}");
                    return @"
                        <style>
                        html, body, table, div, span, p, td, th, b, i, u, strong, em, h1, h2, h3, h4, h5, h6 {
                            font-family: Arial, sans-serif !important;
                        }
                        </style>";
                }

                // Read font file and convert to base64
                byte[] fontBytes = File.ReadAllBytes(fontPath);
                string base64Font = Convert.ToBase64String(fontBytes);

                Console.WriteLine($"[FONT] Embedded DejaVu Sans font as base64 data URI ({fontBytes.Length} bytes)");

                return $@"
                    <style>
                    @font-face {{
                        font-family: 'DejaVu Sans';
                        src: url('data:font/truetype;charset=utf-8;base64,{base64Font}') format('truetype');
                        font-weight: normal;
                        font-style: normal;
                    }}
                    html, body, table, div, span, p, td, th, b, i, u, strong, em, h1, h2, h3, h4, h5, h6 {{
                        font-family: 'DejaVu Sans', Arial, sans-serif !important;
                    }}
                    </style>";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[FONT] Error embedding DejaVu Sans font: {ex.Message}");
                return @"
                    <style>
                    html, body, table, div, span, p, td, th, b, i, u, strong, em, h1, h2, h3, h4, h5, h6 {
                        font-family: Arial, sans-serif !important;
                    }
                    </style>";
            }
        }

        public string BuildEmailHtml(Storage.Message msg, bool extractOriginalOnly = false)
        {
            // Build proper From field with both name and email
            string from = "";
            if (msg.Sender != null)
            {
                if (!string.IsNullOrEmpty(msg.Sender.DisplayName) && !string.IsNullOrEmpty(msg.Sender.Email))
                {
                    from = $"{msg.Sender.DisplayName} <{msg.Sender.Email}>";
                }
                else if (!string.IsNullOrEmpty(msg.Sender.DisplayName))
                {
                    from = msg.Sender.DisplayName;
                }
                else if (!string.IsNullOrEmpty(msg.Sender.Email))
                {
                    from = msg.Sender.Email;
                }
            }

            string sent = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("f") : "";
            string to = string.Join(", ", msg.Recipients?.FindAll(r => r.Type == Storage.Recipient.RecipientType.To)?.ConvertAll(r => r.DisplayName + (string.IsNullOrEmpty(r.Email) ? "" : $" <{r.Email}>")) ?? new List<string>());
            string cc = string.Join(", ", msg.Recipients?.FindAll(r => r.Type == Storage.Recipient.RecipientType.Cc)?.ConvertAll(r => r.DisplayName + (string.IsNullOrEmpty(r.Email) ? "" : $" <{r.Email}>")) ?? new List<string>());
            string subject = msg.Subject ?? "";
            string body = GetEmailBodyWithProperEncoding(msg) ?? "";

            if (extractOriginalOnly)
            {
                body = ExtractOriginalEmailContent(body);
            }

            string attachmentsLine = BuildAttachmentsLine(msg);

            // Embed DejaVu Sans font
            string fontStyle = GetDejaVuSansFontStyle();

            string header =
                "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:12pt;margin-bottom:16px;'>" +
                $"<div><b>From:</b> {System.Net.WebUtility.HtmlEncode(from)}</div>" +
                $"<div><b>Sent:</b> {System.Net.WebUtility.HtmlEncode(sent)}</div>" +
                $"<div><b>To:</b> {System.Net.WebUtility.HtmlEncode(to)}</div>" +
                (string.IsNullOrWhiteSpace(cc) ? "" : $"<div><b>Cc:</b> {System.Net.WebUtility.HtmlEncode(cc)}</div>") +
                $"<div><b>Subject:</b> {System.Net.WebUtility.HtmlEncode(subject)}</div>" +
                attachmentsLine +
                "</div>";

            return "<!DOCTYPE html>" +
                   "<html>" +
                   "<head>" +
                   "<meta charset=\"UTF-8\">" +
                   "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">" +
                   "<title>Email</title>" +
                   fontStyle +
                   "</head>" +
                   "<body>" +
                   header + body +
                   "</body>" +
                   "</html>";
        }

        public string BuildAttachmentsLine(Storage.Message msg)
        {
            if (msg.Attachments == null || msg.Attachments.Count == 0)
                return string.Empty;

            var inlineContentIds = GetInlineContentIds(msg.BodyHtml ?? "");
            var attachmentNames = new List<string>();

            attachmentNames.AddRange(msg.Attachments
                .OfType<Storage.Attachment>()
                .Where(a =>
                    !string.IsNullOrEmpty(a.FileName) &&
                    (string.IsNullOrEmpty(a.ContentId) || !inlineContentIds.Contains(a.ContentId.Trim('<', '>', '"', '\'', ' '))) &&
                    !IsLikelySignatureImage(a) &&
                    !new[] { ".p7s", ".p7m", ".smime", ".asc", ".sig" }.Contains(Path.GetExtension(a.FileName).ToLowerInvariant())
                )
                .Select(a => System.Net.WebUtility.HtmlEncode(a.FileName)));

            attachmentNames.AddRange(msg.Attachments
                .OfType<Storage.Message>()
                .Select(nestedMsg => System.Net.WebUtility.HtmlEncode(nestedMsg.Subject ?? "[Attached Email]")));

            if (attachmentNames.Count > 0)
            {
                return $"<div><b>Attachments:</b> {string.Join(", ", attachmentNames)}</div>";
            }
            return string.Empty;
        }

        /// <summary>
        /// Gets email body with proper encoding handling for Unicode characters like Greek
        /// </summary>
        public string GetEmailBodyWithProperEncoding(Storage.Message msg)
        {
            try
            {
                // Try to get HTML body first
                string htmlBody = msg.BodyHtml;
                if (!string.IsNullOrEmpty(htmlBody))
                {
                    // Check if the HTML body appears to have encoding issues
                    if (HasEncodingIssues(htmlBody))
                    {
                        // Try to re-interpret with different encodings
                        string fixedHtml = TryFixEncoding(htmlBody);
                        if (!string.IsNullOrEmpty(fixedHtml) && !HasEncodingIssues(fixedHtml))
                        {
                            return fixedHtml;
                        }
                    }
                    return htmlBody;
                }

                // Fall back to text body
                string textBody = msg.BodyText;
                if (!string.IsNullOrEmpty(textBody))
                {
                    if (HasEncodingIssues(textBody))
                    {
                        string fixedText = TryFixEncoding(textBody);
                        if (!string.IsNullOrEmpty(fixedText) && !HasEncodingIssues(fixedText))
                        {
                            return fixedText;
                        }
                    }
                    return textBody;
                }

                return string.Empty;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ENCODING] Error getting body with proper encoding: {ex.Message}");
                return msg.BodyHtml ?? msg.BodyText ?? string.Empty;
            }
        }

        /// <summary>
        /// Checks if text has encoding issues (like Greek characters showing as garbage)
        /// </summary>
        public bool HasEncodingIssues(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            // Look for patterns that indicate encoding issues
            // Do NOT treat the Euro sign ("€") as an encoding issue
            // Common encoding artifacts: "Ã", "Î", "Ï", "â", "™" (but not "€")
            return text.Contains("Ã") || text.Contains("Î") || text.Contains("Ï") ||
                   text.Contains("â") || text.Contains("™");
        }

        /// <summary>
        /// Attempts to fix encoding issues by trying different encoding interpretations
        /// </summary>
        public string TryFixEncoding(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            try
            {
                // Try converting from different encodings to UTF-8
                var encodings = new[]
                {
                    System.Text.Encoding.GetEncoding("windows-1252"),
                    System.Text.Encoding.GetEncoding("iso-8859-1"),
                    System.Text.Encoding.GetEncoding("iso-8859-7"), // Greek
                    System.Text.Encoding.UTF8
                };

                foreach (var encoding in encodings)
                {
                    try
                    {
                        // Convert string back to bytes using current encoding assumption
                        byte[] bytes = System.Text.Encoding.GetEncoding("iso-8859-1").GetBytes(text);
                        // Reinterpret as target encoding
                        string result = encoding.GetString(bytes);

                        // Check if this looks better (has fewer encoding issue patterns)
                        if (!HasEncodingIssues(result) && result != text)
                        {
                            return result;
                        }
                    }
                    catch
                    {
                        // Continue to next encoding
                    }
                }
            }
            catch
            {
                // Ignore errors and return original
            }

            return text; // Return original if no fix found
        }

        // Returns all ContentIds referenced as inline images in the HTML
        public HashSet<string> GetInlineContentIds(string html)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(html)) return set;
            // Match src='cid:...' or src="cid:..." and optional angle brackets
            var regex = new System.Text.RegularExpressions.Regex("<img[^>]+src=['\"]cid:<?([^'\">]+)>?['\"]", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            foreach (System.Text.RegularExpressions.Match match in regex.Matches(html))
            {
                if (match.Groups.Count > 1)
                    set.Add(match.Groups[1].Value.Trim('<', '>', '\"', '\'', ' '));
            }
            return set;
        }

        // Extracts the original email content from a reply/forward chain
        public string ExtractOriginalEmailContent(string emailBody)
        {
            if (string.IsNullOrEmpty(emailBody))
                return emailBody;

            // Special case: Outlook/Word HTML reply block
            var specialMarker = "<div id=\"mail-editor-reference-message-container\"";
            int specialIdx = emailBody.IndexOf(specialMarker, StringComparison.OrdinalIgnoreCase);
            if (specialIdx > 0)
            {
                return emailBody.Substring(0, specialIdx).Trim();
            }

            var replyIndicators = new[]
            {
                @"-----Original Message-----",
                @"From:.*Sent:.*To:.*Subject:",
                @"On .* wrote:",
                @"On .* at .* .* wrote:",
                @"> .*wrote:",
                @"<.*> wrote:",
                @"From: .*[\r\n]+.*Sent: .*[\r\n]+.*To: .*[\r\n]+.*Subject:",
                @"________________________________",
                @"From:.*[\r\n]Sent:.*[\r\n]To:.*[\r\n]Subject:",
                @"Begin forwarded message:",
                @"---------- Forwarded message ----------",
                @"Forwarded Message",
                @"FW:",
                @"Fwd:",
                "<div class=\"gmail_quote\">",
                "<div class=\"OutlookMessageHeader\">",
                @"<div.*class.*quoted.*>",
                @"<blockquote.*>",
                @"<hr.*>.*From:",
                @"<div.*outlook.*>.*From:",
                @"^-{5,}.*$",
                @"^_{5,}.*$",
                @"^={5,}.*$"
            };

            string originalContent = emailBody;
            int earliestIndex = originalContent.Length;
            foreach (var pattern in replyIndicators)
            {
                try
                {
                    var matches = System.Text.RegularExpressions.Regex.Matches(originalContent, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Multiline | System.Text.RegularExpressions.RegexOptions.Singleline);
                    if (matches.Count > 0)
                    {
                        var firstMatch = matches[0];
                        if (firstMatch.Index < earliestIndex)
                        {
                            earliestIndex = firstMatch.Index;
                        }
                    }
                }
                catch { }
            }
            if (earliestIndex < originalContent.Length)
            {
                originalContent = originalContent.Substring(0, earliestIndex).Trim();
            }
            // Remove trailing empty divs, paragraphs, or line breaks
            if (originalContent.Contains("<") && originalContent.Contains(">"))
            {
                originalContent = System.Text.RegularExpressions.Regex.Replace(originalContent, @"(<br\s*/?>|<p\s*>|<div\s*>|\s)*$", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            }
            return originalContent;
        }

        /// <summary>
        /// Builds email HTML and replaces inline image cids with temp file paths. Returns the HTML and a list of temp files to clean up.
        /// </summary>
        public (string Html, List<string> TempFiles) BuildEmailHtmlWithInlineImages(Storage.Message msg, bool extractOriginalOnly = false)
        {
            // Build proper From field with both name and email
            string from = "";
            if (msg.Sender != null)
            {
                if (!string.IsNullOrEmpty(msg.Sender.DisplayName) && !string.IsNullOrEmpty(msg.Sender.Email))
                {
                    from = $"{msg.Sender.DisplayName} <{msg.Sender.Email}>";
                }
                else if (!string.IsNullOrEmpty(msg.Sender.DisplayName))
                {
                    from = msg.Sender.DisplayName;
                }
                else if (!string.IsNullOrEmpty(msg.Sender.Email))
                {
                    from = msg.Sender.Email;
                }
            }

            string sent = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("f") : "";
            string to = string.Join(", ", msg.Recipients?.FindAll(r => r.Type == Storage.Recipient.RecipientType.To)?.ConvertAll(r => r.DisplayName + (string.IsNullOrEmpty(r.Email) ? "" : $" <{r.Email}>")) ?? new List<string>());
            string cc = string.Join(", ", msg.Recipients?.FindAll(r => r.Type == Storage.Recipient.RecipientType.Cc)?.ConvertAll(r => r.DisplayName + (string.IsNullOrEmpty(r.Email) ? "" : $" <{r.Email}>")) ?? new List<string>());
            string subject = msg.Subject ?? "";
            string body = GetEmailBodyWithProperEncoding(msg) ?? "";

            if (extractOriginalOnly)
            {
                body = ExtractOriginalEmailContent(body);
            }

            string attachmentsLine = BuildAttachmentsLine(msg);

            // Inline image handling
            var tempFiles = new List<string>();
            if (!string.IsNullOrEmpty(body) && msg.Attachments != null && msg.Attachments.Count > 0)
            {
                var inlineContentIds = GetInlineContentIds(body);
                foreach (var att in msg.Attachments.OfType<Storage.Attachment>())
                {
                    if (!string.IsNullOrEmpty(att.ContentId) && inlineContentIds.Contains(att.ContentId.Trim('<', '>', '\"', '\'', ' ')))
                    {
                        // Save the inline image to a temp file
                        string ext = Path.GetExtension(att.FileName) ?? ".bin";
                        string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ext);
                        File.WriteAllBytes(tempFile, att.Data);
                        tempFiles.Add(tempFile);
                        // Replace src="cid:..." with src="file:///..."
                        string cidPattern = $"cid:{att.ContentId.Trim('<', '>', '\"', '\'', ' ')}";
                        body = body.Replace($"src=\"{cidPattern}\"", $"src=\"file:///{tempFile.Replace("\\", "/")}\"")
                                   .Replace($"src='{cidPattern}'", $"src='file:///{tempFile.Replace("\\", "/")}'");
                    }
                }
            }

            // Embed DejaVu Sans font
            string fontStyle = GetDejaVuSansFontStyle();

            string header =
                "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:12pt;margin-bottom:16px;'>" +
                $"<div><b>From:</b> {System.Net.WebUtility.HtmlEncode(from)}</div>" +
                $"<div><b>Sent:</b> {System.Net.WebUtility.HtmlEncode(sent)}</div>" +
                $"<div><b>To:</b> {System.Net.WebUtility.HtmlEncode(to)}</div>" +
                (string.IsNullOrWhiteSpace(cc) ? "" : $"<div><b>Cc:</b> {System.Net.WebUtility.HtmlEncode(cc)}</div>") +
                $"<div><b>Subject:</b> {System.Net.WebUtility.HtmlEncode(subject)}</div>" +
                attachmentsLine +
                "</div>";

            string html = "<!DOCTYPE html>" +
                   "<html>" +
                   "<head>" +
                   "<meta charset=\"UTF-8\">" +
                   "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">" +
                   "<title>Email</title>" +
                   fontStyle +
                   "</head>" +
                   "<body>" +
                   header + body +
                   "</body>" +
                   "</html>";

            // Save HTML to temp file for debugging
            try
            {
                string debugHtmlPath = Path.Combine(Path.GetTempPath(), $"debug_email_{DateTime.Now:yyyyMMdd_HHmmss}.html");
                File.WriteAllText(debugHtmlPath, html, System.Text.Encoding.UTF8);
                Console.WriteLine($"[DEBUG-HTML] Saved generated HTML to: {debugHtmlPath}");
                Console.WriteLine($"[DEBUG-HTML] Sample body content: {(body?.Length > 200 ? body.Substring(0, 200) + "..." : body)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG-HTML] Failed to save debug HTML: {ex.Message}");
            }

            return (html, tempFiles);
        }

        /// <summary>
        /// Determines if an attachment is likely a signature image or decorative element that should be skipped
        /// </summary>
        private bool IsLikelySignatureImage(Storage.Attachment attachment)
        {
            try
            {
                string fileName = attachment.FileName ?? "";
                string ext = Path.GetExtension(fileName).ToLowerInvariant();

                // Only check image files
                if (ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".gif" && ext != ".bmp")
                {
                    return false; // Not an image, so not a signature image
                }

                // Check file size - signature images are typically small (less than 50KB)
                int fileSizeKB = (attachment.Data?.Length ?? 0) / 1024;
                bool isSmallImage = fileSizeKB < 50;

                // Check for common signature image patterns in filename
                string lowerFileName = fileName.ToLowerInvariant();
                bool hasSignaturePattern = lowerFileName.Contains("image") ||
                                         lowerFileName.Contains("signature") ||
                                         lowerFileName.Contains("logo") ||
                                         lowerFileName.Contains("banner") ||
                                         lowerFileName.StartsWith("oledata.mso");

                // If it's a small image with signature patterns, likely a signature
                if (isSmallImage && hasSignaturePattern)
                {
                    return true;
                }

                // If it's marked as inline AND small, likely decorative/signature
                if (attachment.IsInline == true && isSmallImage)
                {
                    return true;
                }

                return false;
            }
            catch (Exception)
            {
                return false; // If in doubt, don't filter out
            }
        }
    }
}
