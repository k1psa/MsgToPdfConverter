using System;
using System.Collections.Generic;
using System.Windows;
using MsgToPdfConverter.Utils;
using System.IO;
using MsgReader.Outlook;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Diagnostics;
using DinkToPdf;
using DinkToPdf.Contracts;
using System.Text.RegularExpressions;

namespace MsgToPdfConverter
{
    public partial class MainWindow : Window
    {
        private List<string> selectedFiles = new List<string>();

        public MainWindow()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            InitializeComponent();
            CheckDotNetRuntime();
        }

        private void CheckDotNetRuntime()
        {
            // Only check if not running in design mode
            if (!System.ComponentModel.DesignerProperties.GetIsInDesignMode(this))
            {
                if (!IsDotNetDesktopRuntimeInstalled())
                {
                    var result = MessageBox.Show(
                        ".NET Desktop Runtime 5.0 is required to run this application. Would you like to download it now?",
                        ".NET Runtime Required",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                    {
                        string url = "https://dotnet.microsoft.com/en-us/download/dotnet/5.0/runtime";
                        try
                        {
                            Process.Start("explorer", url);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Could not open browser. Please visit this URL manually:\n{url}\nError: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        Application.Current.Shutdown();
                    }
                    else
                    {
                        Application.Current.Shutdown();
                    }
                }
            }
        }

        private bool IsDotNetDesktopRuntimeInstalled()
        {
            // Simple check: look for a known .NET 5+ runtime folder
            string windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string dotnetDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet", "shared", "Microsoft.WindowsDesktop.App");
            if (Directory.Exists(dotnetDir))
            {
                var versions = Directory.GetDirectories(dotnetDir);
                foreach (var v in versions)
                {
                    if (v.Contains("5.0")) return true;
                }
            }
            return false;
        }

        private void SelectFilesButton_Click(object sender, RoutedEventArgs e)
        {
            selectedFiles = FileDialogHelper.OpenMsgFileDialog();
            FilesListBox.Items.Clear();
            if (selectedFiles != null && selectedFiles.Count > 0)
            {
                foreach (var file in selectedFiles)
                {
                    FilesListBox.Items.Add(file);
                }
                ConvertButton.IsEnabled = true;
            }
            else
            {
                ConvertButton.IsEnabled = false;
            }
        }

        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            selectedFiles = FileDialogHelper.OpenMsgFolderDialog();
            FilesListBox.Items.Clear();
            if (selectedFiles != null && selectedFiles.Count > 0)
            {
                foreach (var file in selectedFiles)
                {
                    FilesListBox.Items.Add(file);
                }
                ConvertButton.IsEnabled = true;
            }
            else
            {
                ConvertButton.IsEnabled = false;
            }
        }

        private string GetMimeTypeFromFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return "image/png";
            string ext = System.IO.Path.GetExtension(fileName).ToLowerInvariant();
            switch (ext)
            {
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".png": return "image/png";
                case ".gif": return "image/gif";
                case ".bmp": return "image/bmp";
                case ".tif":
                case ".tiff": return "image/tiff";
                case ".svg": return "image/svg+xml";
                default: return "image/png";
            }
        }

        private string EmbedInlineImages(Storage.Message msg)
        {
            string html = msg.BodyHtml ?? msg.BodyText;
            if (string.IsNullOrEmpty(html) || msg.Attachments == null || msg.Attachments.Count == 0)
                return html;

            var regex = new Regex("<img[^>]+src=\"cid:([^\"]+)\"", RegexOptions.IgnoreCase);
            return regex.Replace(html, match =>
            {
                string cid = match.Groups[1].Value;
                Storage.Attachment found = null;
                foreach (var att in msg.Attachments)
                {
                    if (att is Storage.Attachment attachment && attachment.ContentId != null && attachment.ContentId.Trim('<', '>') == cid.Trim('<', '>'))
                    {
                        found = attachment;
                        break;
                    }
                }
                if (found != null)
                {
                    string mimeType = GetMimeTypeFromFileName(found.FileName);
                    string base64 = Convert.ToBase64String(found.Data);
                    return match.Value.Replace($"cid:{cid}", $"data:{mimeType};base64,{base64}");
                }
                return match.Value;
            });
        }

        private string BuildEmailHtml(Storage.Message msg)
        {
            string from = msg.Sender?.DisplayName ?? msg.Sender?.Email ?? "";
            string sent = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("f") : "";
            string to = string.Join(", ", msg.Recipients?.FindAll(r => r.Type == Storage.Recipient.RecipientType.To)?.ConvertAll(r => r.DisplayName + (string.IsNullOrEmpty(r.Email) ? "" : $" <{r.Email}>")) ?? new List<string>());
            string cc = string.Join(", ", msg.Recipients?.FindAll(r => r.Type == Storage.Recipient.RecipientType.Cc)?.ConvertAll(r => r.DisplayName + (string.IsNullOrEmpty(r.Email) ? "" : $" <{r.Email}>")) ?? new List<string>());
            string subject = msg.Subject ?? "";
            string body = EmbedInlineImages(msg) ?? "";

            string header = $@"
                <div style='font-family:Segoe UI,Arial,sans-serif;font-size:12pt;margin-bottom:16px;'>
                    <div><b>From:</b> {System.Net.WebUtility.HtmlEncode(from)}</div>
                    <div><b>Sent:</b> {System.Net.WebUtility.HtmlEncode(sent)}</div>
                    <div><b>To:</b> {System.Net.WebUtility.HtmlEncode(to)}</div>
                    {(string.IsNullOrWhiteSpace(cc) ? "" : $"<div><b>Cc:</b> {System.Net.WebUtility.HtmlEncode(cc)}</div>")}
                    <div><b>Subject:</b> {System.Net.WebUtility.HtmlEncode(subject)}</div>
                </div>";
            return header + body;
        }

        private async void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            if (selectedFiles == null || selectedFiles.Count == 0)
            {
                MessageBox.Show("No files selected.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            int success = 0, fail = 0;
            ProgressBar.Visibility = Visibility.Visible;
            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = selectedFiles.Count;
            ProgressBar.Value = 0;
            var converter = new SynchronizedConverter(new PdfTools());

            await System.Threading.Tasks.Task.Run(() =>
            {
                for (int i = 0; i < selectedFiles.Count; i++)
                {
                    var msgFilePath = selectedFiles[i];
                    try
                    {
                        var msg = new Storage.Message(msgFilePath);
                        string datePart = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmms");
                        string baseName = Path.GetFileNameWithoutExtension(msgFilePath);
                        string dir = Path.GetDirectoryName(msgFilePath);
                        string pdfFilePath = Path.Combine(dir, $"{baseName} - {datePart}.pdf");
                        string htmlWithHeader = BuildEmailHtml(msg);
                        var doc = new HtmlToPdfDocument()
                        {
                            GlobalSettings = new GlobalSettings
                            {
                                ColorMode = ColorMode.Color,
                                Orientation = Orientation.Portrait,
                                PaperSize = PaperKind.A4,
                                Out = pdfFilePath
                            },
                            Objects = {
                                new ObjectSettings
                                {
                                    PagesCount = true,
                                    HtmlContent = htmlWithHeader,
                                    WebSettings = { DefaultEncoding = "utf-8" }
                                }
                            }
                        };
                        converter.Convert(doc);
                        success++;
                    }
                    catch (Exception ex)
                    {
                        fail++;
                        // Show error on UI thread
                        Dispatcher.Invoke(() =>
                        {
                            MessageBox.Show($"Failed to convert: {msgFilePath}\nError: {ex.Message}", "Conversion Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        });
                    }
                    // Update progress bar on UI thread
                    Dispatcher.Invoke(() => ProgressBar.Value = i + 1);
                }
            });
            ProgressBar.Visibility = Visibility.Collapsed;
            MessageBox.Show($"Conversion completed. Success: {success}, Failed: {fail}", "Result", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}