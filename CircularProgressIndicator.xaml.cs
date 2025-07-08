using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace MsgToPdfConverter
{
    public partial class CircularProgressIndicator : UserControl
    {
        public static readonly DependencyProperty ProgressProperty =
            DependencyProperty.Register("Progress", typeof(double), typeof(CircularProgressIndicator),
                new PropertyMetadata(0.0, OnProgressChanged));

        public static readonly DependencyProperty MaximumProperty =
            DependencyProperty.Register("Maximum", typeof(double), typeof(CircularProgressIndicator),
                new PropertyMetadata(100.0, OnProgressChanged));

        public static readonly DependencyProperty DisplayModeProperty =
            DependencyProperty.Register("DisplayMode", typeof(ProgressDisplayMode), typeof(CircularProgressIndicator),
                new PropertyMetadata(ProgressDisplayMode.Percent, OnProgressChanged));

        public double Progress
        {
            get { return (double)GetValue(ProgressProperty); }
            set { SetValue(ProgressProperty, value); }
        }

        public double Maximum
        {
            get { return (double)GetValue(MaximumProperty); }
            set { SetValue(MaximumProperty, value); }
        }

        public ProgressDisplayMode DisplayMode
        {
            get { return (ProgressDisplayMode)GetValue(DisplayModeProperty); }
            set { SetValue(DisplayModeProperty, value); }
        }

        private Path _progressPath;
        private TextBlock _progressText;
        private PathFigure _progressFigure;
        private ArcSegment _progressArc;

        public CircularProgressIndicator()
        {
            InitializeComponent();
            Loaded += CircularProgressIndicator_Loaded;
        }

        private void CircularProgressIndicator_Loaded(object sender, RoutedEventArgs e)
        {
            _progressPath = FindName("ProgressPath") as Path;
            _progressText = FindName("ProgressText") as TextBlock;
            if (_progressPath != null && _progressPath.Data is PathGeometry geometry &&
                geometry.Figures.Count > 0 &&
                geometry.Figures[0].Segments.Count > 0 &&
                geometry.Figures[0].Segments[0] is ArcSegment arc)
            {
                _progressFigure = geometry.Figures[0];
                _progressArc = arc;
            }
            UpdateProgress();
        }

        private static void OnProgressChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = d as CircularProgressIndicator;
            control?.UpdateProgress();
        }

        private void UpdateProgress()
        {
            if (_progressFigure == null || _progressArc == null || _progressText == null || _progressPath == null)
                return;

            double max = Maximum > 0 ? Maximum : 100.0;
            double value = Math.Max(0, Math.Min(Progress, max));
            double progress = value / max;
            double angle = progress * 360;

            if (progress <= 0)
            {
                _progressPath.Visibility = Visibility.Hidden;
                _progressText.Text = DisplayMode == ProgressDisplayMode.Percent ? "0%" : $"0/{(int)max}";
                return;
            }

            _progressPath.Visibility = Visibility.Visible;

            double radius = 7;
            double centerX = 7;
            double centerY = 7;

            // Calculate start point (top of circle)
            double startX = centerX;
            double startY = centerY - radius;

            // Calculate end point
            double radians = (angle - 90) * Math.PI / 180; // -90 to start from top
            double endX = centerX + radius * Math.Cos(radians);
            double endY = centerY + radius * Math.Sin(radians);

            bool isLargeArc = angle > 180;

            _progressFigure.StartPoint = new Point(startX, startY);
            _progressArc.Point = new Point(endX, endY);
            _progressArc.Size = new Size(radius, radius);
            _progressArc.IsLargeArc = isLargeArc;

            // Set progress text
            if (DisplayMode == ProgressDisplayMode.Percent)
            {
                int percent = (int)Math.Round(progress * 100);
                _progressText.Text = $"{percent}%";
            }
            else
            {
                _progressText.Text = $"{(int)value}/{(int)max}";
            }
        }
    }

    public enum ProgressDisplayMode
    {
        Percent,
        ValueOverMax
    }
}
