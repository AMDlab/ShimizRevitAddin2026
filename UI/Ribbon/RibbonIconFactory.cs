using System;
using System.Windows;
using System.Windows.Media;

namespace ShimizRevitAddin2026.UI.Ribbon
{
    internal class RibbonIconFactory
    {
        public ImageSource CreateIcon16()
        {
            return CreateIcon(16.0);
        }

        public ImageSource CreateIcon32()
        {
            return CreateIcon(32.0);
        }

        private ImageSource CreateIcon(double size)
        {
            if (size <= 0)
            {
                size = 16.0;
            }

            var group = new DrawingGroup();
            group.Children.Add(BuildBackground(size));
            group.Children.Add(BuildRebarBars(size));
            group.Children.Add(BuildCheckMark(size));
            group.Freeze();

            var image = new DrawingImage(group);
            image.Freeze();
            return image;
        }

        private Drawing BuildBackground(double size)
        {
            var bg = new SolidColorBrush(Color.FromRgb(245, 245, 245));
            bg.Freeze();

            var stroke = new SolidColorBrush(Color.FromRgb(120, 120, 120));
            stroke.Freeze();

            var pen = new Pen(stroke, Math.Max(1.0, size / 20.0))
            {
                LineJoin = PenLineJoin.Round
            };
            pen.Freeze();

            var radius = size * 0.12;
            var rect = new Rect(0, 0, size, size);
            var geometry = new RectangleGeometry(rect, radius, radius);
            geometry.Freeze();

            var drawing = new GeometryDrawing(bg, pen, geometry);
            drawing.Freeze();
            return drawing;
        }

        private Drawing BuildRebarBars(double size)
        {
            var stroke = new SolidColorBrush(Color.FromRgb(0, 120, 215));
            stroke.Freeze();

            var pen = new Pen(stroke, Math.Max(1.2, size / 12.0))
            {
                StartLineCap = PenLineCap.Round,
                EndLineCap = PenLineCap.Round,
                LineJoin = PenLineJoin.Round
            };
            pen.Freeze();

            var gap = size * 0.18;
            var left = size * 0.22;
            var top = size * 0.20;
            var bottom = size * 0.80;

            var bars = new GeometryGroup();
            bars.Children.Add(new LineGeometry(new Point(left + 0 * gap, top), new Point(left + 0 * gap, bottom)));
            bars.Children.Add(new LineGeometry(new Point(left + 1 * gap, top), new Point(left + 1 * gap, bottom)));
            bars.Children.Add(new LineGeometry(new Point(left + 2 * gap, top), new Point(left + 2 * gap, bottom)));
            bars.Freeze();

            var drawing = new GeometryDrawing(null, pen, bars);
            drawing.Freeze();
            return drawing;
        }

        private Drawing BuildCheckMark(double size)
        {
            var stroke = new SolidColorBrush(Color.FromRgb(0, 170, 90));
            stroke.Freeze();

            var pen = new Pen(stroke, Math.Max(1.4, size / 10.0))
            {
                StartLineCap = PenLineCap.Round,
                EndLineCap = PenLineCap.Round,
                LineJoin = PenLineJoin.Round
            };
            pen.Freeze();

            var p1 = new Point(size * 0.38, size * 0.60);
            var p2 = new Point(size * 0.50, size * 0.72);
            var p3 = new Point(size * 0.78, size * 0.40);

            var figure = new PathFigure { StartPoint = p1, IsClosed = false, IsFilled = false };
            figure.Segments.Add(new LineSegment(p2, true));
            figure.Segments.Add(new LineSegment(p3, true));

            var geometry = new PathGeometry();
            geometry.Figures.Add(figure);
            geometry.Freeze();

            var drawing = new GeometryDrawing(null, pen, geometry);
            drawing.Freeze();
            return drawing;
        }
    }
}

