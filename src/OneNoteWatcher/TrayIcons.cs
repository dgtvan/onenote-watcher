using System.Drawing;
using System.Drawing.Drawing2D;

namespace OneNoteWatcher;

/// <summary>
/// Icons drawn at runtime (no binary assets): a OneNote-purple rounded square with a status dot.
/// TWO STATES ONLY — green = success, red pulsing = error (something to look at). A middle "warning"
/// tier is deliberately absent: it is the tier people learn to ignore.
/// </summary>
public static class TrayIcons
{
    public enum State { Ok, Error }

    private static readonly Color Purple = Color.FromArgb(0x80, 0x39, 0x7B); // OneNote purple

    public static Icon Make(State state, bool pulseBright = false)
    {
        using var bmp = new Bitmap(32, 32);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);

            using (var body = new SolidBrush(Purple))
            using (var path = Rounded(new Rectangle(3, 3, 26, 26), 6))
                g.FillPath(body, path);

            // status dot
            Color dot = state == State.Error
                ? (pulseBright ? Color.FromArgb(0xFF, 0x3B, 0x30) : Color.FromArgb(0x99, 0x1B, 0x14))
                : Color.FromArgb(0x3C, 0xB3, 0x71);
            var r = state == State.Error && pulseBright ? 11 : 9;
            var cx = 20; var cy = 20;
            if (state == State.Error && pulseBright)
            {
                using var glow = new SolidBrush(Color.FromArgb(80, dot));
                g.FillEllipse(glow, cx - 14, cy - 14, 28, 28);
            }
            using (var db = new SolidBrush(dot))
                g.FillEllipse(db, cx - r / 2, cy - r / 2, r, r);
        }
        return Icon.FromHandle(bmp.GetHicon());
    }

    private static GraphicsPath Rounded(Rectangle r, int radius)
    {
        var p = new GraphicsPath();
        int d = radius * 2;
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }
}
