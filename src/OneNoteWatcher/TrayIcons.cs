using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OneNoteWatcher;

/// <summary>
/// Tray icons built at runtime from one embedded source image (Assets/cloud.png — a black cloud with
/// a sync arrow, transparent background).
/// TWO STATES ONLY — success = the cloud in its original black, outlined in white so it reads on a
/// dark taskbar; error = the same cloud in red, pulsing (something to look at). A middle "warning"
/// tier is deliberately absent: it is the tier people learn to ignore.
/// </summary>
public static class TrayIcons
{
    public enum State { Ok, Error }

    private const int Size = 32;          // tray icons are consumed at 16px; 32 leaves room to shrink
    private const int OutlineRadius = 2;  // white halo thickness, in canvas pixels

    private static readonly Color Outline = Color.White;
    private static readonly Color ErrorDim = Color.FromArgb(0xC0, 0x22, 0x1A);
    private static readonly Color ErrorBright = Color.FromArgb(0xFF, 0x3B, 0x30);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyIcon(IntPtr handle);

    // There are only three icons in the whole app, and the error state repaints twice a second, so
    // they are built once and shared. Callers must NOT dispose them.
    private static readonly Dictionary<(State, bool), Icon> Cache = new();
    private static readonly object Gate = new();

    // The source art, loaded once. Everything else is a recolour of its alpha channel.
    private static Bitmap? _source;
    // The artwork's opaque bounds inside that frame. The PNG is a 56x44 cloud centred in a 64x64
    // canvas, so scaling the whole frame down left the tray icon visibly smaller than its neighbours.
    private static Rectangle _sourceBounds;

    /// <summary>The icon for this state. Cached and shared — do not dispose the result.</summary>
    public static Icon Make(State state, bool pulseBright = false)
    {
        lock (Gate)
        {
            if (Cache.TryGetValue((state, pulseBright), out var cached)) return cached;
            var icon = Render(state, pulseBright);
            Cache[(state, pulseBright)] = icon;
            return icon;
        }
    }

    private static Bitmap Source
    {
        get
        {
            if (_source != null) return _source;
            var asm = Assembly.GetExecutingAssembly();
            using var s = asm.GetManifestResourceStream("OneNoteWatcher.Assets.cloud.png")
                ?? throw new InvalidOperationException("Embedded resource OneNoteWatcher.Assets.cloud.png is missing.");
            using var raw = new Bitmap(s);
            // Copy out of the stream-backed bitmap: Bitmap keeps the stream alive and we are closing it.
            _source = new Bitmap(raw);
            _sourceBounds = OpaqueBounds(_source);
            return _source;
        }
    }

    private static Icon Render(State state, bool pulseBright)
    {
        var body = state == State.Error
            ? (pulseBright ? ErrorBright : ErrorDim)
            : Color.Black;                                    // success = the original artwork

        // The halo dims with the pulse too, otherwise the white ring makes the dark phase look bright.
        var halo = state == State.Error && !pulseBright
            ? Color.FromArgb(0xB0, Outline)
            : Outline;

        using var bmp = new Bitmap(Size, Size, PixelFormat.Format32bppArgb);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.Clear(Color.Transparent);

            // Only the artwork itself is drawn — the source frame's padding is cropped away, then the
            // cloud is scaled to fill the canvas apart from the margin the halo needs. Without this the
            // icon sat noticeably smaller than every other icon in the tray.
            var art = Source;                 // touching Source also populates _sourceBounds
            var src = _sourceBounds;
            var dest = Fit(src.Size, OutlineRadius);

            // The outline is the silhouette stamped around a circle and then the body drawn over it,
            // i.e. a dilation of the alpha channel.
            using (var haloArt = Tint(art, halo))
            {
                for (var i = 0; i < 16; i++)
                {
                    var a = i * Math.PI / 8;
                    var dx = (float)(Math.Cos(a) * OutlineRadius);
                    var dy = (float)(Math.Sin(a) * OutlineRadius);
                    g.DrawImage(haloArt,
                        new RectangleF(dest.X + dx, dest.Y + dy, dest.Width, dest.Height),
                        src, GraphicsUnit.Pixel);
                }
            }
            using (var bodyArt = Tint(art, body))
                g.DrawImage(bodyArt, dest, src, GraphicsUnit.Pixel);
        }

        // Bitmap.GetHicon returns a native handle that Icon.FromHandle does NOT take ownership of, so
        // disposing that Icon leaks the HICON. Left unfixed the pulsing error icon exhausted the
        // process GDI quota in about an hour and the tray died with "A generic error occurred in GDI+"
        // — it crashed precisely while it was reporting a problem. Clone into a managed icon that owns
        // its own handle, then destroy the native one.
        var h = bmp.GetHicon();
        try
        {
            using var borrowed = Icon.FromHandle(h);
            return (Icon)borrowed.Clone();
        }
        finally { DestroyIcon(h); }
    }

    /// <summary>
    /// The largest rectangle of the given aspect ratio that fits the canvas inside <paramref name="margin"/>,
    /// centred. Aspect is preserved: the cloud is wider than it is tall, so it fills the width and is
    /// centred vertically rather than being stretched square.
    /// </summary>
    private static RectangleF Fit(Size art, int margin)
    {
        float avail = Size - margin * 2;
        var scale = Math.Min(avail / art.Width, avail / art.Height);
        var w = art.Width * scale;
        var h = art.Height * scale;
        return new RectangleF((Size - w) / 2f, (Size - h) / 2f, w, h);
    }

    /// <summary>The bounding box of everything not effectively transparent, i.e. the artwork itself.</summary>
    private static Rectangle OpaqueBounds(Bitmap bmp)
    {
        // Runs once, over a 64x64 image — GetPixel is more than fast enough and keeps this unsafe-free.
        int minX = bmp.Width, minY = bmp.Height, maxX = -1, maxY = -1;
        for (var y = 0; y < bmp.Height; y++)
        for (var x = 0; x < bmp.Width; x++)
        {
            if (bmp.GetPixel(x, y).A <= 8) continue;   // ignore near-transparent antialiasing
            if (x < minX) minX = x;
            if (x > maxX) maxX = x;
            if (y < minY) minY = y;
            if (y > maxY) maxY = y;
        }

        // A fully transparent image would leave the box inverted; fall back to the whole frame.
        if (maxX < minX || maxY < minY) return new Rectangle(0, 0, bmp.Width, bmp.Height);
        return Rectangle.FromLTRB(minX, minY, maxX + 1, maxY + 1);
    }

    /// <summary>
    /// The source image's shape in a flat colour: alpha is kept, RGB is replaced. Done with a colour
    /// matrix so the anti-aliased edges of the artwork stay smooth.
    /// </summary>
    private static Bitmap Tint(Bitmap src, Color color)
    {
        var outBmp = new Bitmap(src.Width, src.Height, PixelFormat.Format32bppArgb);
        var matrix = new ColorMatrix(new[]
        {
            new[] { 0f, 0f, 0f, 0f, 0f },
            new[] { 0f, 0f, 0f, 0f, 0f },
            new[] { 0f, 0f, 0f, 0f, 0f },
            new[] { 0f, 0f, 0f, color.A / 255f, 0f },
            new[] { color.R / 255f, color.G / 255f, color.B / 255f, 0f, 1f },
        });
        using var attrs = new ImageAttributes();
        attrs.SetColorMatrix(matrix);
        using var g = Graphics.FromImage(outBmp);
        g.Clear(Color.Transparent);
        g.DrawImage(src, new Rectangle(0, 0, src.Width, src.Height),
            0, 0, src.Width, src.Height, GraphicsUnit.Pixel, attrs);
        return outBmp;
    }
}
