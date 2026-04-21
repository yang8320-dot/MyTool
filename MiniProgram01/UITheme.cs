using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

public static class UITheme {
    // iOS 標準配色
    public static readonly Color AppleBlue = Color.FromArgb(0, 122, 255);
    public static readonly Color AppleGreen = Color.FromArgb(52, 199, 89);
    public static readonly Color AppleRed = Color.FromArgb(255, 59, 48);
    public static readonly Color AppleYellow = Color.FromArgb(255, 204, 0);
    public static readonly Color BgGray = Color.FromArgb(242, 242, 247); // iOS 系統背景色
    public static readonly Color CardWhite = Color.White;
    public static readonly Color TextMain = Color.FromArgb(28, 28, 30);
    public static readonly Color TextSub = Color.FromArgb(142, 142, 147);

    // 取得支援 DPI 縮放的字型
    public static Font GetFont(float size, FontStyle style = FontStyle.Regular) {
        return new Font("Microsoft JhengHei UI", size, style, GraphicsUnit.Point);
    }

    // 繪製 iOS 風格的圓角矩形路徑
    public static GraphicsPath CreateRoundedRectanglePath(Rectangle rect, int cornerRadius) {
        GraphicsPath path = new GraphicsPath();
        int diameter = cornerRadius * 2;
        Rectangle arc = new Rectangle(rect.X, rect.Y, diameter, diameter);

        // 左上角
        path.AddArc(arc, 180, 90);
        // 右上角
        arc.X = rect.Right - diameter;
        path.AddArc(arc, 270, 90);
        // 右下角
        arc.Y = rect.Bottom - diameter;
        path.AddArc(arc, 0, 90);
        // 左下角
        arc.X = rect.Left;
        path.AddArc(arc, 90, 90);

        path.CloseFigure();
        return path;
    }

    // 為控制項繪製圓角背景 (供 Panel/Button 的 OnPaint 使用)
    public static void DrawRoundedBackground(Graphics g, Rectangle bounds, int radius, Color bgColor) {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (GraphicsPath path = CreateRoundedRectanglePath(bounds, radius))
        using (SolidBrush brush = new SolidBrush(bgColor)) {
            g.FillPath(brush, path);
        }
    }
}
