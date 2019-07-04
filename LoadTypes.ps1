Try {
  [void][Window]
}
Catch {
  Add-Type @"
  using System;
  using System.Runtime.InteropServices;
  public class Window {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("User32.dll")]
    public extern static bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);
  }
  public struct RECT
  {
    public int Left;        // x position of upper-left corner
    public int Top;         // y position of upper-left corner
    public int Right;       // x position of lower-right corner
    public int Bottom;      // y position of lower-right corner
  }
"@
}