using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
namespace PPTAIStudio.Probe.Native
{
 internal sealed class ComLease:IDisposable
 {
  private object? value;
  public dynamic Value=>value??throw new ObjectDisposedException(nameof(ComLease));
  public ComLease(object value){this.value=value;}
  public void Dispose(){var old=value;value=null;if(old!=null&&Marshal.IsComObject(old))Marshal.ReleaseComObject(old);}
 }
 public static class ProbeLog
 {
  public static string Folder=>Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"PPTAIStudio","NativeProbe","logs");
  public static void Write(string message){try{Directory.CreateDirectory(Folder);var path=Path.Combine(Folder,"probe.log");if(File.Exists(path)&&new FileInfo(path).Length>1024*1024)File.Move(path,Path.Combine(Folder,"probe-"+DateTime.UtcNow.Ticks+".log"));File.AppendAllText(path,DateTime.UtcNow.ToString("o")+" "+message+Environment.NewLine);}catch{/* Logging must not interrupt PowerPoint. */}}
  public static void Error(string operation,Exception e)=>Write(operation+" error="+e.GetType().Name+" hresult=0x"+e.HResult.ToString("X8"));
 }
 internal static class Win32
 {
  [DllImport("user32.dll")] internal static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] internal static extern IntPtr GetAncestor(IntPtr hwnd,uint flags);
  [DllImport("user32.dll")] internal static extern bool GetWindowRect(IntPtr hwnd,out Rect rect);
  [DllImport("user32.dll")] internal static extern short GetAsyncKeyState(int key);
  [DllImport("user32.dll")] internal static extern bool SetForegroundWindow(IntPtr hwnd);
  [DllImport("user32.dll")] internal static extern bool SetWindowPos(IntPtr hwnd,IntPtr after,int x,int y,int width,int height,uint flags);
  [StructLayout(LayoutKind.Sequential)] internal struct Rect{public int Left,Top,Right,Bottom;public Rectangle Rectangle=>Rectangle.FromLTRB(Left,Top,Right,Bottom);}
 }
}
