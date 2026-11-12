using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
[assembly:ComVisible(false)]
namespace PPTAIStudio.Probe.Native
{
 // Standard Office IDTExtensibility2 ABI. See native/README.md for references.
 [ComImport,ComVisible(true),Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"),InterfaceType(ComInterfaceType.InterfaceIsDual),TypeLibType(4160)]
 public interface IDTExtensibility2
 {
  [DispId(1),MethodImpl(MethodImplOptions.InternalCall)] void OnConnection([In,MarshalAs(UnmanagedType.IDispatch)] object application,int connectMode,[In,MarshalAs(UnmanagedType.IDispatch)] object addIn,[In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
  [DispId(2),MethodImpl(MethodImplOptions.InternalCall)] void OnDisconnection(int removeMode,[In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
  [DispId(3),MethodImpl(MethodImplOptions.InternalCall)] void OnAddInsUpdate([In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
  [DispId(4),MethodImpl(MethodImplOptions.InternalCall)] void OnStartupComplete([In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
  [DispId(5),MethodImpl(MethodImplOptions.InternalCall)] void OnBeginShutdown([In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
 }
 [ComVisible(true),Guid(ClassId),ProgId(ProgramId),ClassInterface(ClassInterfaceType.None),ComDefaultInterface(typeof(IDTExtensibility2))]
 public sealed class Connect:IDTExtensibility2
 {
  public const string ClassId="813CFA48-5BE7-4CBB-8B4C-13F167A64509";
  public const string ProgramId="PPTAIStudio.NativeProbe";
  private ProbeController? controller;
  public void OnConnection(object application,int connectMode,object addIn,ref Array custom)
  {
   try{System.Windows.Forms.Application.EnableVisualStyles();controller?.Dispose();controller=new ProbeController(new PowerPointHost(application,false));controller.Start();ProbeLog.Write("addin-connected");}
   catch(Exception e){ProbeLog.Error("addin-connect",e);}
  }
  public void OnDisconnection(int mode,ref Array custom){controller?.Dispose();controller=null;ProbeLog.Write("addin-disconnected");}
  public void OnAddInsUpdate(ref Array custom){}
  public void OnStartupComplete(ref Array custom){}
  public void OnBeginShutdown(ref Array custom){controller?.Dispose();controller=null;}
 }
}
