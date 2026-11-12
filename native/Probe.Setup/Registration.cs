using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using Microsoft.Win32;
using PPTAIStudio.Probe.Native;
namespace PPTAIStudio.Probe.Setup
{
 internal static class Registration
 {
  private static string InstallDir=>Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"PPTAIStudio","NativeProbe","app");
  private static string DllPath=>Path.Combine(InstallDir,"PPTAIStudio.Probe.Addin.dll");
  private static string Clsid=>"{"+Connect.ClassId+"}";
  private static string Classes=>"Software\\Classes\\";
  private static string Addin=>"Software\\Microsoft\\Office\\PowerPoint\\Addins\\"+Connect.ProgramId;
  private static string UninstallKey=>"Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\"+Connect.ProgramId;
  private static RegistryView[] Views=>Environment.Is64BitOperatingSystem?new[]{RegistryView.Registry32,RegistryView.Registry64}:new[]{RegistryView.Registry32};
  private static string[] Keys=>new[]{Classes+"CLSID\\"+Clsid,Classes+Connect.ProgramId,Addin,UninstallKey};
  private static void RequireOfficeClosed(){if(Process.GetProcessesByName("POWERPNT").Length>0)throw new InvalidOperationException("연결 시험을 중지하고 PowerPoint를 완전히 종료한 뒤 등록/해제해 주세요. 파일을 강제로 닫지는 않습니다.");}
  public static void Install()
  {
   RequireOfficeClosed();
   foreach(var view in Views)using(var root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))EnsureOwned(root);
   Directory.CreateDirectory(InstallDir);string source=AppDomain.CurrentDomain.BaseDirectory;
   foreach(var view in Views)using(var root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))using(var existing=root.OpenSubKey(Addin))
    if(existing!=null&&!string.Equals(Path.GetFullPath(source).TrimEnd(Path.DirectorySeparatorChar),Path.GetFullPath(InstallDir).TrimEnd(Path.DirectorySeparatorChar),StringComparison.OrdinalIgnoreCase))
     foreach(var name in new[]{"PPTAIStudio.Probe.Addin.dll","PPTAIStudio.Probe.Core.dll","PPTAIStudio-Probe.exe"})
      if(!SameFile(Path.Combine(source,name),Path.Combine(InstallDir,name)))throw new InvalidOperationException("다른 빌드가 이미 등록되어 있습니다. 기존 등록을 먼저 해제한 뒤 새 검증본을 등록해 주세요.");
   foreach(var name in new[]{"PPTAIStudio.Probe.Addin.dll","PPTAIStudio.Probe.Core.dll","PPTAIStudio-Probe.exe","PPTAIStudio-Probe.exe.config"})
   {
    string from=Path.Combine(source,name),to=Path.Combine(InstallDir,name);
    if(!File.Exists(from)){if(name.EndsWith(".config",StringComparison.Ordinal))continue;throw new FileNotFoundException("검증 ZIP의 파일을 모두 같은 폴더에 풀어 주세요.",name);}
    if(!string.Equals(Path.GetFullPath(from),Path.GetFullPath(to),StringComparison.OrdinalIgnoreCase))File.Copy(from,to,true);
   }
   var assembly=AssemblyName.GetAssemblyName(DllPath);if(assembly.GetPublicKeyToken()?.Length!=8)throw new InvalidOperationException("추가 기능 DLL의 개발 강력한 이름 서명을 확인할 수 없습니다.");
   var backups=new List<Backup>();
   try
   {
    // Capture both views before writing either, including any existing owned settings.
    foreach(var view in Views)using(var root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))foreach(var path in Keys)backups.Add(new Backup(view,path,Capture(root,path)));
    foreach(var view in Views)using(var root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
    {
     Put(root,Classes+Connect.ProgramId,"", "PPT AI Studio Native Probe");Put(root,Classes+Connect.ProgramId+"\\CLSID","",Clsid);
     Put(root,Classes+"CLSID\\"+Clsid,"","PPT AI Studio Native Probe");Put(root,Classes+"CLSID\\"+Clsid+"\\ProgId","",Connect.ProgramId);
     string server=Classes+"CLSID\\"+Clsid+"\\InprocServer32";
     foreach(var key in new[]{server,server+"\\"+assembly.Version})
     {
      Put(root,key,"","mscoree.dll");Put(root,key,"ThreadingModel","Both");Put(root,key,"Class",typeof(Connect).FullName!);Put(root,key,"Assembly",assembly.FullName!);Put(root,key,"RuntimeVersion","v4.0.30319");Put(root,key,"CodeBase",new Uri(DllPath).AbsoluteUri);
     }
     Put(root,Addin,"FriendlyName","PPT AI Studio 설치형 검증");Put(root,Addin,"Description","개체 근처 도구막대 가능성 검증");Put(root,Addin,"LoadBehavior",3);Put(root,Addin,"CommandLineSafe",0);
     Put(root,UninstallKey,"DisplayName","PPT AI Studio 설치형 검증");Put(root,UninstallKey,"DisplayVersion","0.1.0");Put(root,UninstallKey,"InstallLocation",InstallDir);Put(root,UninstallKey,"UninstallString","\""+Path.Combine(InstallDir,"PPTAIStudio-Probe.exe")+"\" --uninstall");Put(root,UninstallKey,"NoModify",1);Put(root,UninstallKey,"NoRepair",1);
    }
   }
   catch{foreach(var backup in backups.AsEnumerable().Reverse()){try{using(var root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,backup.View)){root.DeleteSubKeyTree(backup.Path,false);if(backup.State!=null)Restore(root,backup.Path,backup.State);}}catch{/* Original error is reported; registry restoration can be policy-blocked. */}}throw;}
   ProbeLog.Write("registered-current-user");
  }
  public static void Uninstall()
  {
   RequireOfficeClosed();foreach(var view in Views)using(var root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))EnsureOwned(root);
   foreach(var view in Views)using(var root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))foreach(var path in Keys)root.DeleteSubKeyTree(path,false);
   ProbeLog.Write("unregistered-current-user");
  }
  private static void EnsureOwned(RegistryKey root)
  {
   using(var key=root.OpenSubKey(Classes+"CLSID\\"+Clsid+"\\InprocServer32"))
   {
    if(key==null)return;string? code=key.GetValue("CodeBase") as string;string? assembly=key.GetValue("Assembly") as string;
    if(code!=new Uri(DllPath).AbsoluteUri||assembly==null||!assembly.StartsWith("PPTAIStudio.Probe.Addin,",StringComparison.Ordinal))throw new InvalidOperationException("다른 경로의 등록이 같은 식별자를 사용합니다. 기존 등록을 보존하기 위해 작업을 중단했습니다.");
   }
  }
  private static bool SameFile(string a,string b){if(!File.Exists(a)||!File.Exists(b))return false;using(var sha=SHA256.Create())using(var x=File.OpenRead(a))using(var y=File.OpenRead(b))return sha.ComputeHash(x).SequenceEqual(sha.ComputeHash(y));}
  private static void Put(RegistryKey root,string path,string name,object value){using(var key=root.CreateSubKey(path,true)){if(key==null)throw new UnauthorizedAccessException("현재 사용자 등록 권한이 없습니다.");key.SetValue(name,value,value is int?RegistryValueKind.DWord:RegistryValueKind.String);}}
  private sealed class State
  {
   public readonly Dictionary<string,Tuple<object,RegistryValueKind>> Values=new Dictionary<string,Tuple<object,RegistryValueKind>>();public readonly Dictionary<string,State> Children=new Dictionary<string,State>();
  }
  private sealed class Backup{public RegistryView View{get;}public string Path{get;}public State? State{get;}public Backup(RegistryView view,string path,State? state){View=view;Path=path;State=state;}}
  private static State? Capture(RegistryKey root,string path)
  {
   using(var key=root.OpenSubKey(path)){if(key==null)return null;var state=new State();foreach(var name in key.GetValueNames()){var value=key.GetValue(name,null,RegistryValueOptions.DoNotExpandEnvironmentNames);if(value!=null)state.Values[name]=Tuple.Create(value,key.GetValueKind(name));}foreach(var name in key.GetSubKeyNames()){var child=Capture(key,name);if(child!=null)state.Children[name]=child;}return state;}
  }
  private static void Restore(RegistryKey root,string path,State state){using(var key=root.CreateSubKey(path,true)){if(key==null)throw new UnauthorizedAccessException();foreach(var pair in state.Values)key.SetValue(pair.Key,pair.Value.Item1,pair.Value.Item2);foreach(var child in state.Children)Restore(key,child.Key,child.Value);}}
 }
}
