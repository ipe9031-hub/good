using System;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using PPTAIStudio.Probe.Core;
namespace PPTAIStudio.Probe.Native
{
 public sealed class ProbeController:IDisposable
 {
  private readonly PowerPointHost host;private readonly System.Windows.Forms.Timer timer=new System.Windows.Forms.Timer{Interval=200};
  private readonly ToolbarForm toolbar=new ToolbarForm();private HostFrame? frame;private Box? copiedSize;private bool busy,disposed;private string? suppressed;private int? lastError;
  public event Action<string>? Status;
  public ProbeController(PowerPointHost host){if(Thread.CurrentThread.GetApartmentState()!=ApartmentState.STA)throw new InvalidOperationException("PowerPoint 연동은 STA 스레드에서 실행해야 합니다.");this.host=host;timer.Tick+=(_,__)=>Refresh();toolbar.Command+=Run;}
  public void Start(){timer.Start();Status?.Invoke("PowerPoint에서 개체를 선택해 주세요.");}
  private void Refresh()
  {
   if(disposed||busy||toolbar.MenuOpen)return;
   try
   {
    if(!host.IsForeground()){toolbar.Hide();return;}
    if(Win32.GetAsyncKeyState(1)<0&&!toolbar.Bounds.Contains(Cursor.Position)){toolbar.Hide();return;}
    var current=host.ReadFrame();frame=current;lastError=null;
    if(current.Selection.Shapes.Count==0||current.Selection.Signature==suppressed){toolbar.Hide();return;}
    if(current.Selection.Signature!=suppressed)suppressed=null;
    toolbar.Configure(current.Selection.Shapes.Count,current.Selection.Shapes.Count==2&&current.Selection.Shapes.All(x=>x.IsPicture),copiedSize!=null,current.Selection.Shapes.All(x=>x.Bounds.Width>0&&x.Bounds.Height>0));
    var work=Screen.FromRectangle(current.ScreenSelection).WorkingArea;
    if(Win32.GetWindowRect(current.Window,out var native))work=Rectangle.Intersect(work,native.Rectangle);
    if(work.Width<toolbar.Width+8||work.Height<toolbar.Height+8){toolbar.Hide();return;}
    var p=Placement.Near(new Box(current.ScreenSelection.X,current.ScreenSelection.Y,current.ScreenSelection.Width,current.ScreenSelection.Height),toolbar.PreferredSize.Width,toolbar.PreferredSize.Height,new Box(work.X,work.Y,work.Width,work.Height));
    toolbar.Place(new Rectangle((int)p.X,(int)p.Y,toolbar.Width,toolbar.Height),current.Window);Status?.Invoke("연결됨 · "+current.Selection.Shapes.Count+"개 선택");
   }
   catch(Exception e){toolbar.Hide();if(lastError!=e.HResult){lastError=e.HResult;ProbeLog.Error("selection-read",e);Status?.Invoke("PowerPoint 편집 보기에서 다시 선택해 주세요. (0x"+e.HResult.ToString("X8")+")");}}
  }
  private void Run(string action)
  {
   if(busy||frame==null)return;busy=true;
   try
   {
    var expected=frame.Selection;expected.AssertSame(host.GetSelection());
    if(action=="close"){suppressed=expected.Signature;toolbar.Hide();return;}
    if(action=="copy"){if(expected.Shapes.Count!=1)throw new InvalidOperationException("기준 개체 하나를 선택해 주세요.");copiedSize=expected.Shapes[0].Bounds;toolbar.Notify("크기를 복사했습니다.");}
    else if(action=="find")
    {
     if(expected.Shapes.Count!=1)throw new InvalidOperationException("기준 개체 하나를 선택해 주세요.");var found=Editing.Similar(expected.Shapes[0],host.ReadSlide(expected));host.Select(expected,found);Status?.Invoke(found.Count+"개 선택됨");
    }
    else
    {
     var changes=action=="swap"?Editing.Swap(expected):action=="paste"?Editing.CloneSize(expected,copiedSize??throw new InvalidOperationException("먼저 크기를 복사해 주세요.")):Editing.Arrange(expected,action);
     Editing.Execute(host,expected,changes);ProbeLog.Write("edit-"+action+" count="+changes.Count);Status?.Invoke("작업 완료 · "+changes.Count+"개");
    }
   }
   catch(Exception e){ProbeLog.Error("command-"+action,e);MessageBox.Show(e is InvalidOperationException?e.Message:"PowerPoint가 작업을 완료하지 못했습니다. 잠금·편집 권한을 확인해 주세요. 오류가 나면 복제본에서 Ctrl+Z로 결과를 확인하세요.","PPT AI Studio 검증",MessageBoxButtons.OK,MessageBoxIcon.Information);}
   finally{busy=false;}
   Refresh();
  }
  public void Dispose(){if(disposed)return;disposed=true;timer.Stop();timer.Dispose();toolbar.Dispose();host.Dispose();}
 }
}
