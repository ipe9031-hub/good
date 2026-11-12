using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using PPTAIStudio.Probe.Core;
namespace PPTAIStudio.Probe.Native
{
 public sealed class HostFrame
 {
  public SelectionState Selection {get;} public Rectangle ScreenSelection {get;} public IntPtr Window {get;}
  public HostFrame(SelectionState selection,Rectangle bounds,IntPtr window){Selection=selection;ScreenSelection=bounds;Window=window;}
 }
 public sealed class PowerPointHost:IEditHost,IDisposable
 {
  private object? application;private readonly bool ownsApplication;private IntPtr presentationIdentity;private string presentationToken=Guid.NewGuid().ToString("N");
  private dynamic App=>application??throw new ObjectDisposedException(nameof(PowerPointHost));
  public PowerPointHost(object application,bool ownsApplication){this.application=application;this.ownsApplication=ownsApplication;}
  public static PowerPointHost Attach()=>new PowerPointHost(Marshal.GetActiveObject("PowerPoint.Application"),true);
  private string DocumentKey(object presentation)
  {
   // Keep exactly one native identity reference so out-of-process COM proxies
   // cannot receive a new address between selection reads.
   IntPtr p=Marshal.GetIUnknownForObject(presentation);
   if(p!=presentationIdentity){if(presentationIdentity!=IntPtr.Zero)Marshal.Release(presentationIdentity);presentationIdentity=p;presentationToken=Guid.NewGuid().ToString("N");}
   else Marshal.Release(p);
   return presentationToken;
  }
  public IntPtr Window=>Win32.GetAncestor(new IntPtr(Convert.ToInt64(App.HWND)),2);
  public bool IsForeground(){try{return Window!=IntPtr.Zero&&Win32.GetAncestor(Win32.GetForegroundWindow(),2)==Window;}catch{return false;}}
  public SelectionState GetSelection()=>ReadFrame().Selection;
  public HostFrame ReadFrame()
  {
   using(var window=new ComLease((object)App.ActiveWindow))
   using(var selection=new ComLease((object)window.Value.Selection))
   using(var presentation=new ComLease((object)App.ActivePresentation))
   using(var page=new ComLease((object)presentation.Value.PageSetup))
   using(var view=new ComLease((object)window.Value.View))
   using(var slide=new ComLease((object)view.Value.Slide))
   {
    var shapes=new List<ShapeState>();
    if(Convert.ToInt32(selection.Value.Type)==2)
     using(var range=new ComLease((object)selection.Value.ShapeRange))
     {int count=Convert.ToInt32(range.Value.Count);if(count>500)throw new InvalidOperationException("검증본은 한 번에 500개까지 지원합니다.");for(int i=1;i<=count;i++)using(var shape=new ComLease((object)range.Value.Item(i)))shapes.Add(ReadShape(shape.Value));}
    var state=new SelectionState(DocumentKey(presentation.Value),(int)slide.Value.SlideID,(double)page.Value.SlideWidth,(double)page.Value.SlideHeight,shapes);
    Rectangle screen=Rectangle.Empty;
    if(shapes.Count>0)
    {
     double left=shapes.Min(x=>x.Bounds.X),top=shapes.Min(x=>x.Bounds.Y),right=shapes.Max(x=>x.Bounds.X+x.Bounds.Width),bottom=shapes.Max(x=>x.Bounds.Y+x.Bounds.Height);
     int x1=Convert.ToInt32(window.Value.PointsToScreenPixelsX((float)left)),y1=Convert.ToInt32(window.Value.PointsToScreenPixelsY((float)top)),x2=Convert.ToInt32(window.Value.PointsToScreenPixelsX((float)right)),y2=Convert.ToInt32(window.Value.PointsToScreenPixelsY((float)bottom));
     screen=Rectangle.FromLTRB(Math.Min(x1,x2),Math.Min(y1,y2),Math.Max(x1,x2),Math.Max(y1,y2));
    }
    return new HostFrame(state,screen,Window);
   }
  }
  private static ShapeState ReadShape(dynamic shape)
  {
   int type=Convert.ToInt32(shape.Type);int? preset=null,fill=null;
   if(type==14)using(var placeholder=new ComLease((object)shape.PlaceholderFormat)){int contained=Convert.ToInt32(placeholder.Value.ContainedType);if(contained==13||contained==11)type=contained;}
   if(type==1)preset=Convert.ToInt32(shape.AutoShapeType);
   if(type==1||type==17)
    using(var format=new ComLease((object)shape.Fill))
     if(Convert.ToInt32(format.Value.Visible)!=0&&Convert.ToInt32(format.Value.Type)==1)
      using(var color=new ComLease((object)format.Value.ForeColor))fill=Convert.ToInt32(color.Value.RGB);
   return new ShapeState(Convert.ToInt32(shape.Id),type,new Box((double)shape.Left,(double)shape.Top,(double)shape.Width,(double)shape.Height),preset,fill);
  }
  private void WithSlide(string document,int slideId,Action<dynamic> action)
  {
   using(var presentation=new ComLease((object)App.ActivePresentation))
   using(var window=new ComLease((object)App.ActiveWindow))
   using(var view=new ComLease((object)window.Value.View))
   using(var slide=new ComLease((object)view.Value.Slide))
   {
    if(DocumentKey(presentation.Value)!=document||Convert.ToInt32(slide.Value.SlideID)!=slideId)throw new InvalidOperationException("작업 문서나 슬라이드가 바뀌었습니다.");action(slide.Value);
   }
  }
  public IReadOnlyList<ShapeState> ReadSlide(SelectionState expected)
  {
   var result=new List<ShapeState>();WithSlide(expected.Document,expected.Slide,slide=>{using(var shapes=new ComLease((object)slide.Shapes)){int count=Convert.ToInt32(shapes.Value.Count);if(count>500)throw new InvalidOperationException("검증본의 슬라이드 검색 한도는 개체 500개입니다.");for(int i=1;i<=count;i++)using(var shape=new ComLease((object)shapes.Value.Item(i)))result.Add(ReadShape(shape.Value));}});return result;
  }
  public void Select(SelectionState expected,IReadOnlyList<ShapeState> found)
  {
   expected.AssertSame(GetSelection());
   WithSlide(expected.Document,expected.Slide,slide=>{
    using(var shapes=new ComLease((object)slide.Shapes))
    {
     var indices=new List<object>();int count=Convert.ToInt32(shapes.Value.Count);for(int i=1;i<=count;i++)using(var shape=new ComLease((object)shapes.Value.Item(i)))if(found.Any(x=>x.Id==Convert.ToInt32(shape.Value.Id)))indices.Add(i);
     if(indices.Count!=found.Count||indices.Count==0)throw new InvalidOperationException("검색 대상이 바뀌었습니다. 다시 선택해 주세요.");
     using(var range=new ComLease((object)shapes.Value.Range(indices.ToArray())))range.Value.Select(-1);
    }
   });
  }
  public void StartUndo()=>App.StartNewUndoEntry();
  public void Write(string document,int slideId,int id,Box bounds)
  {
   WithSlide(document,slideId,slide=>{
    using(var shapes=new ComLease((object)slide.Shapes))
    {
     int count=Convert.ToInt32(shapes.Value.Count);for(int i=1;i<=count;i++)using(var shape=new ComLease((object)shapes.Value.Item(i)))
     {
      if(Convert.ToInt32(shape.Value.Id)!=id)continue;
      bool resize=Math.Abs((double)shape.Value.Width-bounds.Width)>.001||Math.Abs((double)shape.Value.Height-bounds.Height)>.001;
      if(resize){int locked=Convert.ToInt32(shape.Value.LockAspectRatio);try{shape.Value.LockAspectRatio=0;shape.Value.Width=(float)bounds.Width;shape.Value.Height=(float)bounds.Height;}finally{shape.Value.LockAspectRatio=locked;}}
      shape.Value.Left=(float)bounds.X;shape.Value.Top=(float)bounds.Y;return;
     }
     throw new InvalidOperationException("선택한 개체를 찾을 수 없습니다.");
    }
   });
  }
  public void Dispose(){if(presentationIdentity!=IntPtr.Zero){Marshal.Release(presentationIdentity);presentationIdentity=IntPtr.Zero;}var old=application;application=null;if(ownsApplication&&old!=null&&Marshal.IsComObject(old))Marshal.ReleaseComObject(old);}
 }
}
