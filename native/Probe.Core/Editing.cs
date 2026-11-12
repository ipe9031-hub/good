using System;
using System.Collections.Generic;
using System.Linq;
namespace PPTAIStudio.Probe.Core
{
 public sealed class Box
 {
  public double X {get;} public double Y {get;} public double Width {get;} public double Height {get;}
  public Box(double x,double y,double width,double height){X=x;Y=y;Width=width;Height=height;}
  public Box At(double x,double y)=>new Box(x,y,Width,Height);
  public Box Size(double width,double height)=>new Box(X,Y,width,height);
  public bool Near(Box b)=>Math.Abs(X-b.X)<0.01&&Math.Abs(Y-b.Y)<0.01&&Math.Abs(Width-b.Width)<0.01&&Math.Abs(Height-b.Height)<0.01;
  public void Validate(double slideWidth,double slideHeight)
  {
   if(new[]{X,Y,Width,Height,slideWidth,slideHeight}.Any(v=>double.IsNaN(v)||double.IsInfinity(v))||Width<=0||Height<=0||X<0||Y<0||X+Width>slideWidth+0.01||Y+Height>slideHeight+0.01)throw new InvalidOperationException("개체가 슬라이드 밖으로 나가거나 크기가 올바르지 않습니다.");
  }
 }
 public sealed class ShapeState
 {
  public int Id {get;} public int Type {get;} public int? Preset {get;} public int? Fill {get;} public Box Bounds {get;}
  public bool IsPicture=>Type==13||Type==11;
  public ShapeState(int id,int type,Box bounds,int? preset=null,int? fill=null){Id=id;Type=type;Bounds=bounds;Preset=preset;Fill=fill;}
 }
 public sealed class SelectionState
 {
  public string Document {get;} public int Slide {get;} public double Width {get;} public double Height {get;} public IReadOnlyList<ShapeState> Shapes {get;}
  public SelectionState(string document,int slide,double width,double height,IEnumerable<ShapeState> shapes){Document=document;Slide=slide;Width=width;Height=height;Shapes=shapes.ToArray();}
  public string Signature=>Document+"/"+Slide+"/"+string.Join(",",Shapes.Select(x=>x.Id).OrderBy(x=>x));
  public void AssertSame(SelectionState current)
  {
   if(Signature!=current.Signature||Shapes.Any(x=>!x.Bounds.Near(current.Shapes.Single(y=>y.Id==x.Id).Bounds)))throw new InvalidOperationException("선택 대상이나 위치가 바뀌었습니다. 다시 선택해 주세요.");
  }
 }
 public sealed class Change
 {
  public int Id {get;} public Box Before {get;} public Box After {get;}
  public Change(ShapeState shape,Box after){Id=shape.Id;Before=shape.Bounds;After=after;}
 }
 public interface IEditHost
 {
  SelectionState GetSelection();
  void StartUndo();
  void Write(string document,int slide,int id,Box bounds);
 }
 public static class Editing
 {
  public static IReadOnlyList<Change> Swap(SelectionState s)
  {
   if(s.Shapes.Count!=2||s.Shapes.Any(x=>!x.IsPicture))throw new InvalidOperationException("사진 두 장만 선택해 주세요.");
   var a=s.Shapes[0];var b=s.Shapes[1];return Validate(s,new[]{new Change(a,a.Bounds.At(b.Bounds.X,b.Bounds.Y)),new Change(b,b.Bounds.At(a.Bounds.X,a.Bounds.Y))});
  }
  public static IReadOnlyList<Change> CloneSize(SelectionState s,Box source)
  {
   if(s.Shapes.Count==0)throw new InvalidOperationException("크기를 적용할 개체를 선택해 주세요.");
   return Validate(s,s.Shapes.Select(x=>new Change(x,x.Bounds.Size(source.Width,source.Height))));
  }
  public static IReadOnlyList<Change> Arrange(SelectionState s,string action)
  {
   if(s.Shapes.Count<2)throw new InvalidOperationException("개체 두 개 이상을 선택해 주세요.");
   double left=s.Shapes.Min(x=>x.Bounds.X),top=s.Shapes.Min(x=>x.Bounds.Y),right=s.Shapes.Max(x=>x.Bounds.X+x.Bounds.Width),bottom=s.Shapes.Max(x=>x.Bounds.Y+x.Bounds.Height);
   if(action=="horizontal"||action=="vertical")
   {
    if(s.Shapes.Count<3)throw new InvalidOperationException("간격 배분은 개체 세 개 이상이 필요합니다.");
    bool horizontal=action=="horizontal";var sorted=s.Shapes.OrderBy(x=>horizontal?x.Bounds.X:x.Bounds.Y).ToArray();double cursor=horizontal?left:top;
    double gap=((horizontal?right-left:bottom-top)-sorted.Sum(x=>horizontal?x.Bounds.Width:x.Bounds.Height))/(sorted.Length-1);
    if(gap<0)throw new InvalidOperationException("개체가 겹칩니다. 먼저 선택 영역을 넓혀 주세요.");
    var changes=new List<Change>();foreach(var x in sorted){changes.Add(new Change(x,x.Bounds.At(horizontal?cursor:x.Bounds.X,horizontal?x.Bounds.Y:cursor)));cursor+=(horizontal?x.Bounds.Width:x.Bounds.Height)+gap;}return Validate(s,changes);
   }
   return Validate(s,s.Shapes.Select(x=>{var b=x.Bounds;double nx=b.X,ny=b.Y;switch(action){case "left":nx=left;break;case "center":nx=(left+right-b.Width)/2;break;case "right":nx=right-b.Width;break;case "top":ny=top;break;case "middle":ny=(top+bottom-b.Height)/2;break;case "bottom":ny=bottom-b.Height;break;default:throw new InvalidOperationException("지원하지 않는 정렬입니다.");}return new Change(x,b.At(nx,ny));}));
  }
  private static IReadOnlyList<Change> Validate(SelectionState s,IEnumerable<Change> input)
  {
   var changes=input.ToArray();foreach(var c in changes){if(!s.Shapes.Any(x=>x.Id==c.Id))throw new InvalidOperationException("선택 외 개체는 변경할 수 없습니다.");c.After.Validate(s.Width,s.Height);}return changes;
  }
  public static void Execute(IEditHost host,SelectionState expected,IReadOnlyList<Change> changes)
  {
   expected.AssertSame(host.GetSelection());Validate(expected,changes);host.StartUndo();var attempted=new List<Change>();
   try{foreach(var c in changes){attempted.Add(c);host.Write(expected.Document,expected.Slide,c.Id,c.After);}}
   catch(Exception cause)
   {
    bool restored=true;foreach(var c in attempted.AsEnumerable().Reverse()){try{host.Write(expected.Document,expected.Slide,c.Id,c.Before);}catch{restored=false;}}
    throw new InvalidOperationException(restored?"작업에 실패해 이전 위치·크기로 복원을 시도했습니다. 결과를 확인해 주세요.":"작업 일부가 적용됐을 수 있습니다. 복제본을 확인하고 PowerPoint에서 Ctrl+Z를 사용하세요.",cause);
   }
  }
  public static IReadOnlyList<ShapeState> Similar(ShapeState reference,IEnumerable<ShapeState> candidates)
  {
   double dx=Math.Max(1.5,Math.Min(18,reference.Bounds.Width*.05)),dy=Math.Max(1.5,Math.Min(18,reference.Bounds.Height*.05));
   return candidates.Where(x=>(reference.IsPicture?x.IsPicture:x.Type==reference.Type)&&x.Preset==reference.Preset&&Math.Abs(x.Bounds.Width-reference.Bounds.Width)<=dx&&Math.Abs(x.Bounds.Height-reference.Bounds.Height)<=dy&&(reference.IsPicture||(reference.Fill.HasValue?x.Fill.HasValue&&ColorDistance(x.Fill.Value,reference.Fill.Value)<=24:!x.Fill.HasValue))).ToArray();
  }
  private static double ColorDistance(int a,int b)=>Math.Sqrt(new[]{0,8,16}.Sum(shift=>Math.Pow(((a>>shift)&255)-((b>>shift)&255),2)));
 }
 public static class Placement
 {
  public static Box Near(Box selection,double width,double height,Box work)
  {
   double x=selection.X+selection.Width+8,y=selection.Y-4;
   if(x+width>work.X+work.Width)x=selection.X-width-8;
   x=Math.Max(work.X+4,Math.Min(x,work.X+work.Width-width-4));y=Math.Max(work.Y+4,Math.Min(y,work.Y+work.Height-height-4));return new Box(x,y,width,height);
  }
 }
}
