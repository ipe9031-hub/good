using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.PortableExecutable;
using PPTAIStudio.Probe.Core;
int passed=0;
void Check(string name,Action test){try{test();Console.WriteLine("PASS "+name);passed++;}catch(Exception e){Console.Error.WriteLine("FAIL "+name+": "+e.Message);Environment.Exit(1);}}
void Assert(bool value,string message="Assertion failed"){if(!value)throw new Exception(message);}
void Throws(Action action){try{action();}catch(InvalidOperationException){return;}throw new Exception("Expected operation rejection");}
ShapeState Shape(int id,int type,double x,double y,double w,double h,int? preset=null,int? fill=null)=>new ShapeState(id,type,new Box(x,y,w,h),preset,fill);
SelectionState Selected(params ShapeState[] shapes)=>new SelectionState("doc",256,960,540,shapes);
var a=Shape(1,13,30,100,200,100);var b=Shape(2,13,500,150,300,160);
Check("two photographs swap position and preserve size",()=>{var result=Editing.Swap(Selected(a,b));Assert(result[0].After.X==500&&result[0].After.Width==200);Assert(result[1].After.X==30&&result[1].After.Width==300);});
Check("non-photo and invalid counts are rejected",()=>{Throws(()=>Editing.Swap(Selected(a)));Throws(()=>Editing.Swap(Selected(a,Shape(3,1,300,100,100,100,1))));});
Check("out-of-slide swap is rejected before writes",()=>{Throws(()=>Editing.Swap(Selected(Shape(1,13,0,0,400,100),Shape(2,13,850,0,50,50))));});
Check("size copy preserves target position and original source",()=>{var target=Shape(4,1,40,40,50,30,1);var result=Editing.CloneSize(Selected(target),a.Bounds);Assert(result[0].After.X==40&&result[0].After.Y==40&&result[0].After.Width==200&&target.Bounds.Width==50);});
Check("exact native preset distinguishes arrow and rectangle",()=>{var reference=Shape(1,1,0,0,100,80,33,255);var same=Shape(2,1,200,0,104,81,33,250);var rect=Shape(3,1,400,0,100,80,1,255);Assert(Editing.Similar(reference,new[]{reference,same,rect}).Select(x=>x.Id).SequenceEqual(new[]{1,2}));});
Check("linked and embedded pictures match by size",()=>{var linked=Shape(3,11,0,0,205,103);Assert(Editing.Similar(a,new[]{a,linked,b}).Count==2);});
Check("alignment and spacing use the selection bounds",()=>{var s=Selected(Shape(1,1,10,20,100,50),Shape(2,1,200,80,150,50),Shape(3,1,700,100,100,50));var left=Editing.Arrange(s,"left");Assert(left.All(x=>x.After.X==10));var spaced=Editing.Arrange(s,"horizontal");Assert(Math.Abs((spaced[1].After.X-spaced[0].After.X-100)-(spaced[2].After.X-spaced[1].After.X-150))<.001);Throws(()=>Editing.Arrange(Selected(a,b),"horizontal"));});
Check("mutation rejects stale selection and document",()=>{var host=new FakeHost(Selected(a,b));host.Current=new SelectionState("other",256,960,540,new[]{a,b});Throws(()=>Editing.Execute(host,Selected(a,b),Editing.Swap(Selected(a,b))));Assert(host.Writes==0);});
Check("mutation rejects moved objects",()=>{var host=new FakeHost(Selected(a,b));host.Current=Selected(Shape(1,13,31,100,200,100),b);Throws(()=>Editing.Execute(host,Selected(a,b),Editing.Swap(Selected(a,b))));Assert(host.Writes==0);});
Check("host failure attempts rollback without touching other objects",()=>{var s=Selected(a,b);var host=new FakeHost(s){FailOnWrite=2};Throws(()=>Editing.Execute(host,s,Editing.Swap(s)));Assert(host.Values[1].Near(a.Bounds)&&host.Values[2].Near(b.Bounds));Assert(host.Undos==1);});
Check("toolbar stays within monitor work area",()=>{var work=new Box(-1920,0,1920,1080);var pos=Placement.Near(new Box(-20,1050,10,20),150,42,work);Assert(pos.X>=work.X&&pos.X+150<=0&&pos.Y+42<=1080);});
Check("toolbar can sit to the right of a shape",()=>{var pos=Placement.Near(new Box(100,100,300,200),140,40,new Box(0,0,1920,1080));Assert(pos.X==408&&pos.Y==96);});
if(args.Length==1)
{
 Check("Windows binaries are managed AnyCPU and the COM library is strong-name signed",()=>{
  foreach(var name in new[]{"PPTAIStudio.Probe.Addin.dll","PPTAIStudio.Probe.Core.dll","PPTAIStudio-Probe.exe"})
   using(var stream=File.OpenRead(Path.Combine(args[0],name)))using(var pe=new PEReader(stream))
   {var header=pe.PEHeaders.CorHeader??throw new Exception("Missing CLR header");Assert((header.Flags&CorFlags.ILOnly)!=0);Assert((header.Flags&CorFlags.Requires32Bit)==0);if(name.EndsWith(".dll",StringComparison.Ordinal))Assert((header.Flags&CorFlags.StrongNameSigned)!=0&&header.StrongNameSignatureDirectory.Size>0);}
 });
}
Console.WriteLine($"{passed} checks passed. Native Windows/PowerPoint UI checks are not run on this OS.");
sealed class FakeHost:IEditHost
{
 public SelectionState Current;public readonly Dictionary<int,Box> Values;public int Writes,Undos,FailOnWrite;
 public FakeHost(SelectionState state){Current=state;Values=state.Shapes.ToDictionary(x=>x.Id,x=>x.Bounds);}
 public SelectionState GetSelection()=>Current;
 public void StartUndo()=>Undos++;
 public void Write(string document,int slide,int id,Box bounds){Writes++;Values[id]=bounds;if(Writes==FailOnWrite)throw new InvalidOperationException("Simulated COM partial write");}
}
