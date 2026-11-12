using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
namespace PPTAIStudio.Probe.Native
{
 internal sealed class ToolbarForm:Form
 {
  private readonly FlowLayoutPanel row=new FlowLayoutPanel{AutoSize=true,WrapContents=false,Padding=new Padding(3),Margin=Padding.Empty};
  private readonly ToolTip tips=new ToolTip{ShowAlways=true,InitialDelay=300,ReshowDelay=100,AutoPopDelay=5000};
  private readonly Button find,swap,copy,paste,arrange,close;
  private readonly ContextMenuStrip menu=new ContextMenuStrip();
  public event Action<string>? Command;
  public bool MenuOpen=>menu.Visible;
  public ToolbarForm()
  {
   FormBorderStyle=FormBorderStyle.None;ShowInTaskbar=false;AutoScaleMode=AutoScaleMode.None;BackColor=Color.FromArgb(226,232,241);Padding=new Padding(1);AutoSize=true;AutoSizeMode=AutoSizeMode.GrowAndShrink;TopMost=true;
   find=Add("find","비슷한 개체 찾기");swap=Add("swap","사진 두 장 위치 맞바꾸기");copy=Add("copy","기준 크기 복사");paste=Add("paste","복사한 크기 적용");arrange=Add("arrange","선택 개체 정렬·간격");close=Add("close","이 선택에서 도구막대 닫기");Controls.Add(row);
   foreach(var pair in new[]{("left","왼쪽 정렬"),("center","가로 중앙"),("right","오른쪽 정렬"),("top","위쪽 정렬"),("middle","세로 중앙"),("bottom","아래쪽 정렬"),("horizontal","가로 간격"),("vertical","세로 간격")})
   {var action=pair.Item1;menu.Items.Add(pair.Item2,null,(_,__)=>Command?.Invoke(action)).Name=action;}
   arrange.Click-=OnButton;arrange.Click+=(_,__)=>menu.Show(arrange,new Point(0,arrange.Height));
  }
  protected override bool ShowWithoutActivation=>true;
  protected override CreateParams CreateParams{get{var p=base.CreateParams;p.ExStyle|=0x08000000|0x80;return p;}}
  protected override void WndProc(ref Message m){if(m.Msg==0x21){m.Result=new IntPtr(3);return;}base.WndProc(ref m);}
  private Button Add(string id,string tip)
  {
   var b=new Button{Tag=id,Width=32,Height=32,Margin=new Padding(1),FlatStyle=FlatStyle.Flat,BackColor=Color.White,Image=DrawIcon(id),TabStop=false,AccessibleName=tip};b.FlatAppearance.BorderSize=0;b.FlatAppearance.MouseOverBackColor=Color.FromArgb(226,237,254);tips.SetToolTip(b,tip);b.Click+=OnButton;row.Controls.Add(b);return b;
  }
  private void OnButton(object? sender,EventArgs e){if(sender is Button b&&b.Tag is string id)Command?.Invoke(id);}
  public void Configure(int count,bool twoPhotos,bool hasCopy,bool canCopy)
  {
   find.Visible=count==1;swap.Visible=twoPhotos;copy.Visible=count==1&&canCopy;paste.Visible=hasCopy&&canCopy;arrange.Visible=count>=2;close.Visible=true;
   foreach(ToolStripItem item in menu.Items)item.Enabled=(item.Name!="horizontal"&&item.Name!="vertical")||count>=3;
  }
  public void Place(Rectangle bounds,IntPtr owner)
  {
   if(!Visible)Show(new WindowOwner(owner));
   Win32.SetWindowPos(Handle,new IntPtr(-1),bounds.X,bounds.Y,Width,Height,0x10|0x40);
  }
  public void Notify(string message)=>tips.Show(message,row,new Point(0,Height+3),1800);
  protected override void Dispose(bool disposing){if(disposing){foreach(Control c in row.Controls)if(c is Button b)b.Image?.Dispose();tips.Dispose();menu.Dispose();}base.Dispose(disposing);}
  private sealed class WindowOwner:IWin32Window{public IntPtr Handle{get;}public WindowOwner(IntPtr handle){Handle=handle;}}
  private static Bitmap DrawIcon(string id)
  {
   var image=new Bitmap(24,24);using(var g=Graphics.FromImage(image))using(var pen=new Pen(Color.FromArgb(35,68,112),1.8f)){g.SmoothingMode=SmoothingMode.AntiAlias;
    switch(id){
     case "find":g.DrawEllipse(pen,3,3,11,11);g.DrawLine(pen,13,13,21,21);break;
     case "swap":g.DrawLine(pen,3,7,21,7);g.DrawLine(pen,17,3,21,7);g.DrawLine(pen,17,11,21,7);g.DrawLine(pen,21,17,3,17);g.DrawLine(pen,3,17,7,13);g.DrawLine(pen,3,17,7,21);break;
     case "copy":g.DrawRectangle(pen,3,3,12,11);g.DrawRectangle(pen,9,10,12,11);break;
     case "paste":g.DrawRectangle(pen,4,4,16,16);g.DrawLine(pen,7,12,11,16);g.DrawLine(pen,11,16,18,8);break;
     case "arrange":g.DrawLine(pen,4,2,4,22);g.DrawRectangle(pen,7,4,13,4);g.DrawRectangle(pen,7,11,9,4);g.DrawRectangle(pen,7,18,13,3);break;
     default:g.DrawLine(pen,6,6,18,18);g.DrawLine(pen,18,6,6,18);break;
    }
   }return image;
  }
 }
}
