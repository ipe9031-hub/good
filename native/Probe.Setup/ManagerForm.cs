using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using PPTAIStudio.Probe.Native;
namespace PPTAIStudio.Probe.Setup
{
 internal sealed class ManagerForm:Form
 {
  private ProbeController? controller;private readonly Label status=new Label{AutoSize=false,Height=54,Dock=DockStyle.Bottom,Padding=new Padding(12),Text="PowerPoint를 열고 복제한 시험용 파일을 준비하세요."};
  public ManagerForm()
  {
   Text="PPT AI Studio · Windows 설치형 검증";Width=520;Height=380;MinimumSize=new Size(480,360);StartPosition=FormStartPosition.CenterScreen;Font=new Font("맑은 고딕",9f);
   var panel=new FlowLayoutPanel{Dock=DockStyle.Fill,FlowDirection=FlowDirection.TopDown,WrapContents=false,Padding=new Padding(18),AutoScroll=true};
   panel.Controls.Add(new Label{Text="설치형 가능성 검증 · 0.1.0",AutoSize=true,Font=new Font(Font,FontStyle.Bold)});
   panel.Controls.Add(new Label{Text="미서명 개발 검증본입니다. Node.js·웹 서버·루트 인증서가 필요 없습니다.\n먼저 '연결 시험'으로 도구막대를 확인하세요.",AutoSize=true,MaximumSize=new Size(450,0),Margin=new Padding(0,10,0,12)});
   Add(panel,"① 등록 없이 PowerPoint 연결 시험",()=>{controller?.Dispose();controller=new ProbeController(PowerPointHost.Attach());controller.Status+=message=>status.Text=message;controller.Start();status.Text="연결했습니다. PowerPoint로 돌아가 개체를 선택하세요.";});
   Add(panel,"연결 시험 중지",()=>{controller?.Dispose();controller=null;status.Text="연결 시험을 중지했습니다. PowerPoint는 그대로 유지됩니다.";});
   Add(panel,"② PowerPoint 시작 시 로드 등록",()=>{Registration.Install();status.Text="등록했습니다. PowerPoint를 새로 열고 개체를 선택하세요.";});
   Add(panel,"추가 기능 등록 해제",()=>{Registration.Uninstall();status.Text="등록을 해제했습니다. 기존 웹 애드인은 변경하지 않았습니다.";});
   Add(panel,"로컬 진단 로그 폴더",()=>{System.IO.Directory.CreateDirectory(ProbeLog.Folder);Process.Start("explorer.exe",ProbeLog.Folder);});
   Controls.Add(panel);Controls.Add(status);FormClosed+=(_,__)=>controller?.Dispose();
  }
  private void Add(FlowLayoutPanel panel,string text,Action action)
  {
   var button=new Button{Text=text,AutoSize=false,Width=440,Height=33,Margin=new Padding(0,3,0,3)};
   button.Click+=(_,__)=>{try{action();}catch(Exception e){ProbeLog.Error("manager",e);MessageBox.Show(e is System.Runtime.InteropServices.COMException?"PowerPoint에 연결하지 못했습니다. PowerPoint를 먼저 열고 관리자 권한이 아닌 일반 사용자로 실행했는지 확인하세요.":e.Message,"검증 안내",MessageBoxButtons.OK,MessageBoxIcon.Information);}};panel.Controls.Add(button);
  }
 }
}
