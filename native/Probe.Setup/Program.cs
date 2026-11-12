using System;
using System.Windows.Forms;
namespace PPTAIStudio.Probe.Setup
{
 internal static class Program
 {
  [STAThread] private static void Main(string[] args)
  {
   Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);
   if(args.Length==1&&args[0]=="--uninstall"){try{Registration.Uninstall();MessageBox.Show("추가 기능 등록을 해제했습니다. 파일은 직접 삭제할 수 있습니다.");}catch(Exception e){MessageBox.Show(e.Message,"등록 해제");}return;}
   Application.Run(new ManagerForm());
  }
 }
}
