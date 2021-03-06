using System;
using System.Collections.Generic;
using System.Text;
using WindowsLive.Writer.Api;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Drawing;
using System.Runtime.InteropServices;

namespace LiveWriterSnippingPlugin
{
  [
    WriterPlugin(
      "5E3ADDAE-EFDF-4f05-B227-296C903520C7", 
      "Snipped Image",
      PublisherUrl = "http://www.mtaulty.com",
      Description = "Launches the snipping tool and inserts the picture.",
      ImagePath="snip.png")]

  [InsertableContentSource("Snipped Image")]



  public class PlugIn : ContentSource
  {
    
      public static bool Is64BitOperatingSystem()
      {
          if (IntPtr.Size == 8)  // 64-bit programs run only on Win64
          {
              return true;
          }
          else  // 32-bit programs run on both 32-bit and 64-bit Windows
          {
              // Detect whether the current process is a 32-bit process 
              // running on a 64-bit system.
              bool flag;
              return ((DoesWin32MethodExist("kernel32.dll", "IsWow64Process") &&
                  IsWow64Process(GetCurrentProcess(), out flag)) && flag);
          }
      }

      /// <summary>
      /// The function determins whether a method exists in the export 
      /// table of a certain module.
      /// </summary>
      /// <param name="moduleName">The name of the module</param>
      /// <param name="methodName">The name of the method</param>
      /// <returns>
      /// The function returns true if the method specified by methodName 
      /// exists in the export table of the module specified by moduleName.
      /// </returns>
      static bool DoesWin32MethodExist(string moduleName, string methodName)
      {
          IntPtr moduleHandle = GetModuleHandle(moduleName);
          if (moduleHandle == IntPtr.Zero)
          {
              return false;
          }
          return (GetProcAddress(moduleHandle, methodName) != IntPtr.Zero);
      }

      [DllImport("kernel32.dll")]
      static extern IntPtr GetCurrentProcess();

      [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
      static extern IntPtr GetModuleHandle(string moduleName);

      [DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
      static extern IntPtr GetProcAddress(IntPtr hModule,
          [MarshalAs(UnmanagedType.LPStr)]string procName);

      [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
      [return: MarshalAs(UnmanagedType.Bool)]
      static extern bool IsWow64Process(IntPtr hProcess, out bool wow64Process);


    public override DialogResult CreateContent(
      IWin32Window dialogOwner, ref string newContent)
    {
      DialogResult result = DialogResult.Abort;
      newContent = "";

      try
      {
          string path = "";

          if (Is64BitOperatingSystem())
          {
              path = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\sysnative\SnippingTool.exe");
          }
          else
          {
              path = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\SnippingTool.exe");
          }
          Process p = Process.Start(path);
          p.WaitForExit();

        // Grab the image data from the clipboard.
        if (Clipboard.ContainsImage())
        {
          string snipName = string.Format("{0}snip.png",Guid.NewGuid().ToString());
          string tempImagePath = Environment.ExpandEnvironmentVariables(string.Format(@"%temp%\{0}",snipName));
          Image tempImage = Clipboard.GetImage();
          tempImage.Save(tempImagePath);

          //newContent.Files.AddImage(snipName, tempImage,
          //  ImageFormat.Bmp);
          newContent=string.Format(@"<img src='file://{0}' />" , tempImagePath);
          result = DialogResult.OK;
        }
        else
        {
          PluginDiagnostics.DisplayError("Error finding image data",
            "No image data found on clipboard");
        }
      }
      catch (Exception ex)
      {
        PluginDiagnostics.DisplayError("Launching Snipping Tool",
          ex.Message);
      }
      return (result);
    }
  }
}