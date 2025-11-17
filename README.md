# Analyse_Recursive-VB.NET
  Texte de l'exception 
---------------
System.ArgumentException: Impossible de marshaler le type, car la longueur d'une instance de tableau incorporée ne correspond pas à la longueur déclarée dans la disposition.
   à System.StubHelpers.ValueClassMarshaler.ConvertToNative(IntPtr dst, IntPtr src, IntPtr pMT, CleanupWorkList& pCleanupWorkList)
   à Analyse_Recursive.Form1.FindNextFile(Int32 hFindFile, WIN32_FIND_DATA& lpFindFileData)
   à Analyse_Recursive.Form1.FindFilesAPI(String Path, String SearchStr, Int16 FileCount, Int16 DirCount)
   à Analyse_Recursive.Form1.DoFileSystemSearch(String sPath, String sFilter, ListPaths ListAction)
   à Analyse_Recursive.Form1.Command1_Click(Object eventSender, EventArgs eventArgs)
   à System.Windows.Forms.Control.OnClick(EventArgs e)
   à System.Windows.Forms.Button.OnClick(EventArgs e)
   à System.Windows.Forms.Button.OnMouseUp(MouseEventArgs mevent)
   à System.Windows.Forms.Control.WmMouseUp(Message& m, MouseButtons button, Int32 clicks)
   à System.Windows.Forms.Control.WndProc(Message& m)
   à System.Windows.Forms.ButtonBase.WndProc(Message& m)
   à System.Windows.Forms.Button.WndProc(Message& m)
   à System.Windows.Forms.Control.ControlNativeWindow.OnMessage(Message& m)
   à System.Windows.Forms.Control.ControlNativeWindow.WndProc(Message& m)
   à System.Windows.Forms.NativeWindow.Callback(IntPtr hWnd, Int32 msg, IntPtr wparam, IntPtr lparam)
