using System.Runtime.InteropServices;

namespace ClipImage
{
    static class NativeMethods
    {
        //―― GUID 群 ――
        public static class ShellGUID
        {
            public const string CLSID_FileSaveDialog = "C0B4E2F3-BA21-4773-8DBA-335EC946EB8B";
            public const string IID_IFileDialog = "42F85136-DB7E-439C-85F1-E4075D135FC8";
            public const string IID_IFileSaveDialog = "84BCCD23-5FDE-4CDB-AEA4-AF64B83D78AB";
            public const string IID_IShellItem = "43826D1E-E718-42EE-BC55-A1E261C37BFE";
        }

        //―― 共通構造体／列挙体 ――
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct COMDLG_FILTERSPEC
        {
            public string pszName;
            public string pszSpec;
        }

        public enum SIGDN : uint
        {
            SIGDN_FILESYSPATH = 0x80058000,
        }

        [Flags]
        public enum FOS : uint
        {
            FOS_OVERWRITEPROMPT = 0x2,
        }

        public enum FDAP : uint
        {
            FDAP_BOTTOM = 0,
            FDAP_TOP = 1,
        }

        [ComImport]
        [Guid(ShellGUID.IID_IFileDialog)]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFileDialog
        {
            // ----- IModalWindow -----
            [PreserveSig]
            int Show(IntPtr hwndOwner);

            // ----- IFileDialog -----
            void SetFileTypes(uint cFileTypes, [MarshalAs(UnmanagedType.LPArray)] COMDLG_FILTERSPEC[] rgFilterSpec);
            void SetFileTypeIndex(uint iFileType);
            void GetFileTypeIndex(out uint piFileType);

            void Advise(IntPtr pfde, out uint pdwCookie);        // IFileDialogEvents は null で OK
            void Unadvise(uint dwCookie);

            void SetOptions(FOS fos);
            void GetOptions(out FOS pfos);

            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder(out IShellItem ppsi);

            void GetCurrentSelection(out IShellItem ppsi);

            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName(out IntPtr pszName);

            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);

            void GetResult(out IShellItem ppsi);
            void AddPlace(IShellItem psi, FDAP fdap);

            void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);

            void Close(int hr);
            void SetClientGuid(ref Guid guid);
            void ClearClientData();
            void SetFilter(IntPtr pFilter);  // IFileDialog events 用
        }

        //―― IFileSaveDialog ――
        [ComImport]
        [Guid(ShellGUID.IID_IFileSaveDialog)]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFileSaveDialog : IFileDialog
        {
            // Save 用の追加メソッド（今回使わなくても v-table 合わせのために定義）
            void SetSaveAsItem(IShellItem psi);
            void SetProperties(IntPtr pStore);
            void SetCollectedProperties(IntPtr pList, int fAppendDefault);
            void GetProperties(out IntPtr ppStore);
            void ApplyProperties(IShellItem psi, IntPtr pStore, IntPtr hwnd, uint flags);
        }

        //―― IShellItem ――
        [ComImport]
        [Guid(ShellGUID.IID_IShellItem)]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IShellItem
        {
            void BindToHandler();      // 省略
            void GetParent();          // 省略
            void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
        }

        [DllImport("ole32.dll")]
        public static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

        [DllImport("ole32.dll")]
        public static extern void CoUninitialize();

        [DllImport("ole32.dll", ExactSpelling = true)]
        public static extern int CoCreateInstance(
          [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
          IntPtr pUnkOuter,
          uint dwClsContext,
          [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
          out IFileSaveDialog ppv);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        public static extern void SHCreateItemFromParsingName(
           string pszPath,
           IntPtr pbc,  // NULL で OK
           [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
           out IShellItem ppv
        );

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int MessageBox(IntPtr hWnd, string lpText, string lpCaption, uint uType);

        public const uint COINIT_APARTMENTTHREADED = 0x2; // STA
        public const uint CLSCTX_INPROC_SERVER = 0x1;

        public const uint MB_ICONERROR = 0x00000010;
        public const uint MB_ICONWARNING = 0x00000030;
        public const uint MB_ICONINFORMATION = 0x00000040;
    }

}
