using System.Runtime.InteropServices;
using static ClipImage.NativeMethods;

namespace ClipImage
{
    static class UI
    {
        public static void ShowMessage(string message)
        {
            NativeMethods.MessageBox(IntPtr.Zero, message, "ClipImage", NativeMethods.MB_ICONINFORMATION);
        }

        public static void ShowError(string message)
        {
            NativeMethods.MessageBox(IntPtr.Zero, message, "ClipImage", NativeMethods.MB_ICONERROR);
        }

        public static bool Initialize()
        {
            var hr = NativeMethods.CoInitializeEx(IntPtr.Zero, NativeMethods.COINIT_APARTMENTTHREADED);
            if (hr < 0)
            {
                Console.Error.WriteLine($"初期化に失敗しました。{hr:X8}");
                return false;
            }
            return true;
        }

        public static void UnInitialize()
        {
            NativeMethods.CoUninitialize();
        }

        public static bool ShowSaveFiledialog(string location, out string dest)
        {
            dest = string.Empty;
            var hr = NativeMethods.CoCreateInstance(
                new Guid(ShellGUID.CLSID_FileSaveDialog),
                IntPtr.Zero,
                NativeMethods.CLSCTX_INPROC_SERVER,
                new Guid(ShellGUID.IID_IFileSaveDialog),
                out IFileSaveDialog dialog);

            if (hr != 0)
            {
                ShowError("保存ダイアログが作成できません。");
                return false;
            }

            IShellItem? initialFolderItem = null;
            try
            {
                NativeMethods.SHCreateItemFromParsingName(
                    location,
                    IntPtr.Zero,
                    new Guid(ShellGUID.IID_IShellItem),
                    out initialFolderItem
                );
                dest = ShowSaveFileDialogInternal(dialog, initialFolderItem);
                return !string.IsNullOrEmpty(dest);
            }
            catch (Exception ex)
            {
                ShowError($"保存ダイアログの表示に失敗しました。{ex.Message}");
                return false;
            }
            finally
            {
                Marshal.ReleaseComObject(dialog);
                if (initialFolderItem != null)
                    Marshal.ReleaseComObject(initialFolderItem);
            }
        }

        private static string ShowSaveFileDialogInternal(IFileSaveDialog dialog, IShellItem initialFolderItem)
        {
            dialog.SetTitle("保存場所を選択");
            dialog.SetDefaultFolder(initialFolderItem);
            dialog.SetFolder(initialFolderItem);

            var filters = new[]
            {
                new COMDLG_FILTERSPEC { pszName = "PNG (*.png)", pszSpec = "*.png" },
                new COMDLG_FILTERSPEC { pszName = "JPEG (*.jpg;*.jpeg)", pszSpec = "*.jpg;*.jpeg" },
                new COMDLG_FILTERSPEC { pszName = "すべて (*.*)", pszSpec = "*.*" },
            };
            dialog.SetFileTypes((uint)filters.Length, filters);

            if (dialog.Show(IntPtr.Zero) != 0)
                return string.Empty;

            IntPtr pPath = IntPtr.Zero;
            IShellItem? result = null;
            try
            {
                dialog.GetResult(out result);
                result.GetDisplayName(SIGDN.FILESYSPATH, out pPath);
                return Marshal.PtrToStringUni(pPath) ?? string.Empty;
            }
            finally
            {
                if (pPath != IntPtr.Zero)
                    Marshal.FreeHGlobal(pPath);
                if (result != null)
                    Marshal.ReleaseComObject(result);
            }
        }
    }
}