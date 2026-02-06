using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace lnk_to_EXE
{
    public static class ShortcutParser
    {
        public static ShortcutInfo Parse(string lnkPath)
        {
            if (!File.Exists(lnkPath))
                throw new FileNotFoundException("Shortcut file not found.", lnkPath);

            IShellLink? link = (IShellLink?)new ShellLink();
            if (link == null)
                throw new InvalidOperationException("Failed to create ShellLink instance.");

            try
            {
                IPersistFile? file = (IPersistFile?)link;
                file?.Load(lnkPath, 0);

                StringBuilder targetPath = new(260);
                link.GetPath(targetPath, targetPath.Capacity, IntPtr.Zero, 0);

                StringBuilder arguments = new(260);
                link.GetArguments(arguments, arguments.Capacity);

                StringBuilder workingDir = new(260);
                link.GetWorkingDirectory(workingDir, workingDir.Capacity);

                StringBuilder iconPath = new(260);
                link.GetIconLocation(iconPath, iconPath.Capacity, out int iconIndex);

                return new ShortcutInfo
                {
                    TargetPath = targetPath.ToString(),
                    Arguments = arguments.ToString(),
                    WorkingDirectory = workingDir.ToString(),
                    IconPath = iconPath.ToString(),
                    IconIndex = iconIndex
                };
            }
            finally
            {
                Marshal.ReleaseComObject(link);
            }
        }

        [ComImport]
        [Guid("00021401-0000-0000-C000-000000000046")]
        private class ShellLink { }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214F9-0000-0000-C000-000000000046")]
        private interface IShellLink
        {
            void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, IntPtr pfd, int fFlags);
            void GetIDList(out IntPtr ppidl);
            void SetIDList(IntPtr pidl);
            void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
            void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
            void GetHotkey(out short pwHotkey);
            void SetHotkey(short wHotkey);
            void GetShowCmd(out int piShowCmd);
            void SetShowCmd(int iShowCmd);
            void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
            void Resolve(IntPtr hwnd, int fFlags);
            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("0000010b-0000-0000-C000-000000000046")]
        private interface IPersistFile
        {
            void GetClassID(out Guid pClassID);
            void IsDirty();
            void Load([In, MarshalAs(UnmanagedType.LPWStr)] string pszFileName, int dwMode);
            void Save([In, MarshalAs(UnmanagedType.LPWStr)] string pszFileName, [In, MarshalAs(UnmanagedType.Bool)] bool fRemember);
            void SaveCompleted([In, MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
            void GetCurFile([In, MarshalAs(UnmanagedType.LPWStr)] string ppszFileName);
        }
    }
}
