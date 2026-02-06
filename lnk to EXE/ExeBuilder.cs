using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Drawing;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using System.Linq;

namespace lnk_to_EXE
{
    public static class ExeBuilder
    {
        public static void Build(ShortcutInfo shortcut, string outputPath, bool embedIcon)
        {
            if (string.IsNullOrWhiteSpace(shortcut.TargetPath))
                throw new ArgumentException("Target path cannot be empty.");

            if (!File.Exists(shortcut.TargetPath))
                throw new FileNotFoundException("Target executable not found.", shortcut.TargetPath);

            string sourceCode = GenerateLauncherCode(shortcut);
            byte[] exeBytes = CompileToExe(sourceCode);
            File.WriteAllBytes(outputPath, exeBytes);

            if (embedIcon)
            {
                // Use icon from IconPath if specified, otherwise use target executable's icon
                string iconPath = !string.IsNullOrEmpty(shortcut.IconPath) 
                    ? shortcut.IconPath 
                    : shortcut.TargetPath;
                
                // Only try to embed if the icon source exists
                if (File.Exists(iconPath))
                {
                    TryEmbedIcon(outputPath, iconPath, shortcut.IconIndex);
                }
            }
        }

        private static string GenerateLauncherCode(ShortcutInfo shortcut)
        {
            // For verbatim strings (@""), we only need to escape quotes by doubling them
            string escapedTarget = shortcut.TargetPath.Replace("\"", "\"\"");
            string escapedArgs = shortcut.Arguments.Replace("\"", "\"\"");
            string escapedWorkDir = shortcut.WorkingDirectory.Replace("\"", "\"\"");

            return $$"""
                using System;
                using System.Diagnostics;

                class Program
                {
                    static void Main(string[] args)
                    {
                        try
                        {
                            var startInfo = new ProcessStartInfo
                            {
                                FileName = @"{{escapedTarget}}",
                                Arguments = @"{{escapedArgs}}",
                                WorkingDirectory = @"{{escapedWorkDir}}",
                                UseShellExecute = true
                            };

                            Process.Start(startInfo);
                        }
                        catch
                        {
                            // Silent fail - this is a launcher, not a diagnostic tool
                        }
                    }
                }
                """;
        }

        private static byte[] CompileToExe(string sourceCode)
        {
            SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(sourceCode);

            string assemblyName = Path.GetRandomFileName();
            
            // Use .NET Framework 4.x which is pre-installed on Windows
            string frameworkPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                @"Microsoft.NET\Framework64\v4.0.30319");
            
            // Fallback to 32-bit framework path if 64-bit not found
            if (!Directory.Exists(frameworkPath))
            {
                frameworkPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                    @"Microsoft.NET\Framework\v4.0.30319");
            }
            
            var references = new[]
            {
                MetadataReference.CreateFromFile(Path.Combine(frameworkPath, "mscorlib.dll")),
                MetadataReference.CreateFromFile(Path.Combine(frameworkPath, "System.dll")),
                MetadataReference.CreateFromFile(Path.Combine(frameworkPath, "System.Core.dll"))
            };

            CSharpCompilation compilation = CSharpCompilation.Create(
                assemblyName,
                syntaxTrees: new[] { syntaxTree },
                references: references,
                options: new CSharpCompilationOptions(
                    OutputKind.WindowsApplication,
                    optimizationLevel: OptimizationLevel.Release,
                    platform: Platform.AnyCpu));

            using var ms = new MemoryStream();
            EmitResult result = compilation.Emit(ms);

            if (!result.Success)
            {
                var failures = result.Diagnostics.Where(d => d.Severity == DiagnosticSeverity.Error);
                throw new InvalidOperationException($"Compilation failed: {string.Join(", ", failures)}");
            }

            return ms.ToArray();
        }

        private static void TryEmbedIcon(string exePath, string iconPath, int iconIndex)
        {
            try
            {
                if (string.IsNullOrEmpty(iconPath) || !File.Exists(iconPath))
                {
                    System.Diagnostics.Debug.WriteLine($"Icon embedding skipped: iconPath='{iconPath}'");
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"Starting icon embedding: {iconPath} (index {iconIndex})");

                // Extract icon to temporary .ico file
                string tempIconPath = Path.Combine(Path.GetTempPath(), $"lnk2exe_{Guid.NewGuid()}.ico");
                
                try
                {
                    // Extract and save icon as proper .ico file
                    if (ExtractIconToFile(iconPath, iconIndex, tempIconPath))
                    {
                        System.Diagnostics.Debug.WriteLine($"Icon extracted to temp file: {tempIconPath}");
                        
                        // Embed the icon into the executable
                        EmbedIconResource(exePath, tempIconPath);
                        
                        System.Diagnostics.Debug.WriteLine("? Icon embedding completed successfully!");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("? Failed to extract icon");
                    }
                }
                finally
                {
                    // Clean up temp file
                    try
                    {
                        if (File.Exists(tempIconPath))
                            File.Delete(tempIconPath);
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"? Icon embedding failed: {ex.Message}");
                // Don't throw - icon embedding is non-critical
            }
        }

        private static bool ExtractIconToFile(string sourcePath, int iconIndex, string destPath)
        {
            try
            {
                // If source is already a .ico file, just copy it
                if (sourcePath.EndsWith(".ico", StringComparison.OrdinalIgnoreCase))
                {
                    File.Copy(sourcePath, destPath, true);
                    return true;
                }

                // Extract all available icon sizes
                IntPtr[] largeIcons = new IntPtr[1];
                IntPtr[] smallIcons = new IntPtr[1];
                
                int count = ExtractIconEx(sourcePath, iconIndex, largeIcons, smallIcons, 1);
                
                if (count <= 0)
                {
                    System.Diagnostics.Debug.WriteLine($"No icons found at index {iconIndex}");
                    return false;
                }

                // Use large icon if available, otherwise small icon
                IntPtr hIcon = largeIcons[0] != IntPtr.Zero ? largeIcons[0] : smallIcons[0];
                
                if (hIcon == IntPtr.Zero)
                {
                    System.Diagnostics.Debug.WriteLine("Icon handle is null");
                    return false;
                }

                try
                {
                    // Create icon from handle and save
                    using (var icon = Icon.FromHandle(hIcon))
                    {
                        // Save icon to file stream
                        using (var fs = new FileStream(destPath, FileMode.Create, FileAccess.Write))
                        {
                            icon.Save(fs);
                        }
                    }
                    
                    return File.Exists(destPath) && new FileInfo(destPath).Length > 0;
                }
                finally
                {
                    if (largeIcons[0] != IntPtr.Zero) DestroyIcon(largeIcons[0]);
                    if (smallIcons[0] != IntPtr.Zero) DestroyIcon(smallIcons[0]);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ExtractIconToFile error: {ex.Message}");
                return false;
            }
        }

        private static void EmbedIconResource(string exePath, string iconPath)
        {
            // Read the icon file
            byte[] iconBytes = File.ReadAllBytes(iconPath);
            
            if (iconBytes.Length < 6)
            {
                System.Diagnostics.Debug.WriteLine("Invalid icon file - too small");
                return;
            }

            // Verify it's a valid ICO file (starts with 0x00 0x00 0x01 0x00)
            if (iconBytes[0] != 0 || iconBytes[1] != 0 || iconBytes[2] != 1 || iconBytes[3] != 0)
            {
                System.Diagnostics.Debug.WriteLine($"Invalid icon header: {iconBytes[0]:X2} {iconBytes[1]:X2} {iconBytes[2]:X2} {iconBytes[3]:X2}");
                return;
            }

            // Begin resource update
            IntPtr hUpdate = BeginUpdateResource(exePath, false);
            if (hUpdate == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                System.Diagnostics.Debug.WriteLine($"BeginUpdateResource failed with error {error}");
                return;
            }

            try
            {
                const int RT_ICON = 3;
                const int RT_GROUP_ICON = 14;

                using (var ms = new MemoryStream(iconBytes))
                using (var reader = new BinaryReader(ms))
                {
                    // Read ICONDIR
                    reader.ReadUInt16(); // Reserved
                    reader.ReadUInt16(); // Type
                    ushort iconCount = reader.ReadUInt16();

                    if (iconCount == 0 || iconCount > 20)
                    {
                        System.Diagnostics.Debug.WriteLine($"Invalid icon count: {iconCount}");
                        EndUpdateResource(hUpdate, true);
                        return;
                    }

                    System.Diagnostics.Debug.WriteLine($"Processing {iconCount} icon(s)");

                    // Build GRPICONDIR
                    using (var groupMs = new MemoryStream())
                    using (var groupWriter = new BinaryWriter(groupMs))
                    {
                        groupWriter.Write((ushort)0);  // Reserved
                        groupWriter.Write((ushort)1);  // Type
                        groupWriter.Write(iconCount);  // Count

                        for (ushort i = 0; i < iconCount; i++)
                        {
                            // Read ICONDIRENTRY
                            byte width = reader.ReadByte();
                            byte height = reader.ReadByte();
                            byte colorCount = reader.ReadByte();
                            byte reserved = reader.ReadByte();
                            ushort planes = reader.ReadUInt16();
                            ushort bitCount = reader.ReadUInt16();
                            uint imageSize = reader.ReadUInt32();
                            uint imageOffset = reader.ReadUInt32();

                            // Write GRPICONDIRENTRY
                            groupWriter.Write(width);
                            groupWriter.Write(height);
                            groupWriter.Write(colorCount);
                            groupWriter.Write(reserved);
                            groupWriter.Write(planes);
                            groupWriter.Write(bitCount);
                            groupWriter.Write(imageSize);
                            groupWriter.Write((ushort)(i + 1)); // Resource ID

                            // Save position and read icon image data
                            long currentPos = ms.Position;
                            ms.Seek(imageOffset, SeekOrigin.Begin);
                            byte[] imageData = reader.ReadBytes((int)imageSize);

                            // Add RT_ICON resource
                            if (!UpdateResource(hUpdate, (IntPtr)RT_ICON, (IntPtr)(i + 1), 0, imageData, (uint)imageData.Length))
                            {
                                int error = Marshal.GetLastWin32Error();
                                System.Diagnostics.Debug.WriteLine($"UpdateResource RT_ICON #{i + 1} failed: error {error}");
                            }

                            ms.Seek(currentPos, SeekOrigin.Begin);
                        }

                        // Add RT_GROUP_ICON resource
                        byte[] groupData = groupMs.ToArray();
                        if (!UpdateResource(hUpdate, (IntPtr)RT_GROUP_ICON, (IntPtr)1, 0, groupData, (uint)groupData.Length))
                        {
                            int error = Marshal.GetLastWin32Error();
                            System.Diagnostics.Debug.WriteLine($"UpdateResource RT_GROUP_ICON failed: error {error}");
                        }
                    }
                }

                // Commit the changes
                if (!EndUpdateResource(hUpdate, false))
                {
                    int error = Marshal.GetLastWin32Error();
                    System.Diagnostics.Debug.WriteLine($"EndUpdateResource failed: error {error}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Exception in EmbedIconResource: {ex.Message}");
                EndUpdateResource(hUpdate, true); // Discard on error
            }
        }



        #region Win32 API

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern int ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, int nIcons);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr BeginUpdateResource(string pFileName, bool bDeleteExistingResources);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool UpdateResource(IntPtr hUpdate, IntPtr lpType, IntPtr lpName, ushort wLanguage, byte[] lpData, uint cbData);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool EndUpdateResource(IntPtr hUpdate, bool fDiscard);

        #endregion
    }
}
