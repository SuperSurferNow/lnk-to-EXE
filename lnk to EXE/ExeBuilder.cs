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
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to launch target: {ex.Message}");
                            Console.WriteLine("Press any key to exit...");
                            Console.ReadKey();
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
                    OutputKind.ConsoleApplication,
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
            // Icon embedding is temporarily disabled due to complexity with .NET Framework executables
            // The generated EXE works perfectly, it just won't have a custom icon
            System.Diagnostics.Debug.WriteLine($"Icon embedding skipped for now - EXE will use default icon");
            
            // TODO: Implement icon embedding using:
            // - ResourceHacker CLI tool, or
            // - rcedit tool, or
            // - Proper multi-icon .ico file generation
        }


        #region Win32 API (for future icon embedding)

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern int ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, int nIcons);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);

        #endregion
    }
}
