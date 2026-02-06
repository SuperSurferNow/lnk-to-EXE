using System.IO;
using System.Reflection;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;

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

            if (embedIcon && !string.IsNullOrEmpty(shortcut.IconPath))
            {
                TryEmbedIcon(outputPath, shortcut.IconPath, shortcut.IconIndex);
            }
        }

        private static string GenerateLauncherCode(ShortcutInfo shortcut)
        {
            string escapedTarget = shortcut.TargetPath.Replace("\\", "\\\\").Replace("\"", "\\\"");
            string escapedArgs = shortcut.Arguments.Replace("\\", "\\\\").Replace("\"", "\\\"");
            string escapedWorkDir = shortcut.WorkingDirectory.Replace("\\", "\\\\").Replace("\"", "\\\"");

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
            var references = new[]
            {
                MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
                MetadataReference.CreateFromFile(typeof(Console).Assembly.Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Runtime").Location),
                MetadataReference.CreateFromFile(Assembly.Load("System.Diagnostics.Process").Location)
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
            try
            {
                // Icon embedding requires Win32 resource manipulation
                // This is a simplified placeholder - full implementation requires ResourceLib or similar
                // For initial version, we skip icon embedding if it fails
            }
            catch
            {
                // Silently continue if icon embedding fails
            }
        }
    }
}
