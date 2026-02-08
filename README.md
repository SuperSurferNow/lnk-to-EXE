# LNK to EXE Converter

Converts Windows shortcut (.lnk) files to standalone executables (.exe).

## Features
- ?? Convert .lnk shortcuts to portable .exe files
- ?? Automatically embeds icons from target applications
- ?? Dark mode support
- ?? Edit shortcut properties before conversion
- ?? Batch conversion support

## Requirements
- Windows 7 or later
- .NET 10 Runtime (for running the converter app)

## Generated Executables
The generated .exe files target .NET Framework 4.x, which is pre-installed on all modern Windows systems. No additional runtime is required to run the generated executables!

## Publishing as Portable

To create a portable single-file version of this application:

### Option 1: Using Visual Studio
1. Right-click on the project
2. Select "Publish"
3. Choose the "PortableWin64" profile
4. Click "Publish"

### Option 2: Using Command Line
```bash
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=true
```

The output will be in: `bin\Release\net10.0-windows\win-x64\publish\`

The single .exe file can be copied to any Windows system with .NET 10 runtime installed.

## Building from Source
```bash
dotnet build -c Release
```

## Usage
1. Launch the application
2. Add .lnk files (drag & drop or use the buttons)
3. Edit properties if needed (optional)
4. Select output folder
5. Click "Build All" or "Build Selected"
6. Your portable .exe files are ready!

## Icon Embedding
**Current Status:** Icon embedding is temporarily disabled due to _technical complexity._ (The code self-destruting every time I ask claude to fix it)

The generated .exe files will use the default Windows executable icon for now. The executables work perfectly - they just won't have custom icons yet.

**Workaround:** You can manually add icons using tools like:
- [Resource Hacker](http://www.angusj.com/resourcehacker/) (Free)
- [rcedit](https://github.com/electron/rcedit) (CLI tool)

**Future Plans:** Icon embedding will be re-implemented using a more robust approach.

## License
MIT License
