# Copilot Instructions for MsCrmTools.MetadataDocumentGenerator

## Project Overview

This is an XrmToolBox plugin that generates Excel and Word documentation from Microsoft Dynamics CRM/Dataverse metadata. The plugin allows users to select entities from their CRM environment and export detailed documentation about entities, attributes, relationships, and other metadata.

## Technology Stack

- **Language**: C# 
- **Target Framework**: .NET Framework 4.8
- **Build System**: MSBuild (Visual Studio project format)
- **Key Dependencies**:
  - `XrmToolBoxPackage` (1.2025.10.74) - XrmToolBox plugin framework
  - `EPPlus` (5.4.2) - Excel document generation
  - `DocumentFormat.OpenXml` (2.11.3) - Office Open XML document manipulation
  - `ILMerge` (3.0.41) - Assembly merging for deployment
  - Microsoft Dynamics CRM SDK libraries

## Project Structure

```
MsCrmTools.MetadataDocumentGenerator/
├── Forms/                          # UI forms
│   ├── PublisherSelector.cs        # Publisher selection dialog
│   └── SolutionPicker.cs           # Solution picker dialog
├── Generation/                     # Document generation
│   ├── ExcelDocument.cs            # Excel generation logic
│   ├── IDocument.cs                # Document interface
│   ├── WordDocument.cs             # Base Word generation
│   ├── WordDocumentDocX.cs         # Word generation (DocX)
│   └── WordDocumentOpenXml.cs      # Word generation (OpenXML)
├── Helper/                         # Utility classes
│   ├── AttributeMetadataComparer.cs
│   ├── Extensions.cs
│   ├── ListViewItemComparer.cs
│   ├── MetadataHelper.cs
│   └── OpenXmlHelper.cs
├── Resources/                      # Embedded resources
├── Properties/                     # Assembly info and resources
├── Plugin.cs                       # XrmToolBox plugin entry point
├── MetadataDocumentGenerator.cs    # Main user control
└── GenerationSettings.cs           # Configuration settings
```

## Build Requirements

1. **Visual Studio 2017 or later** with .NET Framework 4.8 Developer Pack
2. **NuGet Package Restore** must be enabled
3. **ILMerge** is used in Release builds to merge dependencies

### Build Process

- **Debug**: Copies output to `Plugins` subdirectory
- **Release**: Uses ILMerge to create a single DLL with merged dependencies (EPPlus.dll, DocumentFormat.OpenXml.dll)

### Build Commands

```bash
# Restore packages
nuget restore MsCrmTools.MetadataDocumentGenerator.sln

# Build Debug
msbuild MsCrmTools.MetadataDocumentGenerator.sln /p:Configuration=Debug

# Build Release
msbuild MsCrmTools.MetadataDocumentGenerator.sln /p:Configuration=Release
```

## Architecture

### Plugin Architecture
- **Entry Point**: `Plugin.cs` - Exports the plugin metadata for XrmToolBox
- **Main Control**: `MetadataDocumentGenerator.cs` - Windows Forms UserControl
- **Document Generation**: Strategy pattern with `IDocument` interface for different output formats

### Key Components

1. **Metadata Retrieval**: Uses Microsoft Dynamics CRM SDK to query entity metadata
2. **Document Generation**: 
   - `ExcelDocument` - Generates .xlsx files using EPPlus
   - `WordDocument*` classes - Generate .docx files using OpenXML
3. **UI Controls**: Windows Forms for entity selection and configuration

## Coding Conventions

### General Style
- **Namespace**: Single namespace per file matching folder structure
- **Indentation**: Default C# formatting (4 spaces)
- **Braces**: Opening brace on new line (Allman style)
- **Access Modifiers**: Always explicit

### Naming Conventions
- **Classes/Methods**: PascalCase
- **Private fields**: camelCase
- **Parameters**: camelCase
- **Constants**: PascalCase

### XrmToolBox Plugin Conventions
- Plugin class must implement `IXrmToolBoxPlugin` and be decorated with `[Export]` and `[ExportMetadata]` attributes
- Main control inherits from `PluginBase` or `PluginControlBase`
- Plugin metadata includes Name, Description, and Base64-encoded images

### Entity/Metadata Conventions
- Use `SchemaName` property for entity/attribute identifiers
- ListViews with entity data use the `Tag` property to store `SchemaName`
- Check for rollup derived columns using the `IsRollupDerivedColumn` extension method

## Common Patterns

### Extension Methods
The project uses extension methods for metadata operations:
```csharp
// Example: Check if an attribute is a rollup-derived column
if (amd.IsRollupDerivedColumn(allAttributes)) { ... }
```

### ListView Usage
- `lvEntities.CheckBoxes = true` - Entities are selectable via checkboxes
- `Tag` property contains entity SchemaName for later reference

## Testing

⚠️ **Note**: This project does not currently have automated tests. Testing is performed manually through XrmToolBox.

### Manual Testing Process
1. Build the solution in Debug mode
2. Copy output to XrmToolBox Plugins folder
3. Launch XrmToolBox and load the plugin
4. Connect to a CRM/Dataverse environment
5. Test entity selection and document generation

## Dependencies and Security

- Always check for security vulnerabilities in NuGet packages before updating
- Be cautious when updating XrmToolBoxPackage - ensure compatibility
- EPPlus 5.x uses the Polyform Noncommercial license - it's free for non-commercial use but requires a commercial license for commercial applications. Version 5.4.2 is currently in use.

## Common Tasks

### Adding a New Document Format
1. Create a new class implementing `IDocument` in the `Generation/` folder
2. Implement required methods for document generation
3. Update the main control to support the new format

### Modifying Entity Metadata Display
1. Check `MetadataHelper.cs` for metadata retrieval logic
2. Update document generation classes (`ExcelDocument.cs`, `WordDocument*.cs`)
3. Test with various entity types

### Updating XrmToolBox Integration
1. Modify `Plugin.cs` for plugin metadata changes
2. Update `MetadataDocumentGenerator.cs` for UI/control changes
3. Ensure compatibility with current XrmToolBox version

## Important Notes

- This is a **Windows Forms** application, not WPF
- The plugin must be compatible with XrmToolBox's plugin architecture
- ILMerge is used in production to create a single DLL
- The project uses Microsoft Dynamics CRM SDK for metadata operations
- Generated documents support multiple languages based on CRM language codes

## Resources

- [XrmToolBox Documentation](https://www.xrmtoolbox.com/)
- [Microsoft Dynamics CRM SDK](https://learn.microsoft.com/en-us/dynamics365/)
- [EPPlus Documentation](https://github.com/EPPlusSoftware/EPPlus)
- [Open XML SDK](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk)
