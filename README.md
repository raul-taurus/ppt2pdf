# ppt2pdf

The program only works for Windows with Microsoft Office installed.


## Build

Windows has the latest Visual Studio installed.

Need Office primary interop assemblies

ppt2pdf.csproj: 
```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net6.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <Platforms>AnyCPU</Platforms>
  </PropertyGroup>

  <ItemGroup>
    <Reference Include="Microsoft.Office.Interop.PowerPoint">
      <HintPath>C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15\Microsoft.Office.Interop.PowerPoint.dll</HintPath>
    </Reference>
    <Reference Include="office">
      <HintPath>C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15\Office.dll</HintPath>
    </Reference>
  </ItemGroup>

</Project>

```