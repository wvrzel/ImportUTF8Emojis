$source = [IO.File]::ReadAllText("$PSScriptRoot\ImportUTF8Emojis.cs")

$refs = (
  "System","System.Management",
  "Microsoft.Office.Interop.Word",
  "windowsbase",
  "PresentationFramework",
  "PresentationCore",
  "System.Xaml",
  "System.Reflection"
)

Add-Type -TypeDefinition $source -ReferencedAssemblies $refs -Language CSharp -OutputAssembly "$PSScriptRoot\ImportUTF8Emojis.exe" -OutputType "ConsoleApplication"