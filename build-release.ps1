# PowerShell script to build and publish the tool
$project = "OfficeToPdf.csproj"
$output = "publish"

Write-Host "Restoring and building project..."
dotnet restore $project
dotnet build $project -c Release

Write-Host "Publishing self-contained Windows x64 executable..."
dotnet publish $project -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o $output

Write-Host "Publish complete. Output in /$output"
Write-Host "You can now zip that folder as a deliverable."
