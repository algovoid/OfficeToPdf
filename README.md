OfficeToPdf - minimal Office Interop multi-threaded converter [Testing Phase]
================================================================

What it does
------------
Converts .doc/.docx and .xls/.xlsx/.xlsm files to PDF using Microsoft Office Interop.
Multiple worker threads run in parallel; each worker creates its own Office Application instance on an STA thread.

IMPORTANT WARNINGS
------------------
- Microsoft does NOT support Office Interop automation from unattended server processes. This tool must be run in an interactive desktop session whenever possible.
- Keep concurrency small (recommended 1-4). High concurrency can cause failures and orphaned processes.
- Run tests on a machine with Office installed and with the same user account that runs the program.

Prerequisites
-------------
- Windows Desktop OS (Office Interop requires Windows + Microsoft Office)
- Microsoft Office installed (Word and Excel). Tested with Office 2016 / Office 2019 / Microsoft 365 desktop installs.
- .NET 8.0 SDK for build (project targets net8.0 LTS in sample).
- NuGet packages: Microsoft.Office.Interop.Word and Microsoft.Office.Interop.Excel (see csproj example).

Build
-----
1. Clone/copy the project into a folder.
2. From project folder run:
   dotnet restore
   dotnet build -c Release

Run
---
Console usage:
  OfficeToPdf.exe <input-file-or-folder> <output-folder> [maxWorkers]

Examples:
- OfficeToPdf.exe "C:\toConvert" "C:\pdfs" 2
- OfficeToPdf.exe "C:\toConvert\report.docx" "C:\pdfs"

Notes
-----
- The program processes files in the top-level of the input folder only (no recursion). Modify `GatherOfficeFiles` to add recursion.
- Worker threads are STA and each owns its own Office Application instance.
- Logs to console; redirect console output to file to keep persistent logs.

Troubleshooting
---------------
- If Word/Excel processes remain after exit:
  - Check you didn't run multiple incompatible versions of Office on the same machine.
  - Ensure the process had permission to access desktop and COM. Restart machine if needed.
- If fonts are missing: ensure the machine has the required fonts installed.
- If file is locked: the converter will log and skip. Consider adding a retry/backoff in `OfficeWorker`.

Extending
---------
- Add PowerPoint support .
