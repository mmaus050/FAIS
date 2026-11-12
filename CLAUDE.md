# FAIS – instructions for Claude

## Environment limits (read first)
- This is a Windows PowerShell 5.1 + WPF + ImportExcel app. Claude's cloud sandbox is
  Linux with NO PowerShell, NO Pester, NO Excel, and no GUI.
- Do NOT try to install pwsh/Pester/ImportExcel, download anything, or run the tests
  or FAIS.ps1. Do not report "blocked" or "couldn't run tests" – that is expected.
- Make the code change by reading and editing the code carefully, then stop.

## Workflow
- Keep changes minimal and match the existing style in Modules/.
- If a change affects behavior covered by a test in Tests/, update or add the test
  in the matching *.Tests.ps1 file, but don't try to execute it.
- At the end, give a 2–3 line summary of what changed and which test file(s) I should
  run on my Windows machine (Tests\Invoke-SelfTest.ps1, or Invoke-Pester).
- Don't create a pull request unless I ask.
