# Static checks that catch mistakes no unit test would: syntax Windows PowerShell 5.1
# can't run, non-ASCII bytes (5.1 reads BOM-less files as ANSI), $var / $VAR clashes
# (PowerShell variables ignore case), commands that don't exist, and XAML that doesn't
# match the code that drives it.

BeforeDiscovery {
                $script:Root = Split-Path $PSScriptRoot -Parent
                $script:SourceFiles = @(Get-ChildItem -Path (Join-Path $script:Root 'Modules') -Filter '*.psm1') + @(Get-Item (Join-Path $script:Root 'FAIS.ps1')) |
                                Where-Object { $_ } | ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } }
}

BeforeAll {
                $script:Root = Split-Path $PSScriptRoot -Parent

                function Get-Ast {
                                param ([string]$Path)
                                $tokens = $null; $errors = $null
                                $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
                                return @{ Ast = $ast; Tokens = $tokens; Errors = $errors }
                }
}

Describe 'Source file <Name>' -ForEach $script:SourceFiles {
                BeforeAll {
                                $parsed = Get-Ast $Path
                }

                It 'parses without errors' {
                                @($parsed.Errors | ForEach-Object { "$($_.Extent.StartLineNumber): $($_.Message)" }) | Should -BeNullOrEmpty
                }

                It 'is plain ASCII' {
                                $bytes = [System.IO.File]::ReadAllBytes($Path)
                                $bad = @(for ($i = 0; $i -lt $bytes.Length; $i++) { if ($bytes[$i] -gt 127) { $i; break } })
                                $bad | Should -BeNullOrEmpty
                }

                It 'uses no PowerShell 7-only syntax' {
                                $found = $parsed.Ast.FindAll({
                                                param($n)
                                                $t = $n.GetType().Name
                                                $t -eq 'TernaryExpressionAst' -or $t -eq 'PipelineChainAst' -or
                                                ($n -is [System.Management.Automation.Language.BinaryExpressionAst] -and $n.Operator.ToString() -eq 'QuestionQuestion') -or
                                                ($n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Operator.ToString() -eq 'QuestionQuestionEquals') -or
                                                ($n -is [System.Management.Automation.Language.MemberExpressionAst] -and $n.NullConditional)
                                }, $true)
                                @($found | ForEach-Object { "$($_.Extent.StartLineNumber): $($_.Extent.Text)" }) | Should -BeNullOrEmpty
                }

                It 'has no variables that differ only by case inside one function' {
                                $clashes = foreach ($fn in $parsed.Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
                                                $names = @{}
                                                foreach ($v in $fn.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
                                                                $raw = $v.VariablePath.UserPath
                                                                if ($raw -match '^(script|global|env|using):') { continue }
                                                                $k = $raw.ToLowerInvariant()
                                                                if (-not $names.ContainsKey($k)) { $names[$k] = New-Object System.Collections.Generic.HashSet[string] }
                                                                [void]$names[$k].Add($raw)
                                                }
                                                foreach ($k in $names.Keys) {
                                                                $variants = @($names[$k])
                                                                if ($variants.Count -gt 1) { "$($fn.Name): $($variants -join ' / ')" }
                                                }
                                }
                                @($clashes) | Should -BeNullOrEmpty
                }

                It 'never wraps a generic List in @() (fails in PowerShell 7 and is easy to get wrong)' {
                                $listVars = New-Object System.Collections.Generic.HashSet[string]
                                foreach ($a in $parsed.Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)) {
                                                if ($a.Right.Extent.Text -match 'Generic\.List\[object\]' -and $a.Left -is [System.Management.Automation.Language.VariableExpressionAst]) {
                                                                [void]$listVars.Add($a.Left.VariablePath.UserPath.ToLowerInvariant())
                                                }
                                }
                                $bad = foreach ($sub in $parsed.Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ArrayExpressionAst] }, $true)) {
                                                $inner = $sub.SubExpression.Statements
                                                if ($inner.Count -ne 1) { continue }
                                                $text = $inner[0].Extent.Text.Trim()
                                                if ($text -match '^\$(\w+)$' -and $listVars.Contains($Matches[1].ToLowerInvariant())) { "$($sub.Extent.StartLineNumber): $($sub.Extent.Text)" }
                                }
                                @($bad) | Should -BeNullOrEmpty
                }

                It 'never wraps the board''s Lines list in @() (it is a generic List[object])' {
                                $bad = foreach ($sub in $parsed.Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ArrayExpressionAst] }, $true)) {
                                                if ($sub.SubExpression.Statements.Count -eq 1 -and $sub.SubExpression.Statements[0].Extent.Text.Trim() -match '(?i)board\.(Lines|Undo)$') { "$($sub.Extent.StartLineNumber): $($sub.Extent.Text)" }
                                }
                                @($bad) | Should -BeNullOrEmpty
                }
}

Describe 'Module set' {
                BeforeAll {
                                $script:ModuleFiles = @(Get-ChildItem -Path (Join-Path $script:Root 'Modules') -Filter '*.psm1')
                                $script:Defined = @{}
                                foreach ($f in $script:ModuleFiles) {
                                                $p = Get-Ast $f.FullName
                                                foreach ($fn in $p.Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
                                                                if (-not $script:Defined.ContainsKey($fn.Name)) { $script:Defined[$fn.Name] = New-Object System.Collections.Generic.List[string] }
                                                                $script:Defined[$fn.Name].Add($f.Name)
                                                }
                                }
                }

                It 'defines every function exactly once across Modules' {
                                @($script:Defined.Keys | Where-Object { $script:Defined[$_].Count -gt 1 } | ForEach-Object { "$_ in $($script:Defined[$_] -join ', ')" }) | Should -BeNullOrEmpty
                }

                It 'only calls commands that exist (module functions, built-ins or ImportExcel)' {
                                $importExcel = @('Import-Excel', 'Export-Excel', 'Open-ExcelPackage', 'Close-ExcelPackage')
                                $missing = foreach ($f in @($script:ModuleFiles) + @(Get-Item (Join-Path $script:Root 'FAIS.ps1'))) {
                                                $p = Get-Ast $f.FullName
                                                $nested = @($p.Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
                                                foreach ($cmd in $p.Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
                                                                $name = $cmd.GetCommandName()
                                                                if ([string]::IsNullOrEmpty($name)) { continue }
                                                                if ($script:Defined.ContainsKey($name) -or $nested -contains $name -or $importExcel -contains $name) { continue }
                                                                if ($null -ne (Get-Command $name -ErrorAction SilentlyContinue)) { continue }
                                                                if ($name -like '*.ps1' -or $name -like '*\*' -or $name -like '*/*') { continue }
                                                                "$($f.Name):$($cmd.Extent.StartLineNumber) $name"
                                                }
                                }
                                @($missing | Select-Object -Unique) | Should -BeNullOrEmpty
                }

                It 'imports every module cleanly' {
                                . (Join-Path $PSScriptRoot 'Fixtures/New-FaisTestData.ps1')
                                { Import-FaisModules -IncludeUI } | Should -Not -Throw
                }
}

Describe 'PuckBoard.xaml' {
                BeforeAll {
                                $script:XamlPath = Join-Path $script:Root 'Modules/PuckBoard.xaml'
                                $script:XamlText = [System.IO.File]::ReadAllText($script:XamlPath)
                                $script:BoardCode = [System.IO.File]::ReadAllText((Join-Path $script:Root 'Modules/PuckBoard.psm1'))
                }

                It 'is well-formed XML' {
                                { [xml]$script:XamlText } | Should -Not -Throw
                }

                It 'has no code-behind hooks XamlReader cannot load (x:Class, Click=, mc:)' {
                                $script:XamlText | Should -Not -Match 'x:Class='
                                $script:XamlText | Should -Not -Match '\sClick='
                                $script:XamlText | Should -Not -Match 'mc:Ignorable'
                }

                It 'has every named element the board code looks up' {
                                $names = @([regex]::Matches($script:XamlText, 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
                                $listBlock = [regex]::Match($script:BoardCode, "foreach \(\`$name in @\(([^)]*)\)\)").Groups[1].Value
                                $wanted = @([regex]::Matches($listBlock, "'([^']+)'") | ForEach-Object { $_.Groups[1].Value })
                                $wanted.Count | Should -BeGreaterThan 10
                                @($wanted | Where-Object { $names -notcontains $_ }) | Should -BeNullOrEmpty
                }

                It 'has every resource key the board code asks for' {
                                $keys = @([regex]::Matches($script:XamlText, 'x:Key="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
                                $asked = @([regex]::Matches($script:BoardCode, "FindResource\('([^']+)'\)") | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
                                @($asked | Where-Object { $keys -notcontains $_ }) | Should -BeNullOrEmpty
                }

                It 'only binds to properties the view models provide' {
                                $bound = @([regex]::Matches($script:XamlText, '\{Binding (\w+)') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
                                $provided = @([regex]::Matches($script:BoardCode, '(?m)^\s+(\w+)\s+=\s') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
                                @($bound | Where-Object { $provided -notcontains $_ }) | Should -BeNullOrEmpty
                }

                It 'uses only Tags the board code knows' {
                                $tags = @([regex]::Matches($script:XamlText, '\bTag="([^"{]+)"') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
                                foreach ($t in $tags) { $script:BoardCode | Should -Match ([regex]::Escape("'$t'")) }
                }
}

Describe 'LoadDayDialog.xaml' {
                BeforeAll {
                                $script:DlgText = [System.IO.File]::ReadAllText((Join-Path $script:Root 'Modules/LoadDayDialog.xaml'))
                                $script:BoardCode2 = [System.IO.File]::ReadAllText((Join-Path $script:Root 'Modules/PuckBoard.psm1'))
                }

                It 'is well-formed XML with no code-behind hooks' {
                                { [xml]$script:DlgText } | Should -Not -Throw
                                $script:DlgText | Should -Not -Match 'x:Class='
                                $script:DlgText | Should -Not -Match '\sClick='
                }

                It 'has every named element the board code looks up' {
                                $names = @([regex]::Matches($script:DlgText, 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
                                $listBlock = [regex]::Match($script:BoardCode2, "foreach \(\`$n in @\(([^)]*)\)\)").Groups[1].Value
                                $wanted = @([regex]::Matches($listBlock, "'([^']+)'") | ForEach-Object { $_.Groups[1].Value })
                                $wanted.Count | Should -BeGreaterThan 5
                                @($wanted | Where-Object { $names -notcontains $_ }) | Should -BeNullOrEmpty
                }

                It 'is plain ASCII' {
                                $bytes = [System.IO.File]::ReadAllBytes((Join-Path $script:Root 'Modules/LoadDayDialog.xaml'))
                                @($bytes | Where-Object { $_ -gt 127 }).Count | Should -Be 0
                }
}

Describe 'BoardLogo.png' {
                It 'is a PNG with a transparent background, small enough to ship' {
                                $path = Join-Path $script:Root 'Modules/BoardLogo.png'
                                Test-Path -LiteralPath $path | Should -BeTrue
                                $bytes = [System.IO.File]::ReadAllBytes($path)
                                (@($bytes[0..7]) -join ',') | Should -BeExactly '137,80,78,71,13,10,26,10'
                                $bytes[25] | Should -Be 6 -Because 'colour type 6 = RGBA (the background is see-through)'
                                $bytes.Length | Should -BeLessThan 200KB
                }

                It 'is the file the board code loads' {
                                $code = [System.IO.File]::ReadAllText((Join-Path $script:Root 'Modules/PuckBoard.psm1'))
                                $code | Should -Match ([regex]::Escape('Join-Path $PSScriptRoot ''BoardLogo.png'''))
                }
}

Describe 'BulkLinesDialog.xaml' {
                BeforeAll {
                                $script:BulkText = [System.IO.File]::ReadAllText((Join-Path $script:Root 'Modules/BulkLinesDialog.xaml'))
                                $script:BoardCode3 = [System.IO.File]::ReadAllText((Join-Path $script:Root 'Modules/PuckBoard.psm1'))
                }

                It 'is well-formed XML with no code-behind hooks' {
                                { [xml]$script:BulkText } | Should -Not -Throw
                                $script:BulkText | Should -Not -Match 'x:Class='
                                $script:BulkText | Should -Not -Match '\sClick='
                }

                It 'has every named element the board code looks up' {
                                $names = @([regex]::Matches($script:BulkText, 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
                                $listBlock = [regex]::Match($script:BoardCode3, "foreach \(\`$field in @\(([^)]*)\)\)").Groups[1].Value
                                $wanted = @([regex]::Matches($listBlock, "'([^']+)'") | ForEach-Object { $_.Groups[1].Value })
                                $wanted.Count | Should -BeGreaterThan 5
                                @($wanted | Where-Object { $names -notcontains $_ }) | Should -BeNullOrEmpty
                }

                It 'is plain ASCII' {
                                $bytes = [System.IO.File]::ReadAllBytes((Join-Path $script:Root 'Modules/BulkLinesDialog.xaml'))
                                @($bytes | Where-Object { $_ -gt 127 }).Count | Should -Be 0
                }
}
