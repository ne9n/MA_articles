$ErrorActionPreference = "Stop"

$docxPath = "c:\Users\uline\OneDrive\Documents\ma_articles\model-aviation-kidventure-submission.docx"
$mdPath = "c:\Users\uline\OneDrive\Documents\ma_articles\model-aviation-kidventure-submission.md"

$content = Get-Content -Path $mdPath -Encoding UTF8

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Add()
    
    # 1 inch margins (72 points)
    $doc.PageSetup.TopMargin = 72
    $doc.PageSetup.BottomMargin = 72
    $doc.PageSetup.LeftMargin = 72
    $doc.PageSetup.RightMargin = 72
    
    $selection = $word.Selection
    
    foreach ($line in $content) {
        $trimmed = $line.Trim()
        
        if ($trimmed.StartsWith("# ")) {
            $selection.Font.Name = "Times New Roman"
            $selection.Font.Size = 18
            $selection.Font.Bold = 1
            $selection.ParagraphFormat.SpaceBefore = 12
            $selection.ParagraphFormat.SpaceAfter = 6
            $selection.TypeText($trimmed.Substring(2))
            $selection.TypeParagraph()
        }
        elseif ($trimmed.StartsWith("## ")) {
            $selection.Font.Name = "Times New Roman"
            $selection.Font.Size = 14
            $selection.Font.Bold = 1
            $selection.ParagraphFormat.SpaceBefore = 14
            $selection.ParagraphFormat.SpaceAfter = 4
            $selection.TypeText($trimmed.Substring(3))
            $selection.TypeParagraph()
        }
        elseif ($trimmed.StartsWith("### ")) {
            $selection.Font.Name = "Times New Roman"
            $selection.Font.Size = 12
            $selection.Font.Bold = 1
            $selection.ParagraphFormat.SpaceBefore = 10
            $selection.ParagraphFormat.SpaceAfter = 3
            $selection.TypeText($trimmed.Substring(4))
            $selection.TypeParagraph()
        }
        elseif ($trimmed.StartsWith("> ")) {
            $selection.Font.Name = "Times New Roman"
            $selection.Font.Size = 12
            $selection.Font.Italic = 1
            $selection.Font.Bold = 0
            $selection.ParagraphFormat.LeftIndent = 36
            $selection.ParagraphFormat.SpaceBefore = 6
            $selection.ParagraphFormat.SpaceAfter = 6
            $cleanQuote = $trimmed.Substring(2).Replace("*", "")
            $selection.TypeText($cleanQuote)
            $selection.TypeParagraph()
            $selection.ParagraphFormat.LeftIndent = 0
            $selection.Font.Italic = 0
        }
        elseif ($trimmed -eq "---") {
            # Horizontal divider / spacing
            $selection.ParagraphFormat.SpaceBefore = 6
            $selection.ParagraphFormat.SpaceAfter = 6
            $selection.TypeParagraph()
        }
        elseif ($trimmed.StartsWith("|") -and $trimmed.EndsWith("|")) {
            if ($trimmed -match "^\|\s*:\s*-+") {
                # Separator line in markdown table, skip
                continue
            }
            $selection.Font.Name = "Times New Roman"
            $selection.Font.Size = 11
            $selection.Font.Bold = 0
            $selection.ParagraphFormat.LeftIndent = 18
            $selection.ParagraphFormat.SpaceBefore = 2
            $selection.ParagraphFormat.SpaceAfter = 2
            $cells = $trimmed.Trim('|').Split('|') | ForEach-Object { $_.Trim().Replace("**", "") }
            $tableLine = ($cells -join " : ")
            $selection.TypeText($tableLine)
            $selection.TypeParagraph()
            $selection.ParagraphFormat.LeftIndent = 0
        }
        elseif ($trimmed.Length -eq 0) {
            # Empty line
            continue
        }
        else {
            $selection.Font.Name = "Times New Roman"
            $selection.Font.Size = 12
            $selection.Font.Bold = 0
            $selection.Font.Italic = 0
            $selection.ParagraphFormat.SpaceBefore = 0
            $selection.ParagraphFormat.SpaceAfter = 6
            $selection.ParagraphFormat.LineSpacingRule = 0 # Single space
            
            # Simple bold/italic strip or clean text for Word output
            $cleanText = $trimmed.Replace("**", "").Replace("*", "")
            $selection.TypeText($cleanText)
            $selection.TypeParagraph()
        }
    }
    
    $doc.SaveAs([ref]$docxPath, [ref]16) # 16 = wdFormatDocumentDefault (.docx)
    Write-Output "Successfully generated: $docxPath"
}
finally {
    if ($doc) { $doc.Close([ref]$false) }
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
