$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

trap {
    $ErrorActionPreference = "Continue";   
    Write-Error $_
    exit 1
}

function ExportVba {
    param($book, $exportPath)

    $book.VBProject.VBComponents | % {
        $comp = $_
        Write-Host $_.Name
        switch ($comp.Type) {
            1 {
                $ext = ".bas"
            }
            2 {
                $ext = ".cls"
            }
            3 {
                $ext = ".frm"
            }
            100 {
                if ($comp.CodeModule.CountOfDeclarationLines -eq $comp.CodeModule.CountOfLines) {
                    # 宣言行と同じ行数しかないソースは無視する
                    return
                }
                $ext = ".bas"
            }
            default {
                throw "unknown type : $($comp.Type)"
            }
        }
        $filename = $book.Name -replace "\..*$", ""
        $filename += "_" + $comp.Name + $ext
        [void]$comp.Export( (Join-Path $exportPath $filename ) )
    }

}

function Release-ComObject {
    param ($obj)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj)
}

function main {
    $excel = New-Object -ComObject Excel.Application
    try {

        Get-ChildItem *.xlsm | % {
            $book = $excel.Workbooks.Open($_.FullName)
            try {
                $exportPath = Join-Path $book.Path "src"
                mkdir $exportPath -ErrorAction SilentlyContinue | Out-Null
                ExportVba $book $exportPath
            } finally {
                $book.Saved = $true
                [void]$book.Close()
            }
        }
    } finally {
        [void]$excel.Quit()
        Release-ComObject $excel
    }
}

main
