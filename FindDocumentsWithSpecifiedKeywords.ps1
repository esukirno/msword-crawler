$comments = @' 
Script name: FindDocumentsWithSpecifiedKeywords.ps1 
Created on: Tuesday, August 08, 2014 
Author: Edwin Sukirno
Purpose: Find Keywords inside MS Word documents. For resumes crawling perhaps? 

Adaptation of http://gallery.technet.microsoft.com/scriptcenter/7c463ad7-0eed-4792-8236-38434f891e0e
Reference http://blogs.technet.com/b/heyscriptingguy/archive/2009/05/14/how-can-i-use-windows-powershell-to-look-for-and-replace-a-word-in-a-microsoft-word-document.aspx
'@ 
 
$searchDirectory = "C:\dev.git\MSWord.Crawler\"
$keywords = $("Drupal","NServiceBus","MySql","MVC")

$files = Get-ChildItem $searchDirectory | Where-Object { $_.Name -like "*.docx" } | Select-Object Name 

if ($files.length -eq 0) {
	Write-Host "Could not find *.docx file"
}

$objWord = New-Object -comobject Word.Application  
$objWord.Visible = $True 

foreach ($file in $files) {
	$filePath = $searchDirectory + $file.Name
	$objDoc = $objWord.Documents.Open($filePath) 	
	 	
	$MatchCase = $False 
	$MatchWholeWord = $False 
	$MatchWildcards = $False 
	$MatchSoundsLike = $False 
	$MatchAllWordForms = $False 
	$Forward = $True 
	$Wrap = 1 # WdFindWrap enumeration in MS Word interop
	$Format = $False 
	$Replace = 0 # WdReplace enumeration in MS Word interop
	$ReplaceText = "" 	
	 
	Write-Host "Start Crawling " $filePath
	$count = 0
	foreach ($text in $keywords) {		
		$objSelection = $objWord.Selection 
		$a = $objSelection.Find.Execute($text,$MatchCase,$MatchWholeWord, `
			$MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
			$Wrap,$Format,$ReplaceText,$Replace) 
		If ($a -eq $True) 
		{
			Write-Host `t $text "was found" -foregroundcolor "green"
			$count = $count + 1
		}
		Else {
			Write-Host `t $text "was not found" -foregroundcolor "red"
		}
	}
	Write-Host "Found $count out of" $keywords.length `n
	$objDoc.Close()
}

$objWord.Visible = $False 
$objWord.Quit()