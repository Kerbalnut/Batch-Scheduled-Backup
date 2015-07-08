
# Thanks to:
# http://www.chapmancentral.co.uk/cloudy/2013/02/23/parsing-robocopy-logs-in-powershell/
# for this script!

# ^^ All of this is his work, I merely edited the output naming.

# So basically what I did, without any knowledge of powershell, is add another parameter (a required one actually!) that allows you to specify the output .csv's File Name and Path. 
# Keep in mind I basically just copied & paste'd what he already had there and changed some variable names.
# That way, this whole script becomes way more useful for scripting it from other scripts.
# Now you can call this script from another, specify where you want the output to go and for what it should be called, then after this runs you can easily collect the .csv output in the other language to parse the results you need. 
# ('cause you already know where it is and what it's called, instead of searching for whatever it was named by this script and hope you get it right)
# What I did is add 2 new variables, $OutputPath and $outputfile (I dunno if I did that right)
# Either way, it works. You cannot include the "*.csv" when specifying the output path.
# The script will include that for you. Only include the File Path starting from the letter, and the name of the file, no file format.
# For File Paths with spaces, surround Path in single quotes: e.g.

# PowerShell .\Parse-RobocopyLogs.ps1 -fp '%BAKUPLOGPATH%\last-backup-%DRIVETOBBLETTER%.log' -outputfile '%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results' > "%Temp%\RoboParseOutput.log"
# del "%Temp%\RoboParseOutput.log"

# This (batch) example also redirects *ALL* output to a "RoboParseOutput.log" so nothing shows onscreen. This can be modified (remove the >) so it will show the output messages from this script when you run it within another script.


param(
	[parameter(Position=0,Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Source Path with no trailing slash')][string]$SourcePath,
	[parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Output File Name and Path with no .csv')][string]$OutputPath,
	[switch]$fp,[switch]$outputfile
	)
 
write-host "Robocopy log parser. $(if($fp){"Parsing file entries"} else {"Parsing summaries only, use -fp to parse file entries"})"
 
#Arguments
# -fp			File parse. Counts status flags and oldest file Slower on big files.
# -outputfile	Specify output file Name and Path
 
$ElapsedTime = [System.Diagnostics.Stopwatch]::StartNew()
$refreshrate=1 # progress counter refreshes this often when parsing files (in seconds)
 
# These summary fields always appear in this order in a robocopy log
$HeaderParams = @{
	"04|Started" = "date";	
	"01|Source" = "string";
	"02|Dest" = "string";
	"03|Options" = "string";
	"07|Dirs" = "counts";
	"08|Files" = "counts";
	"09|Bytes" = "counts";
	"10|Times" = "counts";
	"05|Ended" = "date"
	#"06|Duration" = "string"
}
 
$ProcessCounts = @{
	"Processed" = 0;
	"Error" = 0;
	"Incomplete" = 0
}
 
$tab=[char]9
 
$files=get-childitem $SourcePath
 
# Original that is now commented out, my edits are the two lines below.
# $writer=new-object System.IO.StreamWriter("$(get-location)\robocopy-$(get-date -format "dd-MM-yyyy_HH-mm-ss").csv")

IF($outputfile){}ELSE{$outputfile=$OutputPath}
$writer=new-object System.IO.StreamWriter("$("$OutputPath").csv")
 
function Get-Tail([object]$reader, [int]$count = 10) {
 
	$lineCount = 0
	[long]$pos = $reader.BaseStream.Length - 1
 
	while($pos -gt 0)
	{
		$reader.BaseStream.position=$pos
 
		# 0x0D (#13) = CR
		# 0x0A (#10) = LF
		if ($reader.BaseStream.ReadByte() -eq 10)
		{
			$lineCount++
			if ($lineCount -ge $count) { break }
		}
		$pos--
	} 
 
	# tests for file shorter than requested tail
	if ($lineCount -lt $count -or $pos -ge $reader.BaseStream.Length - 1) {
		$reader.BaseStream.Position=0
	} else {
		# $reader.BaseStream.Position = $pos+1
	}
 
	$lines=@()
	while(!$reader.EndOfStream) {
		$lines += $reader.ReadLine()
	}
	return $lines
}
 
function Get-Top([object]$reader, [int]$count = 10)
{
	$lines=@()
	$lineCount = 0
	$reader.BaseStream.Position=0
	while(($linecount -lt $count) -and !$reader.EndOfStream) {
		$lineCount++
		$lines += $reader.ReadLine()		
	}
	return $lines
}
 
function RemoveKey ( $name ) {
	if ( $name -match "|") {
		return $name.split("|")[1]
	} else {
		return ( $name )
	}
}
 
function GetValue ( $line, $variable ) {
 
	if ($line -like "*$variable*" -and $line -like "* : *" ) {
		$result = $line.substring( $line.IndexOf(":")+1 )
		return $result 
	} else {
		return $null
	}
}
 
function UnBodgeDate ( $dt ) {
	# Fixes RoboCopy botched date-times in format Sat Feb 16 00:16:49 2013
	if ( $dt -match ".{3} .{3} \d{2} \d{2}:\d{2}:\d{2} \d{4}" ) {
		$dt=$dt.split(" ")
		$dt=$dt[2],$dt[1],$dt[4],$dt[3]
		$dt -join " "
	}
	if ( $dt -as [DateTime] ) {
		return $dt.ToStr("dd/MM/yyyy hh:mm:ss")
	} else {
		return $null
	}
}
 
function UnpackParams ($params ) {
	# Unpacks file count bloc in the format
	#    Dirs :      1827         0      1827         0         0         0
	#	Files :      9791         0      9791         0         0         0
	#	Bytes :  165.24 m         0  165.24 m         0         0         0
	#	Times :   1:11:23   0:00:00                       0:00:00   1:11:23
	# Parameter name already removed
 
	if ( $params.length -ge 58 ) {
		$params = $params.ToCharArray()
		$result=(0..5)
		for ( $i = 0; $i -le 5; $i++ ) {
			$result[$i]=$($params[$($i*10 + 1) .. $($i*10 + 9)] -join "").trim()
		}
		$result=$result -join ","
	} else {
		$result = ",,,,,"
	}
	return $result
}
 
$sourcecount = 0
$targetcount = 1
 
# Write the header line
$writer.Write("File")
foreach ( $HeaderParam in $HeaderParams.GetEnumerator() | Sort-Object Name ) {
	if ( $HeaderParam.value -eq "counts" ) {
		$tmp="~ Total,~ Copied,~ Skipped,~ Mismatch,~ Failed,~ Extras"
		$tmp=$tmp.replace("~","$(removekey $headerparam.name)")
		$writer.write(",$($tmp)")
	} else {
		$writer.write(",$(removekey $HeaderParam.name)")
	}
}
 
if($fp){
	$writer.write(",Scanned,Newest,Summary")
}
 
$writer.WriteLine()
 
$filecount=0
 
# Enumerate the files
foreach ($file in $files) {  
	$filecount++
    write-host "$filecount/$($files.count) $($file.name) ($($file.length) bytes)"
	$results=@{}
 
	$Stream = $file.Open([System.IO.FileMode]::Open, 
                   [System.IO.FileAccess]::Read, 
                    [System.IO.FileShare]::ReadWrite) 
	$reader = New-Object System.IO.StreamReader($Stream) 
	#$filestream=new-object -typename System.IO.StreamReader -argumentlist $file, $true, [System.IO.FileAccess]::Read
 
	$HeaderFooter = Get-Top $reader 16
 
	if ( $HeaderFooter -match "ROBOCOPY     ::     Robust File Copy for Windows" ) {
		if ( $HeaderFooter -match "Files : " ) {
			$HeaderFooter = $HeaderFooter -notmatch "Files : "
		}
 
		[long]$ReaderEndHeader=$reader.BaseStream.position
 
		$Footer = Get-Tail $reader 16
 
		$ErrorFooter = $Footer -match "ERROR \d \(0x000000\d\d\) Accessing Source Directory"
		if ($ErrorFooter) {
			$ProcessCounts["Error"]++
			write-host -foregroundcolor red "`t $ErrorFooter"
		} elseif ( $footer -match "---------------" ) {
			$ProcessCounts["Processed"]++
			$i=$Footer.count
			while ( !($Footer[$i] -like "*----------------------*") -or $i -lt 1 ) { $i-- }
			$Footer=$Footer[$i..$Footer.Count]
			$HeaderFooter+=$Footer
		} else {
			$ProcessCounts["Incomplete"]++
			write-host -foregroundcolor yellow "`t Log file $file is missing the footer and may be incomplete"
		}
 
		foreach ( $HeaderParam in $headerparams.GetEnumerator() | Sort-Object Name ) {
			$name = "$(removekey $HeaderParam.Name)"
			$tmp = GetValue $($HeaderFooter -match "$name : ") $name
			if ( $tmp -ne "" -and $tmp -ne $null ) {
				switch ( $HeaderParam.value ) {
					"date" { $results[$name]=UnBodgeDate $tmp.trim() }
					"counts" { $results[$name]=UnpackParams $tmp }
					"string" { $results[$name] = """$($tmp.trim())""" }		
					default { $results[$name] = $tmp.trim() }		
				}
			}
		}
 
		if ( $fp ) {
			write-host "Parsing $($reader.BaseStream.Length) bytes" -NoNewLine
 
			# Now go through the file line by line
			$reader.BaseStream.Position=0
			$filesdone = $false
			$linenumber=0
			$FileResults=@{}
			$newest=[datetime]"1/1/1900"
			$linecount++
			$firsttick=$elapsedtime.elapsed.TotalSeconds
			$tick=$firsttick+$refreshrate
			$LastLineLength=1
 
			try {
				do {
					$line = $reader.ReadLine()
					$linenumber++
					if (($line -eq "-------------------------------------------------------------------------------" -and $linenumber -gt 16)  ) { 
						# line is end of job
						$filesdone=$true
					} elseif ($linenumber -gt 16 -and $line -gt "" ) {
						$buckets=$line.split($tab)
 
						# this test will pass if the line is a file, fail if a directory
						if ( $buckets.count -gt 3 ) {
							$status=$buckets[1].trim()
							$FileResults["$status"]++
 
							$SizeDateTime=$buckets[3].trim()
							if ($sizedatetime.length -gt 19 ) {
								$DateTime = $sizedatetime.substring($sizedatetime.length -19)
								if ( $DateTime -as [DateTime] ){
									$DateTimeValue=[datetime]$DateTime
									if ( $DateTimeValue -gt $newest ) { $newest = $DateTimeValue }
								}
							}
						}
					}
 
					if ( $elapsedtime.elapsed.TotalSeconds -gt $tick ) {
						$line=$line.Trim()
						if ( $line.Length -gt 48 ) {
							$line="[...]"+$line.substring($line.Length-48)
						}
						$line="$([char]13)Parsing > $($linenumber) ($(($reader.BaseStream.Position/$reader.BaseStream.length).tostring("P1"))) - $line"
						write-host $line.PadRight($LastLineLength) -NoNewLine
						$LastLineLength = $line.length
						$tick=$tick+$refreshrate						
					}
 
				} until ($filesdone -or $reader.endofstream)
			}
			finally {
				$reader.Close()
			}
 
			$line=$($([string][char]13)).padright($lastlinelength)+$([char]13)
			write-host $line -NoNewLine
		}
 
		$writer.Write("`"$file`"")
		foreach ( $HeaderParam in $HeaderParams.GetEnumerator() | Sort-Object Name ) {
			$name = "$(removekey $HeaderParam.Name)"
			if ( $results[$name] ) {
				$writer.Write(",$($results[$name])")
			} else {
				if ( $ErrorFooter ) {
					#placeholder
				} elseif ( $HeaderParam.Value -eq "counts" ) {
					$writer.Write(",,,,,,") 
				} else {
					$writer.Write(",") 
				}
			}
		}
 
		if ( $ErrorFooter ) {
			$tmp = $($ErrorFooter -join "").substring(20)
			$tmp=$tmp.substring(0,$tmp.indexof(")")+1)+","+$tmp
			$writer.write(",,$tmp")
		} elseif ( $fp ) {
			$writer.write(",$LineCount,$($newest.ToString('dd/MM/yyyy hh:mm:ss'))")			
			foreach ( $FileResult in $FileResults.GetEnumerator() ) {
				$writer.write(",$($FileResult.Name): $($FileResult.Value);")
			}
		}
 
		$writer.WriteLine()
 
	} else {
		write-host -foregroundcolor darkgray "$($file.name) is not recognised as a RoboCopy log file"
	}
}
 
write-host "$filecount files scanned in $($elapsedtime.elapsed.tostring()), $($ProcessCounts["Processed"]) complete, $($ProcessCounts["Error"]) have errors, $($ProcessCounts["Incomplete"]) incomplete"
write-host  "Results written to $($writer.basestream.name)"
$writer.close()
