#Start-transcript -path .\PowershellTranscript
##################################################################################
#          1. Variable declaration
##################################################################################
#Log file
$date=[string](date).Year+"_"+[string]("{0:D2}" -f (date).Month)+"_"+[string]("{0:D2}" -f (date).Day)
$date2=[string]((date).Year)+[string]("{0:D2}" -f ((date).Month))+[string]("{0:D2}" -f (date).Day)
$timestamp= $date+"_"+[string](date).Hour+"_"+[string](date).Minute+"_"+[string](date).Second
$logFile = "Logfile_$date.txt"
   ni $logFile -Type file -Force

   "****************************LOG-BEGINS*************************************** ">>$logFile

#File search pattern eg."6.4.a_HardwareToken_user_ALL_LDAP_NOT_LDAPold_NOT_RSADB_2015-06-30---03-00-01.csv"
$filePattern_a="6.4.a*.csv"
#File search pattern eg."6.4.b_HardwareToken_user_ALL_LDAP_NOT_LDAPold_NOT_RSADB_2015-06-30---03-00-01.csv"
$filePattern_b="6.4.b*.csv"
$filePattern_10admin="10.admin*.csv"
$filePattern_RSA_ext="HardwareToken*Active_External*.csv"
$filePattern_RSA_int="HardwareToken*Active_Internal*.csv"
$filePattern_RSA_all="HardwareToken*.csv"

#Feilds that are present along with data in csv files
$feildIdentifier_a="mail: ", "givenName: ", "sn: ", "supplierName: ", "employeeType: ", "departmentNumber: "
$feildIdentifier_b= "givenName: ", "sn: ", "supplierName: ", "employeeType: ", "departmentNumber: "

#Order of the columns in csv files
$csvHeaders_a="userid", "email", "firstName", "lastName", "company", "departmentNumber"
$csvHeaders_b="userid", "firstName", "lastName", "company", "departmentNumber"
$csvHeaders_FinalList = "userid", "Vorname", "Nachname", "email", "company", "KurzZ"

#Delimiter used in the files
$delimiter=";"
$delimiter_RSA=';'

#Final name of the processed file
$outputIntFile="1_Provisioning_Internal_withoutemail.csv"
$outputExtFile="2_Provisioning_External_withouremail.csv"
$outputAdminFile="3_Provisioning_all_withemail.csv"
$outputCollated="4_Provisioning_Collated_$date.csv"
$outputRSA_DB="5_RSA_Users_with_Tokens.csv"
$outputFinalreport="SecurID_Shipping_$date2.csv"

##################################################################################
#         2. Collating 6.4.a files
##################################################################################

#local variables

$filePattern=$filePattern_a
$csvHeaders=$csvHeaders_a
$feildIdentifier=$feildIdentifier_a
$outputFile=$outputIntFile
$tempFile="temp"
$inputFile="4a_test.csv"

#collating all the files with given pattern
cat $filePattern > .\$inputFile
dir $filePattern -Recurse | %{ $_ | select  name, @{n="lines";e={gc $_ | Measure-Object -line | select -expa lines}}} | ft -AutoSize -wrap  >>$logFile
Write-Host "`nCollated files with pattern $filePattern and total entries are "(gc $inputFile).Length"."
"Collated the above files and total entries are "+(gc $inputFile).Length +"." >>$logFile

#Importing csv file and choosing required coloumns in specific order
$cs =Import-Csv .\$inputFile -Delimiter $delimiter -Header $csvHeaders|select $csvHeaders[0],$csvHeaders[2],$csvHeaders[3],$csvHeaders[1], $csvHeaders[4], $csvHeaders[5] | Export-Csv .\$outputFile -Delimiter ';'

#Removing the feild identifiers from the file
for ($i=0; $i -lt $feildIdentifier_a.Length; $i++){

 $nowFile=[string]($i)+$tempFile
 $nextFile=[string]($i+1)+$tempFile
 #Write-Host $nowFile
 #Write-Host $nextfile

	if($i -eq 0){
		copy $outputFile $nowFile
		}
	(gc $nowFile) -replace $feildIdentifier[$i],'' >.\$nextFile
    #Write-Host "`nRemoved "$feildIdentifier[$i]" from the file." | tee $logFile -Append
	
}

del $outputFile
ren $nextFile $outputFile
del *$tempFile

#Removing the headers for DB import
[int]$len=([convert]::ToInt32((gc $outputFile).Length))-2
#Write-Host $len
(gc $outputFile) | select -Last $len >.\$tempFile
if(!(gc $tempFile).Length -eq 0){

		del $outputFile
		(gc $tempFile) -replace '"','' >.\$outputFile
		del $tempFile
				}

#Deleting the temporary files
if((gc $inputFile).Length -eq (gc $outputFile).Length){
        del $inputFile
                           
                }
Write-Host "`nCreated $outputFile !"
		




##################################################################################
#         3. Collating 6.4.b files
##################################################################################


#local variables
$filePattern=$filePattern_b
$csvHeaders=$csvHeaders_b
$feildIdentifier=$feildIdentifier_b
$outputFile=$outputExtFile
$tempFile="temp"
$inputFile="4b_test.csv"

#collating all the files with given pattern
cat $filePattern > .\$inputFile
dir $filePattern -Recurse | %{ $_ | select  name, @{n="lines";e={gc $_ | Measure-Object -line | select -expa lines}}} | ft -wrap -AutoSize >>$logFile
Write-Host "`nCollated files with pattern $filePattern and total enteries are "(gc $inputFile).Length"."
"Collated the above files and total entries are "+(gc $inputFile).Length +"." >>$logFile

#Importing csv file and choosing required coloumns in specific order
#NUL feild is choosen as no email data is present in csv file
$cs =Import-Csv .\$inputFile -Delimiter $delimiter -Header $csvHeaders|select $csvHeaders[0], $csvHeaders[1],$csvHeaders[2], NUL, $csvHeaders[3], $csvHeaders[4] | Export-Csv .\$outputFile -Delimiter ';'

#Removing the feild identifiers from the file
for ($i=0; $i -lt $feildIdentifier.Length; $i++){

$nowFile=[string]($i)+$tempFile
$nextFile=[string]($i+1)+$tempFile

	if($i -eq 0){
		copy $outputFile $nowFile
		}
	(gc $nowFile) -replace $feildIdentifier[$i],'' >.\$nextFile
  #Write-Host "`nRemoved "$feildIdentifier[$i]" from the file." | tee $logFile -Append
	
}

del $outputFile
ren $nextFile $outputFile
del *$tempFile

#Removing the headers for DB import
[int]$len=([convert]::ToInt32((gc $outputFile).Length))-2
#Write-Host $len

(gc $outputFile)| select -Last $len >.\$tempFile
if(!(gc $tempFile).Length -eq 0){

		del $outputFile
		(gc $tempFile) -replace '"','' >.\$outputFile
		del $tempFile
				}

#Deleting the temporary files
if((gc $inputFile).Length -eq (gc $outputFile).Length){
        del $inputFile
                           
                }
Write-Host "`nCreated $outputFile !"
		
del 8_role*.csv

##################################################################################
#         4. Collating 10.admin, 6_4_a, 6_4_b files
##################################################################################

#local variables
$filePattern=$filePattern_10admin
$outputFile=$outputCollated
$inputFile="10_test.csv"

#collating all the files with given pattern
cat $filePattern > .\$outputAdminFile
dir $filePattern -Recurse | %{ $_ | select  name, @{n="lines";e={gc $_ | Measure-Object -line | select -expa lines}}} | ft -wrap -AutoSize >>$logFile
Write-Host "`nCollated files with pattern "$filePattern" and total enteries are "(gc $outputAdminFile).Length"."
"Collated the above files and total entries are "+(gc $outputAdminFile).Length +"." >>$logFile 

cat $outputIntFile,$outputExtFile,$outputAdminFile | select -unique >.\$outputFile.temp
copy .\$outputFile.temp $outputFile -Force
del *.temp				
del *test.csv
#copy .\$outputCollated ..\$outputCollated -force
Write-Host "`nCreated $outputFile !      -->   "(gc $outputCollated).length"."

"`r`nCreated "+$outputFile+" !      -->   "+(gc $outputCollated).length+"." >>$logFile

##################################################################################
#         5. Collating RSA DB reports
##################################################################################

#local variables
$filePattern=$filePattern_RSA_ext
$outputFile=$outputRSA_DB

if((ls $filePattern).count -gt 1){
            Write-Host "`nMore than one file exists with $filePattern, selecting the recently modified file "
            "`r`nMore than one file exists with "+$filePattern+", selecting the recently modified file `r`n">>$logFile
            $temp = ls $filePattern | sort LastWriteTime | select -Last 1 #>.\temp
            $filePattern_RSA_ext =(( $temp.name) )#| select -Last 3 | select -First 1)
            #del .\temp
            #
            Write-Host (ls $filePattern_RSA_ext).Name
            (ls $filePattern_RSA_ext).Name >>$logFile
            }
$filePattern=$filePattern_RSA_int
if((ls $filePattern).count -gt 1){
             Write-Host "`nMore than one file exists with $filePattern, selecting the recently modified file "
            "`r`nMore than one file exists with "+$filePattern+", selecting the recently modified file `r`n">>$logFile
            $temp = ls $filePattern | sort LastWriteTime | select -Last 1 #>.\temp
            $filePattern_RSA_int =(( $temp.name) )#| select -Last 3 | select -First 1)
            #del .\temp
            Write-Host (ls $filePattern_RSA_int).Name
            (ls $filePattern_RSA_int).Name >>$logFile
            }
Write-Host "`nGenerating $outputFile, please wait . . . "

((cat $filePattern_RSA_ext,$filePattern_RSA_int) -replace ';.*','')  > .\$outputFile

Write-Host "`nRSA DB contains "(gc $outputFile).Length " user(s) with tokens." 

"`r`nRSA DB contains "+(gc $outputFile).Length +" user(s) with tokens." >> $logFile 

Write-Host "`nCreated $outputFile !"


##################################################################################
#         6. Comparing Provisioning list with RSA DB and creating ad-hoc report
##################################################################################

#local variables
$count=0
$validusers="temp.csv"
$csvHeaders=$csvHeaders_FinalList


$provList= Import-Csv $outputCollated -Delimiter $delimiter -Header $csvHeaders
$rsa= Import-Csv $outputRSA_DB -Header $csvHeaders[0]

Write-Host "`nComparing $outputCollated with $outputRSA_DB, please wait . . ."
$validList = diff $rsa $provList -Property $csvHeaders[0] | where {$_.SideIndicator -eq "=>"} 
$validList | select $csvHeaders[0] | Out-File $validusers
Write-Host "`nValid user's list is being processed" 
( gc $validusers) -replace ' ','' >temp
 gc temp | select -Last (((gc temp).count)-3) >$validusers
 del temp
Write-Host "`nValid user's list is processed" 

$validuserlength= ( [convert]::ToInt32((gc $validusers).Length)-2)
Write-Host "`n$validuserlength user(s) are listed out for today's provisioning`n"

"$validuserlength user(s) are listed out for today's provisioning" >>$logFile
$provListlength =($provList.Length)

$outputFile = $provList | where  {
$count = $count +1
Write-Progress "Generating ad-hoc report ... " -Status $count/$provListlength -PercentComplete ($count/($provListlength)*100) 
#the match value '$_.USERID" in the below line is hardcoded. This has to be changed if any changes made in csvHeaders
 (gc $validusers ) -match $_.userid}
 del $validusers

 $outputFile | sort $csvHeaders[5] | Export-Csv $outputFinalreport -NoTypeInformation
 $outputFile | Out-GridView

 
##################################################################################
#         7. Cleaning residual files
##################################################################################
ni "./Residual files" -type directory -force

move $filePattern_a,$filePattern_b,$filePattern_10admin,$filePattern_RSA_all "./Residual files"

" Moved $filePattern_a,$filePattern_b,$filePattern_10admin,$filePattern_RSA_all ">>$logFile

"*****************************LOG-ENDS**************************************** ">>$logFile
#Stop-Transcript