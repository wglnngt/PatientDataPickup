Add-Type -AssemblyName System.Web;
$oADOX;
$oConn;
$oRst;
$nAgeMin = 0;
$nAgeMax = 0;
$strGenderSel = "";
$strHospitalDischargeStatus = "";
$strICD9Code = "";
$strExt = "";
$nIndex = 0;
$nIndexAim = 0;

function OpenDatabase($strFile) {
	if (!(Test-Path "$strFile.mdb")) {
		# Open database for using, is not exist, copy the tmplate file.
		#$script:oADOX = new-object -comobject ADOX.Catalog;
		#$script:oADOX.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$strFile.mdb");
		Copy-Item -path "template.mdb" -dest "$strFile.mdb";
	}

	$script:oConn = new-object -comobject ADODB.Connection;
	$script:oConn.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$strFile.mdb");
}

function OpenTable($strTable) {
	# If the table exist, drop it and create again.
	$strSQL = "DROP TABLE " + $strTable;
	try {
		$script:oConn.Execute($strSQL);
	} catch {
		write-host $_;
	}

	CreateTable($strTable);
}

function CreateTable($strTable) {
	$strSQL = "create table " + $strTable + " (";
	# Get the title
	$strTitle = sed -n "1p" "$strTable.txt"
	$strReg = ":a;
			   s/\([^\`"]\)\(\`"[^\`"\,]*\)\,\([^\`"]*\`"\)/\1\2\;\3/g;
			   ta;";
	$strTitle = echo $strTitle | sed $strReg;
	write-host "strTitle : $strTitle";
	$strTitle.split(',') | %{
		$strSQL += "$_ TEXT,";
	}
	$strSQL = $strSQL.trim(",");
	$strSQL += ");";
	write-host "strSQL : $strSQL";

	$script:oConn.Execute($strSQL);
}

function CloseDatabase() {
	#$script:oRst.close();
	$script:oConn.close();
}


function AddData($strTable) {
	$strSQL = "";
	
	## Read the first line, get the column number of patientunitstayid
	#$nIDColumn = sed -n "1p" "$strTable.txt" | sed "s/,/\n/g" | sed -n "/patientunitstayid/=";
	##write-host "patientunitstayid is col : $strIDColumn"

	$strHead = sed -n "1p" "$strTable.txt";
	$strHeads = $strHead.split(",");
	write-host "strHeads : $strHeads";
	# Change work flow, file content to memory, for fast info getting.[2018/06/21 KOUKON]
	$retIDS = cat _out-diag-ids_${script:strExt}.txt;
	#sed "1d" "$strTable.txt" | %{
	Import-CSV -Path "$strTable.txt" | %{
		#$strLine = $_;
		## ignore the blank line
		#$strLine = echo $strLine | sed 's/\"\"//g';
		#$strReg = ':a;
		#	s/\([^\"]\)\(\"[^\"\,]*\)\,\([^\"]*\"\)/\1\2\;\3/;
		#	ta';
		#$strLine = echo $strLine | sed $strReg;
		##write-host "Current data : $strLine"
		#$nID = GetColumnData ${strLine} ${nIDColumn}

		$script:nIndex++;
		$nID = $_.patientunitstayid;
		write-host "PatientUnitStayID : $nID"
		#$bTag = echo $retIDS | sed -n "/^${nID}`$/p";
		if ($retIDS.indexOf("${nID}") -gt -1) {
			#$bTag = echo $strLine | sed -n '/^\"/p';
			#if (!$bTag) {
			#	$strLine = '"' + $strLine;
			#}
			#$bTag = echo $strLine | sed -n '/\"$/p';
			#if (!$bTag) {
			#	$strLine = $strLine + '"';
			#}
			##write-host "strLine : $strLine";
			##read-host;

			## Add quote to strings
			#$strLine = echo $strLine | sed 's/\([^\",]\),/\1\",/g';
			#$strLine = echo $strLine | sed 's/\,\([^\",]\)/,\"\1/g';
			## null value
			#$strLine = echo $strLine | sed ':a;s/,,/,\"\",/;ta';

			$strLine = "";
			foreach ($i in $strHeads) {
				$strLine += '"' + $_.$i + '",';
			}
			$strLine = $strLine.trim(",");
			$strSQL = "insert into " + $strTable + " values(" + $strLine + ");";
			write-host "strSQL : $strSQL";

			$script:oConn.Execute($strSQL);
			$script:nIndexAim++;
		}
	}
}

function GetAgeRange() {
	cls;
	$ageRange = read-host "Please input the age(e.g. 0-20)";
	$bTag = echo $ageRange | sed -n "/[0-9]\{1,\}-[0-9]\{1,\}/p";
	write-host "`$bTag is $bTag";
	if ($bTag) {
		$script:nAgeMin = echo $ageRange | sed "s/\([0-9]*\)-.*/\1/";
		$script:nAgeMax = echo $ageRange | sed "s/.*-\([0-9]*\)/\1/";
	} else {
		write-host "Error input data, please retry later...";
		read-host;
		GetAgeRange;
	}
}

function GetGender() {
	echo "Please select the gender, M for male and F for famale and 0 for all:";
	$gender = read-host "Your choise";
	$bTag = echo $gender | sed -n "/^[FfMm0]$/p";
	if ($bTag) {
		if ($gender -eq "f" -or $gender -eq "F") {
			$script:strGenderSel = "Female";
		} elseif ($gender -eq "m" -or $gender -eq "M") {
			$script:strGenderSel = "Male";
		} else {
			$script:strGenderSel = "0";
		}
	} else {
		write-host "Error input data, please retry later...";
		read-host;
		GetGender;
	}
}

function GetHospitalDischargeStatus() {
	echo "Please select the hospital discharge status, E for Expired and A for Alive:";
	$hds = read-host "Your choise";
	$bTag = echo $hds | sed -n "/^[AaEe]$/p";
	if ($bTag) {
		if ($hds -eq "e" -or $hds -eq "E") {
			$script:strHospitalDischargeStatus = "Expired";
		} else {
			$script:strHospitalDischargeStatus = "Alive";
		}
	} else {
		write-host "Error input data, please retry later...";
		read-host;
		GetHospitalDischargeStatus;
	}
}

function GetICD9Code() {
	$strTmpCode = read-host "Please input the icd9code [e.g I46.9] ";
	$strTmpCode = echo $strTmpCode | sed "s/ //g";
	#$bTag = echo $strTmpCode | sed -n "/[^\,]*,[^\,]*/p";
	$bTag = echo $strTmpCode | sed -n "/[A-Z][0-9]*\..*/p";
	if ($bTag) {
		$script:strICD9Code = $strTmpCode;
	} else {
		write-host "Error input data, please retry later...";
		read-host;
		GetICD9Code;
	}
}

function GetColumnData($strLine, $nCol) {
	$nSkip = $nCol - 1;
	$strReg = "s/\([^,]*\,\)\{${nSkip}\}\([^,]*\)\,.*/\2/g";
	#write-host "==DEBUG==Raw data :`n$strLine";
	#write-host "==DEBUG==Col num : $nCol";
	#write-host "==DEBUG==Skip col : $nSkip";
	#write-host "==DEBUG==sed $strReg";
	$strValue = echo $strLine | sed $strReg;
	$strValue = echo $strValue | sed "s/\`"//g";
	return $strValue;
}

function Init() {
	GetAgeRange;
	#$script:nAgeMin = 0;
	#$script:nAgeMax = 60;
	write-host "The min age number is : $script:nAgeMin`nThe max age number is : $script:nAgeMax"

	GetGender;
	#$script:strGenderSel = "0";
	write-host "The gender you input is : $script:strGenderSel";

	GetHospitalDischargeStatus;
	#$script:strHospitalDischargeStatus = "Alive";
	write-host "The hospitaldischargestatus you input is : $script:strHospitalDischargeStatus";

	GetICD9Code
	#$script:strICD9Code = "456.0,I85.01";
	write-host "The icd9code you input is : $script:strICD9Code";

	$script:strExt  = "${script:nAgeMin}_${script:nAgeMax}_${script:strGenderSel}_";
	$script:strExt += "${script:strHospitalDischargeStatus}_${script:strICD9Code}_";
	$strDate		= Get-Date -Format 'yyyyMMdd';
	$script:strExt += $strDate
	$script:strExt  = ${script:strExt}.replace(".", "_").replace(" ", "_").replace(",", "_");

	write-host "The extentien string is : ${script:strExt}";
	#read-host;
	write-host "Now begin analysis...`n`n`n";

	$strQuery = "===============================`n";
	$strQuery += "Date : $strDate`n";
	$strQuery += "Age min : ${script:nAgeMin}`n";
	$strQuery += "Age max : ${script:nAgeMax}`n";
	$strQuery += "Gender : ${script:strGenderSel}`n";
	$strQuery += "HospitalDischargeStatue : ${script:strHospitalDischargeStatus}`n";
	$strQuery += "ICD9Code : ${script:strICD9Code}`n";
	$strQuery += "Result file : result_${script:strExt}.mdb`n";
	echo $strQuery | Out-File -Encoding ASCII -Append QueryHistoryLog.txt;
}

function Patient_analisys() {
	#$tmpTable = sed "=" .\patient.txt | sed "N;s/\n/,/" | sed "1d";
	#$tmpTable = sed "1d" .\patient.txt;

	$strIDs = "";
	$strLines = "";
	#echo $tmpTable | %{
	# Change the method to get CSV file content, by Import-CSV, for faster...
	Import-CSV -Path .\patient.txt | %{
		$script:nIndex++;
		#write-host "Current data :`n$_";
		#$strLine = $_;
		## ignore blank line
		#$strLine = echo $strLine | sed 's/\"\"//g';
		#$strReg = ':a;
		#		   s/\([^\"]\)\(\"[^\"\,]*\)\,\([^\"]*\"\)/\1\2\;\3/g;
		#		   ta';
		##write-host "sed $strReg";
		#$strLineDel = echo $strLine | sed $strReg;
		##write-host "Analysis data : $strLineDel"

		#$nID = GetColumnData $strLineDel 1;
		$nID = $_.patientunitstayid;
		write-host "PatientUnitStayID : $nID";

		#$nAge = GetColumnData $strLineDel 4;
		$nAge = $_.age;
		write-host "Age : $nAge";

		#$strGender = GetColumnData $strLineDel 3;
		$strGender = $_.gender;
		write-host "Gender : $strGender";

		#$strHDS = GetColumnData $strLineDel 20;
		$strHDS = $_.hospitaldischargestatus;
		write-host "Discharge status : $strHDS";
		
		#$strErr = GetColumnData $strLineDel 8;
		#write-host "Error-pron col num : $strErr";
		#write-host "`n`n";

		if ($nAge -ge ${script:nAgeMin} -and $nAge -le ${script:nAgeMax} `
			-and $strHDS -eq ${script:strHospitalDischargeStatus}) 
		{
			if (${script:strGenderSel} -eq "0" -or ${script:strGenderSel} -eq $strGender) {
				#$bTag = echo $strIDs | sed -n "/\;*${nID}\;*/p";
				if ($strIDs.indexOf("${nID};") -eq -1) {
					$strIDs += "${nID};";
					write-host "Meet the conditions, add into the inital screening...";
				}
				#$strLines += "$_`n";
			}
		}
		#read-host;
		write-host "================================================";
	}
	$strIDs = $strIDs.trim(";");
	#$strLines = echo $strLines | sed "`$d";
	write-host "The recordings meet the conditions post first screening :"
	echo $strIDs
	#echo "`n`n";
	#echo $strLines;

	# Result out
	echo $strIDs | sed "s/;/\n/g" | out-file -encoding ascii _out-ids_${script:strExt}.txt;
}

function Diagnosis_analysis() {
	$script:strICD9Code = echo $script:strICD9Code | sed "s/`,/\\`, /";
	#$script:strICD9Code = echo $script:strICD9Code | sed "s/\`./\\`./g";
	$strRegICD9Code = "/.*$script:strICD9Code.*/p";
	write-host "strRegICD9Code : ${strRegICD9Code}";
	# New work flow, if ICD9CODE exist, go on[2018/06/21 KOUKON]
	#$tmpTable = sed "1d" .\diagnosis.txt | sed -n "${strRegICD9Code}";

	# Using Import-Csv to get CSV file content, for faster[2018/06/21 KOUKON]
	sed -n "1p;${strRegICD9Code}" .\diagnosis.txt | Out-File -Encoding ascii _rst_${script:strExt}.txt;

	$strDiagIDs = "";
	$strDiagLines = "";
	# Push re-use content to memory, for faster[2018/06/21 KOUKON]
	$retIDS = cat _out-ids_${script:strExt}.txt;
	# CSV file reading, by Import-CSV
	Import-CSV -Path _rst_${script:strExt}.txt | %{
		$script:nIndex++;
		#write-host "Current data :`n$_";
		#$strLine = $_;
		## ignore blank line
		#$strLine = echo $strLine | sed 's/\"\"//g';
		#$strReg = ':a;
		#		   s/\([^\"]\)\(\"[^\"\,]*\)\,\([^\"]*\"\)/\1\2\;\3/g;
		#		   ta';
		##write-host "sed $strReg";
		#$strLineDel = echo $strLine | sed $strReg;
		##write-host "Analysis data : $strLineDel'

		$nID = $_.patientunitstayid;

		#$bTag = echo $retIDS | sed -n "/^${nID}`$/p";
		if ($retIDS.indexOf("${nID}") -gt -1) {
			write-host "Results meet the preliminary screening";
			write-host "PatientUnitStayID : $nID";
			#$bDiagTag = echo $strDiagIDs | sed -n "/\;*${nID}\;*/p";
			if ($strDiagIDs.indexOf("${nID};") -eq -1) {
				$strDiagIDs += "${nID};";
				write-host "Merge into the target library..."
			}
			#$strDiagLines += "$_`n";
		}
		write-host "================================================";
	};
	$strDiagIDs = $strDiagIDs.trim(";");
	write-host "Compliance line post the second screening : "
	echo $strDiagIDs;
	#echo `n`n;
	#echo $strDiagLines;
	# Echo result
	echo $strDiagIDs | sed "s/;/\n/g" | out-file -encoding ascii _out-diag-ids_${script:strExt}.txt;
}

function AddTable($strTable) {
	OpenDatabase "$pwd\result_${script:strExt}";
	OpenTable "$strTable";
	AddData "$strTable";
	CloseDatabase;
}

# Start here...

# Get the begin time...
$dateBegin = date;

Init;
Patient_analisys;
Diagnosis_analysis;
#copy-item -path _out-ids_${script:strExt}.txt  -dest _out-diag-ids_${script:strExt}.txt;

# If valid data exist, go on
$nCount = sed -n "`$=" _out-diag-ids_${script:strExt}.txt;
if ($nCount -ge 1) {
	if ($nCount -eq 1) {
		$bNullLine = sed -n "`$p" _out-diag-ids_${script:strExt}.txt;
		if ("$bNullLine" -eq "") {
			write-host "The result is only one line and is blank";
			$bTag = 0;
		} else {
			write-host "The result is only one row and valid data";
			$bTag = 1;
		}
	} else {
		write-host "There are multiple lines of results";
		$bTag = 1;
	}
} else {
	write-host "The result is less than one line and basically does not appear";
	$bTag = 0
}

if ($bTag) {
	if (test-path "$pwd\result_${script:strExt}.mdb") { rm "$pwd\result_${script:strExt}.mdb"; }
	ls "[a-z]*.txt" | %{
		AddTable "$($_.basename)";
	}
	echo "Mission status : Success...`n" | Out-File -Encoding ASCII -Append QueryHistoryLog.txt;
} else {
	write-host "No valid data, this query process terminated..."
	echo "Mission status : Failed...`n" | Out-File -Encoding ASCII -Append QueryHistoryLog.txt;
}

# Get the end time...
$dateEnd = date;

#read-host;
#exit;
$strDura  = ($dateEnd - $dateBegin).TotalHours;
$strShow  = "Total result :  $nIndex rows`n";
$strShow += "Aim result :  $nIndexAim rows`n";
$strShow += "Start at : $dateBegin`n";
$strShow += "End at : $dateEnd`n";
$strShow += "Total use : $strDura hours`n";
echo $strShow;
echo $strShow | Out-File -Encoding ASCII -Append QueryHistoryLog.txt;
read-host "Press Enter to exit..."
