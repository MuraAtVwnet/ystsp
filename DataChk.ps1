Param( $DepartmentCSV, $EmployeeCSV )

# $DataPath = "D:\Sclipts\kiku\vone\Workflow\ワークフロー\人事組織管理システム\文書管理システム対応データインポート\Retry Data"
$DataPath = $PSScriptRoot


if( $DepartmentCSV -eq $null ){
	$DepartmentCSV = Join-Path $DataPath "組織情報.csv"
}

if( -not (Test-Path $DepartmentCSV)){
	echo "$DepartmentCSV not found"
	exit
}


if( $EmployeeCSV -eq $null ){
	$EmployeeCSV = Join-Path $DataPath "employee_data_20220311_03.csv"
}


if( -not (Test-Path $EmployeeCSV)){
	echo "$EmployeeCSV not found"
	exit
}

$employee = Import-Csv $EmployeeCSV
$department = Import-Csv $DepartmentCSV

# $NoHit = Import-Csv "D:\Sclipts\kiku\vone\Workflow\ワークフロー\人事組織管理システム\文書管理システム対応データインポート\Retry Data\組織割り当てのない従業員一覧.csv"

$ErrorDepartmentInBusyokanriDatas = $department

$ErrStrigngs = @()
$NormalStrigngs = @()

# 整合性確認
foreach( $ErrorDepartmentInBusyokanriData in $ErrorDepartmentInBusyokanriDatas ){

	$BysyoSoshikiName = $ErrorDepartmentInBusyokanriData.'H2事業所名'

	if( $ErrorDepartmentInBusyokanriData.CD -eq [string]$null ){
		continue
	}

	# 上位組織を検索
	[String] $UperSoshikiString = $ErrorDepartmentInBusyokanriData."上位部門CD"
#	$UperSoshiki = ([int]$ErrorDepartmentInBusyokanriData."上位部門CD").ToString()
	$UperSoshiki = $UperSoshikiString
	[array]$TergetUpSoshikis = $department | ? CD -eq $UperSoshiki
	$Count3 = $TergetUpSoshikis.Count
	$Code01 = $ErrorDepartmentInBusyokanriData.CD
	$Name01 = $ErrorDepartmentInBusyokanriData.'部署名'
	$Code02 = $UperSoshiki
	$H2Name = $BysyoSoshikiName
	if( $UperSoshikiString -eq [string]$null ){
			echo "上位組織がセットされていない $Code01 $Name01 H2 $H2Name"
			$ErrStrigngs += "上位組織がセットされていない $Code01 $Name01 H2 $H2Name"
	}
	elseif( $Count3 -eq 0 ){
		if( $Code02 -ne [string]$null ){
			echo "上位組織が存在しない $Code01 $Name01 H2 $H2Name : 上位部門コード $Code02"
			$ErrStrigngs += "上位組織が存在しない $Code01 $Name01 H2 $H2Name : 上位部門コード $Code02"
		}
	}
	elseif( $Count3 -ne 1 ){
		$Code03 = $UperSoshiki
		echo "上位組織が複数ある : 上位部門コード $Code03 : $Count3 "
		$ErrStrigngs += "上位組織が複数ある : 上位部門コード $Code03 : $Count3 "
		# $TergetUpSoshikis | select CD, "部署名"
	}
	else{
		$JigyushoCD = $ErrorDepartmentInBusyokanriData.CD
		$JigyushoName = $ErrorDepartmentInBusyokanriData.'部署名'
		$H2JigyushoName = $BysyoSoshikiName
		$JyouiJigyushoCD = $TergetUpSoshikis[0].CD
		$JyouiJigyushoName = $TergetUpSoshikis[0].'部署名'
		#echo "★★★ 正常データ : 事業所 $JigyushoCD $JigyushoName : H2事業所名 $H2JigyushoName : 上位部署 $JyouiJigyushoCD $JyouiJigyushoName"
		$NormalStrigngs += "★★★ 正常データ : 事業所 $JigyushoCD $JigyushoName : H2事業所名 $H2JigyushoName : 上位部署 $JyouiJigyushoCD $JyouiJigyushoName"
	}
}


# 部署単純重複確認
$ErrorDepartmentInBusyokanriDatas = $department

foreach( $ErrorDepartmentInBusyokanriData in $ErrorDepartmentInBusyokanriDatas ){
	$SoshikiCD = $ErrorDepartmentInBusyokanriData.CD

	if( $SoshikiCD -eq [string]$null ){
		continue
	}

	if($SoshikiCD.Length % 2 -ne 0){
		$ErrorMessage = "組織コード「$SoshikiCD」は、リーディング 0 が欠落している"
		echo $ErrorMessage
		$ErrStrigngs += $ErrorMessage
	}

	# 人事組織システムの組織コードで検索
	[array]$TergetSoshikisByCD = $department | ? CD -eq $SoshikiCD
	$Count2 = $TergetSoshikisByCD.Count
	$Code001 = $ErrorDepartmentInBusyokanriData.CD
	$Name001 = $ErrorDepartmentInBusyokanriData.'部署名'
	$H2Name = $ErrorDepartmentInBusyokanriData.'H2事業所名'
	if( $Count2 -eq 0 ){
		echo "組織コードが存在しない : $Code001"
		$ErrStrigngs += "組織コードが存在しない : $Code001"
	}
	else{
		if( $Count2 -ne 1 ){
			echo "組織コード「$Code001」が複数ある $Name001 H2組織名 $H2Name 重複数 : $Count2"
			$ErrStrigngs += "組織コード「$Code001」が複数ある $Name001 H2組織名 $H2Name 重複数 : $Count2"
			#echo $TergetSoshikisByCD | select CD, 部署名, H2事業所名
		}
	}
}


# 5階層のみ「ひいらぎコード」と「H2事業所名」がセットされている確認
$ErrorDepartmentInBusyokanriDatas = $department

foreach( $ErrorDepartmentInBusyokanriData in $ErrorDepartmentInBusyokanriDatas ){
	$SoshikiCD = $ErrorDepartmentInBusyokanriData.CD

	if( $SoshikiCD -eq [string]$null ){
		continue
	}

	$BuShoName = $ErrorDepartmentInBusyokanriData.'部署名'
	$Kaiso = $ErrorDepartmentInBusyokanriData.'階層'
	$HiragiCD = $ErrorDepartmentInBusyokanriData.'ひいらぎコード'
	$H2JigyoushoName = $ErrorDepartmentInBusyokanriData.'H2事業所名'

	if( $Kaiso -ne "5" ){

		if( ($HiragiCD -ne [string]$null) -or ($H2JigyoushoName -ne [string]$null) ){
			$ErrorMessage = "$SoshikiCD「$BuShoName」 の階層が5階層ではないのに「ひいらぎコード」または「H2事業所名」がセットされている 階層「$Kaiso」 ひいらぎコード「$HiragiCD」 H2事業所名「$H2JigyoushoName」"
			echo $ErrorMessage
			$ErrStrigngs += $ErrorMessage
		}
	}
}



# ひいらぎコードの重複確認
$ErrorDepartmentInBusyokanriDatas = $department

foreach( $ErrorDepartmentInBusyokanriData in $ErrorDepartmentInBusyokanriDatas ){
	$SoshikiCD = $ErrorDepartmentInBusyokanriData.CD
	$BuShoName = $ErrorDepartmentInBusyokanriData.'部署名'
	$Kaiso = $ErrorDepartmentInBusyokanriData.'階層'
	$HiragiCD = $ErrorDepartmentInBusyokanriData.'ひいらぎコード'
	$H2JigyoushoName = $ErrorDepartmentInBusyokanriData.'H2事業所名'

	if( $HiragiCD -eq [string]$null ){
		continue
	}

	# ひいらぎコードので検索
	[array]$TergetSoshikisByCD = $department | ? 'ひいらぎコード' -eq $HiragiCD
	$Count = $TergetSoshikisByCD.Count
	if( $Count -ne 1 ){
		# $ErrorMessage = "$SoshikiCD 「$BuShoName」 の 「ひいらぎコード」が重複している ひいらぎコード「$HiragiCD」 H2事業所名「$H2JigyoushoName」"
		$ErrorMessage = "「ひいらぎコード」が重複している ひいらぎコード「$HiragiCD」"
		echo $ErrorMessage
		$ErrStrigngs += $ErrorMessage
	}
}


# 5 階層の系統が「通所大」「通所小」になっているデータ確認
$ErrorDepartmentInBusyokanriDatas = $department

foreach( $ErrorDepartmentInBusyokanriData in $ErrorDepartmentInBusyokanriDatas ){
	$SoshikiCD = $ErrorDepartmentInBusyokanriData.CD
	$BuShoName = $ErrorDepartmentInBusyokanriData.'部署名'
	$Kaiso = $ErrorDepartmentInBusyokanriData.'階層'
	$Keito = $ErrorDepartmentInBusyokanriData.'系統'
	$HiragiCD = $ErrorDepartmentInBusyokanriData.'ひいらぎコード'
	$H2JigyoushoName = $ErrorDepartmentInBusyokanriData.'H2事業所名'

	if( $Kaiso -ne "5" ){
		continue
	}

	# 系統確認
	if( ($Keito -eq "通所大") -or ($Keito -eq "通所小") ){
		$ErrorMessage = "$SoshikiCD 「$BuShoName」 の 「系統」に誤りがある 階層「$Kaiso」 系統「$Keito」 ひいらぎコード「$HiragiCD」 H2事業所名「$H2JigyoushoName」"
		echo $ErrorMessage
		$ErrStrigngs += $ErrorMessage
	}
}



# 同一階層の部署名重複確認
$ErrorDepartmentInBusyokanriDatas = $department

foreach( $ErrorDepartmentInBusyokanriData in $ErrorDepartmentInBusyokanriDatas ){
	$SoshikiCD = $ErrorDepartmentInBusyokanriData.CD
	$BuShoName = $ErrorDepartmentInBusyokanriData.'部署名'
	$Kaiso = $ErrorDepartmentInBusyokanriData.'階層'
	$Keito = $ErrorDepartmentInBusyokanriData.'系統'
	$HiragiCD = $ErrorDepartmentInBusyokanriData.'ひいらぎコード'
	$H2JigyoushoName = $ErrorDepartmentInBusyokanriData.'H2事業所名'

	if( $BuShoName -eq [string]$null ){
		continue
	}

	# 部署名/階層で検索
	[array]$TergetSoshikisByName = $department | ? '部署名' -eq $BuShoName | ? '階層' -eq $Kaiso
	$Count = $TergetSoshikisByName.Count
	if( $Count -ne 1 ){
		$ErrorMessage = "$SoshikiCD 「$BuShoName」 同一階層に同名の部署がある 階層「$Kaiso」 部署名「$BuShoName」"
		echo $ErrorMessage
		$ErrStrigngs += $ErrorMessage
	}
}


$ErrStrigngs | Set-Content -Path (Join-Path $DataPath "組織エラーData.txt") -Encoding utf8
$NormalStrigngs | Set-Content -Path (Join-Path $DataPath "組織正常Data.txt") -Encoding utf8


# 従業員の組織が存在するか
$AllEmployees = $employee

$ErrorEmployees = @()

foreach( $EmployeeData in $AllEmployees ){
	$JyugyouinCD = $EmployeeData.'従業員ID'
	if( $JyugyouinCD -eq [string]$null ){
		continue
	}

	$JyugyouinName = $EmployeeData.'従業員名'
	$SyozokuSoshikiCD = $EmployeeData.'組織図組織CD'
	$SyozokuSoshikiName = $EmployeeData.'組織図組織名'
	$SyozokuE2SoshikiCD = $EmployeeData.'E2所属部門CD'
	$SyozokuE2SoshikiName = $EmployeeData.'E2所属部門略称'

	if( $SyozokuE2SoshikiCD -eq [String]$null){
		$ErrorString = "E2所属部門CDが従業員情報にセットされていない : $JyugyouinCD $JyugyouinName 組織図組織CD「$SyozokuSoshikiCD」 組織図組織名「$SyozokuSoshikiName」 E2所属部門CD「$SyozokuE2SoshikiCD」 E2所属部門略称「$SyozokuE2SoshikiName」"
		echo $ErrorString
		$ErrorEmployees += $ErrorString
	}
	else{
		# 所属部門コードで検索
#		try{
#			$SyozokuE2SoshikiCD = ([int]$SyozokuE2SoshikiCD).ToString()
#		}
#		catch{
#			#NOP
#		}

		[array]$TergetSoshikisByCD = $department | ? CD -eq $SyozokuE2SoshikiCD
		$Count = $TergetSoshikisByCD.Count
		if( $Count -eq 0 ){
			$ErrorString = "E2所属部門CD $SyozokuE2SoshikiCD が組織情報に存在しない : $JyugyouinCD $JyugyouinName 組織図組織CD「$SyozokuSoshikiCD」 組織図組織名「$SyozokuSoshikiName」 E2所属部門CD「$SyozokuE2SoshikiCD」 E2所属部門略称「$SyozokuE2SoshikiName」"
			echo $ErrorString
			$ErrorEmployees += $ErrorString
		}
		elseif( $Count -ne 1 ){
			$ErrorString = "E2所属部門CD $SyozokuE2SoshikiCD が組織情報に複数存在している : $JyugyouinCD $JyugyouinName 組織図組織CD「$SyozokuSoshikiCD」 組織図組織名「$SyozokuSoshikiName」 E2所属部門CD「$SyozokuE2SoshikiCD」 E2所属部門略称「$SyozokuE2SoshikiName」 重複数 : $Count"
			echo $ErrorString
			$ErrorEmployees += $ErrorString
		}
	}
}

$ErrorEmployees | Set-Content -Path (Join-Path $DataPath "従業員エラーData.txt") -Encoding utf8



