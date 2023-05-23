<# 
.SYNOPSIS Configurar Assinatura do Microsoft Office 365 (local)
.NOTES	Use por sua conta e risco!
  Version:        1.2
  Author:         	  Erich Oliveira https://www.linkedin.com/in/oliveiraerich/
  Editor: 		  Gustavo Alves Pimenta https://www.linkedin.com/in/gustavo-alves-pimenta-dev/
.COMPONENT Requires Module MSOnline and ExchangeOnlineManagement
#>

##### Conectando ao servidor #####

		#Armazena na variável $LiveCred os dados de acesso fornecidos por Linha de Comando
		Write-Host Conectando ao servidor outlook... -ForegroundColor Green
	     
	    	###$Username = "Usuario com adm"
		###$Password = ConvertTo-SecureString 'Senha $Username' -AsPlainText -Force
		###$LiveCred = New-Object System.Management.Automation.PSCredential $Username, $Password
		
		#Armazena na variável $cred os dados de acesso fornecidos pelo usuário mediante ao PopUp exibido
		$LiveCred = get-credential -message "Digite o email com permissoes de administrador do Office 365"

		#Estabelece conexão com o Office365
		Connect-MsolService -Credential $LiveCred
		Connect-ExchangeOnline -Credential $LiveCred
	

### Verifica a versão do office ###
$Usuario = [Environment]::UserName
$Office365Path = 'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE'

if (!(Test-Path -Path $Office365Path)) {
	$destinyPath = 'c:\users\' + $usuario + '\appdata\roaming\microsoft\assinaturas' #office 2013

} else {
	$destinyPath = 'c:\users\' + $usuario + '\appdata\roaming\microsoft\signatures' ##office 365
}

if(Test-Path -Path $destinyPath) {
	Write-Host "Caminho já existe" -ForegroundColor Red
	
} Else {
	Write-Host "Caminho sendo criado" -ForegroundColor Green
	New-Item -ItemType Directory -Path $destinyPath
			
	if (Test-Path -Path $destinyPath) {
		Write-Host "Caminho criado" -ForegroundColor Green

		} else {
			Write-Host "Caminho não pode ser criado. Contate o criador do script para analise" -ForegroundColor RED
	
		}
	}

##### TEMPLATE transfer #####
$FilesOrigin = "Caminho do Template"

Copy-Item $FilesOrigin -Destination $destinyPath

##### TEMPLATES #####
$FileHtml = $destinyPath +'\Signature.htm'
$FileTxt = $destinyPath + '\Signature.txt'



##### APLICANDO AS ASSINATURAS #####
Write-Host Adicionando assinatura de e-mail -ForegroundColor Green
$CurrentUserEmail = $Usuario + '@seu dominio'

$user = Get-MsolUser -UserPrincipalName $CurrentUserEmail
 
	##### Definindo os campos da assinatura #####
        $UserPrincipalName     = $user.UserPrincipalName       #E-mail
        $DisplayName           = $user.DisplayName             #Full name
		$Title				   = $user.Title		 		   #Job Title
        $Department            = $user.Department              #Department / Field
		$PhoneNumber		   = $user.PhoneNumber			   #Phone Number

	##### ASSINATURA HTML #####
	$SignatureHTML = Get-Content -Path $FileHtml -encoding utf8 -ReadCount 0
	$SignatureTXT = Get-Content -Path $FileTxt -encoding utf8 -ReadCount 0

	
	$SignatureHTML = $SignatureHTML.Replace("[DisplayName]", $DisplayName)
	$SignatureHTML = $SignatureHTML.Replace("[Title]", $Title)
	$SignatureHTML = $SignatureHTML.Replace("[Department]", $Department)
	$SignatureHTML = $SignatureHTML.Replace("[UserPrincipalName]", $UserPrincipalName)

	#### Verifica e atribui número de telefone ####
	if ($PhoneNumber.count -gt 0) {
		$SignatureHTML = $SignatureHTML.Replace("[PhoneNumber]", $PhoneNumber)
		$SignatureTXT = $SignatureTXT.Replace("[PhoneNumber]", $PhoneNumber)
	}	else {
		$SignatureHTML = $SignatureHTML.Replace(", Corporativo: [PhoneNumber]", "")
		$SignatureTXT = $SignatureTXT.Replace(", Corporativo: [PhoneNumber]", "")

	}
	##### ASSINATURA TXT #####
	$SignatureTXT = $SignatureTXT.Replace("[DisplayName]", $DisplayName)
	$SignatureTXT = $SignatureTXT.Replace("[Title]", $Title)
	$SignatureTXT = $SignatureTXT.Replace("[Department]", $Department)
	$SignatureTXT = $SignatureTXT.Replace("[UserPrincipalName]", $UserPrincipalName)

	$SignatureHTML | Set-Content -Path $FileHtml
	$SignatureTXT  | Set-Content -Path $FileTxt

	Write-Host "A assinatura de e-mail do usuario $UserPrincipalName foi atribuida ao Outlook." -ForegroundColor Green
