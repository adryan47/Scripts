#Definicion de Variables

#$NetworkPath = \\ivandrago\Software\FontsLotba\Gotham 

#Se define la variable $LocalPath1 la cual es una ubicacion local oculta en el perfil publico que por defecto el usuario no ve.
$LocalPath1 = "C:\FontsToDeploy\Fonts\*.*"

#Se define la variable LocalPath (lugar temporal para guardar fuentes)
$LocalPath = "C:\Users\Public\Fonts\"

$FONTS = 0x14

#Crea un directorio en $LocalPath
New-Item $LocalPath -type directory -Force

#Copia el contenido de la primer carpeta en la segunda
Copy-Item -Path "C:\FontsToDeploy\Fonts\*.*" -Destination "C:\Users\Public\Fonts\" -Force

#Ejcuta dir sobre la variable $LocalPath y guarda el resultado en la variable $Fontdir
$Fontdir = dir $LocalPath


#Notas-`New-Object` provides the most commonly-used functionality of the VBScript CreateObject function. 
#A statement like `Set objShell = CreateObject("Shell.Application")` in VBScript can be   translated 
#to `$objShell = New-Object -COMObject "Shell.Application"` in PowerShell. 

$Objshell = New-Object -COMObject "Shell.Application"

$shell = New-Object -ComObject "WScript.shell"

$objFolder = $objShell.Namespace($FONTS)



foreach($FONTS in $Fontdir) 
{
  if (!(Test-Path "C:\Windows\Fonts\$FONTS") -eq $False)

    {
     if(($FONTS.name -match "Gotham*.ttf") -or ($FONTS.name -match "Gotham*.otf"))
     {
            
         $objFolder.CopyHere($FONTS.fullname)
        
            $shell.Namespace($LocalPath).Items().InvokeVerbEx("Install")
         #$File.InvokeVerb("Install")
    }
}
}