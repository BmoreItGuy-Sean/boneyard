#grab new fonts from share
$newFonts = Get-ChildItem -Path C:\Path\Fonts -Recurse -Include *.ttf,*.otf,*.fon
 
#provide lab pc names
$labPCs = @("cem-ws551393","cem-ws551389")
 
#foreach lab $labPC, check for \tempFont directory. if false, create it and copy fonts. if true, copy new fonts to it
ForEach ($pc in $labPCs) {    
 
    #define $localFolder
    $localFolder = "\\$($pc)\c$\tempFont\"
 
    #test path to $localFolder, create and copy fonts if false, else copy fonts
    if (!(Test-Path -Path $localFolder)) {
 
        New-Item -ItemType Directory -Path $localFolder -Force
        Copy-Item -Path $newFonts -Destination $localFolder
 
    }
 
    Else {
    
    Copy-Item -Path $newFonts -Destination $localFolder
    
    }
 
    #run the following commands on target lab PC's
    Invoke-Command -ComputerName $pc -ScriptBlock {
        
        #set dest path
        $destPath = "C:\Windows\Fonts\"
 
        #grab existing dest objects
        $destFiles = Get-ChildItem -Path $destPath -Recurse         
 
        #set new font source path
        $sourcePath = 'C:\tempFont\'
 
        #use this to grab all new font source objects
        $sourceFiles = Get-ChildItem -Path $sourcePath -Include *.ttf,*.otf,*.fon -Recurse
        
        #set directory for placing fonts to install
        $fontInstall = New-Item -Path $sourcePath -ItemType Directory -Name 'Install' -Force
 
        #set path to \Fonts in registry
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"      
 
        #for each source file, see if it exists in \Fonts on the lab PC... 
        ForEach ($s in $sourceFiles) {
 
            $regName = "$($s.Name)"
 
            #if not...
            if (-not(Test-Path ($destPath + $s.Name))) {
 
                #copy to the Install directory...
                Copy-Item -Path $s.FullName -Destination $fontInstall.FullName -Force;
 
                #grab the full path to the font in it's new location...
                $install = $fontInstall.FullName + '\' + $s.Name;
 
                #for each font in the Install directory, copy it to \Fonts via shell object. This copy method is asynchronous...
                #so I used a While loop with Start-Sleep to make sure the copy finishes before the script moves on...
                #logic is 'While the font to install does not exist in \Fonts, sleep the script for 1 second until the asynch copy finishes...
                ForEach ($i in $install) {
 
                    #copy the font file to the \Fonts folder
                    Copy-Item $i -Destination $destPath -Force
 
                    #register the newly copied fonts in the registry
                    New-ItemProperty -Path $regPath -Name $regName -Value $regName -Force | Out-Null;
 
                    }
 
                #register the newly copied fonts in the registry
                New-ItemProperty -Path $regPath -Name $regName -Value $regName -Force | Out-Null
                
                }                
 
            }
 
            #remove the Install folder from the lab PC.
            #Get-Item -Path $fontInstall.FullName | Remove-Item -Recurse -Force
 
        }
 
    }
 
