Param(      
[Parameter (mandatory=$true,HelpMessage="Defaultbiossettings.xml file path")]
[String]$DefaultBiosSettingsFile,
[Parameter (mandatory=$true,HelpMessage="Flavor of the EG Settings (E.g.: .GN, .AZ)")]
[String]$flavor,
[Parameter (mandatory=$true,HelpMessage="Generation of the SKU (E.g.: 6.0,5.0)")]
[String]$Gen,
[Parameter (mandatory=$true,HelpMessage="Provide Path of the local folder which contains EG Settings files or Press enter if sharepath is input")]
[String]$EGSettingsFilePath
)
#--------------------------------------------------------------------------------------
#Input
#--------------------------------------------------------------------------------------
$DefaultBiosSettingsFile=$DefaultBiosSettingsFile.Trim() -replace '"',''
if(!(Test-Path $DefaultBiosSettingsFile)) {Write-Error "Provide the valid path for Defaultbiossettings.xml" -ErrorAction Stop}

$EGSettingsFilePath=$EGSettingsFilePath.Trim() -replace '"',''
$flavor=$flavor.Trim() -replace '"',''
if(!($flavor -match "(\.+[A-Za-z]{2})")) {Write-Error "Provide the valid flavor for the Bios" -ErrorAction Stop}

$Gen=$Gen.Trim() -replace '"',''
if(!($Gen -as [Int]) -and !($Gen -as [double])) {Write-Error "Provide the valid Generation of the Sku" -ErrorAction Stop}

#--------------------------------------------------------------------------------------
#Get the row and column number of the cell contains the value as token 
#--------------------------------------------------------------------------------------
Function Get-TokenCellValues($EGSettingsExcelData,$EgColumnsList)
{
    for([Int]$row=0;$row -le $EGSettingsExcelData.Count;$row++)
     {
        foreach($property in $EgColumnsList.Name)
            {
               if($EGSettingsExcelData[$row].$property -match "token")
                 {
                     $row
                     $property
                 }
            }
      }
}

#--------------------------------------------------------------------------------------
#return the column of a Cell contains in the format of flavor 
#--------------------------------------------------------------------------------------
Function Check-Flavor($SpecificBiosInList)
{
    $item=$null
    foreach($item in $SpecificBiosInList)
    {
        if($item -match "(\.+[A-Za-z]{2})")
        {
            return $item 
        }
    }
}

#--------------------------------------------------------------------------------------
#Creating a new BiosSetting child in config file for a perticular generation and manufacture
#--------------------------------------------------------------------------------------
Function Append-EGSettings($config,$filePath,$Generation,$xmlBiosSettingsdata)
{
   try
    {
        $fileName=$filePath.Replace(".xlsx","").split('\')[-1]
        try
        {
            $EGSettingsExcelData=Import-XLSX -Path $filePath -Sheet "Gen $Generation.x EG Settings"
        }
        catch
        {
            try
            {
                $EGSettingsExcelData=Import-XLSX -Path $filePath -Sheet "EG Specific Settings"
            }
            catch
            {
                Write-Error "Error: EG Specific Settings or Gen $Generation.x EG Settings Sheet are not available at  $filePath " -ErrorAction Stop
            }
        }

        $EgColumnsList=$EGSettingsExcelData[0] | Get-Member | Where-Object MemberType -EQ NoteProperty | Select-Object Name
        $TokenCellValues= Get-TokenCellValues $EGSettingsExcelData $EgColumnsList
        $rownumber=$TokenCellValues[0]
        $tokenCol=$TokenCellValues[1]

        #BiosSettings Element
        $xmlBiosSettings= $config.CreateElement("BiosSettings")
        $xmlBiosSettingsdata.AppendChild($xmlBiosSettings) | Out-Null
        
        #EgBiosSettings Element
        $xmlEgBiosSettings= $config.CreateElement("EgBiosSettings") 
        $xmlBiosSettings.AppendChild($xmlEgBiosSettings) | Out-Null
        $xmlEgBiosSettings.SetAttribute("FileName","$fileName")
        foreach($property in $EgColumnsList.Name)
         {
            $flavor= $null
            if($EGSettingsExcelData[$rownumber].$property -ne $null)
            {
                $SpecificBiosInList=$EGSettingsExcelData[$rownumber].$property.Split(' ')
                $flavor= Check-Flavor $SpecificBiosInList
            }
            if($flavor -ne $null)
            {
                $EgName=$EGSettingsExcelData[$rownumber].$property -replace '\r*\n', ''
        
                #Eg Element
                $ElementEg= $config.CreateElement("EG") 
                $xmlBiosSettings.AppendChild($ElementEg) | Out-Null
                $ElementEg.SetAttribute("Name","$EgName")

                for([int]$row=1; $row -lt $EGSettingsExcelData.Length;$row++)
                {
                    if($EGSettingsExcelData[$row].$tokenCol -ne $null)
                    {
                        $TokenData=$EGSettingsExcelData[$row].$tokenCol
                        if(($TokenData -match "[a-zA-Z]") -and ($TokenData -ne "N/A"))
                        {
                            $flavorData=$EGSettingsExcelData[$row].$property -replace '\r*\n', ''

                            #Settings node
                            $xmlEgSetting= $config.CreateElement("Settings")
                            $ElementEg.AppendChild($xmlEgSetting) | Out-Null
                            $xmlEgSetting.SetAttribute("Name","$TokenData")
                            $xmlEgSetting.SetAttribute("ExpectedValue","$flavorData")
                        }
                    }
                }
            }
        }
        $Generation=$null 
    }
    catch
    {
        Write-Host "Error: Failed to create the child paths in Bios Setting xml" -ForegroundColor Red 
        continue
    }
}

#--------------------------------------------------------------------------------------
#Creating/updating config.xml with the latest EG Settings 
#--------------------------------------------------------------------------------------
Function Create-ConfigFile($EGSettingsFilePath,$config_Path)
{
    Try
    {
        Write-Host "Updating Config file... "
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
        
    #Creating a config Document object
        [System.Xml.XmlDocument]$config=New-Object System.Xml.XmlDocument 
    #Creating a BiosSettingsData element and added to config doc
        [System.Xml.XmlElement]$xmlBiosSettingsdata= $config.CreateElement("BiosSettingsData")
        $config.AppendChild($xmlBiosSettingsdata) | Out-Null
        

        $Generation=[System.Math]::Floor([convert]::ToDecimal($Gen))
        
        Append-EGSettings $config $EGSettingsFilePath $Generation $xmlBiosSettingsdata
        
        $config.Save($config_Path);

        Write-Host "Successfully updated config file with the new EGSettings data from the file placed at $config_Path"
    }
    catch
    {
        Write-Error "Error: Failed to create/update the config file" -ErrorAction Stop
    }
}

#----------------------------------------------------------------------------------------------------------------------------
#Validate EG Settings with default bios settings
#----------------------------------------------------------------------------------------------------------------------------
Function Validate-EGSettings($DefaultBiosSettingsFileData, $EGSettingsHashTable, $DefaultBiosEGSettingsHashTable, $BiosSettingValidationXmlDocument)
{
  Try
  {
      foreach($DefaultBiosData in $DefaultBiosSettingsFileData.BIOSSetting.Setting)
      {
          $token=$DefaultBiosData.valueDescription.ToString().Trim()
          if($token.ToString().Trim().ToLower() -match "vt for directed i/o")
          {
              $token="Intel VT for Directed I/O (VT-d)"
          }

           if($EGSettingsHashTable.ContainsKey($token))
           {
              $valuefromConfig=$EGSettingsHashTable[$token]
              $value=$DefaultBiosData.value.Trim()
              $value=$value.Split(']')[1]
              $matchFlag="Not Matched"
              try
              {
                  if(($value -match $valuefromConfig) -or ($valuefromConfig -match $value))
                  {        
                      $matchFlag="Matched"
                  }
              }
              catch
              {
                 continue
              }
              $xmlSettings= $BiosSettingValidationXmlDocument.CreateElement("Settings") 
              $xmlValidationResults.AppendChild($xmlSettings) | Out-Null
              $xmlSettings.SetAttribute("Name","$token")
              $xmlSettings.SetAttribute("ExpectedValue","$valuefromConfig")
              if($matchFlag -match "Not Matched")
              {
                  $xmlSettings.SetAttribute("ActualValue","$value")
              }
              $xmlSettings.SetAttribute("Result",$matchFlag)
              $DefaultBiosEGSettingsHashTable+=$token.ToString().ToUpper()
           }  
      }
  }
  Catch
  {
      Write-Host "---------------------------------------------------------------------------------" -ForegroundColor Red
      Try
      {
          $EGSettingsHashTable
      }
      catch
      {
          Write-Error "Error: Unable to Display EG settings. Please check the input files and re-run the validation" -ErrorAction SilentlyContinue
      }
      Write-Host "---------------------------------------------------------------------------------" -ForegroundColor Red

      Write-Error "Error: Failed to validate EGSetting tokens with default bios settings" -ErrorAction SilentlyContinue
  }
  return $DefaultBiosEGSettingsHashTable
}

#----------------------------------------------------------------------------------------------------------------------------
#Get Empty EGSettings from the EGSettings
#----------------------------------------------------------------------------------------------------------------------------
Function Get-EmptyEGSettings($EGSettingsHashTable, $DefaultBiosEGSettingsHashTable, $BiosSettingValidationXmlDocument)
{
    Try
     {
         Foreach($EGtoken in $EGSettingsHashTable.keys)
         {
             If(!($DefaultBiosEGSettingsHashTable.Contains($EGtoken.ToString().Trim().ToUpper())))
             {
                 $xmlSettings= $BiosSettingValidationXmlDocument.CreateElement("Settings") 
                 $xmlValidationResults.AppendChild($xmlSettings) | Out-Null
                 $xmlSettings.SetAttribute("Name",$EGtoken)
                 $xmlSettings.SetAttribute("ExpectedValue",$EGSettingsHashTable[$EGtoken])
                 $xmlSettings.SetAttribute("Result","Setting Not found")
             }
         }
     }
     Catch
     {
         Write-Host "---------------------------------------------------------------------------------" -ForegroundColor Red
         Try
         {
             $EGSettingsHashTable
         }
         catch
         {
             Write-Error "Error: Unable to Display EG settings. Please check the input files and rerun the validation" -ErrorAction SilentlyContinue
         }
         Write-Host "---------------------------------------------------------------------------------" -ForegroundColor Red
         Write-Error "Error: Unable to get the null valued Settings from the EG Settings. Please check with the above EG Settings Data " -ErrorAction SilentlyContinue
     }
}

#----------------------------------------------------------------------------------------------------------------------------
#Creating output(BiosSettingValidation) XML file in the current location By validating EG Settings with default bios settings
#----------------------------------------------------------------------------------------------------------------------------
Function Out-XML($location,$DefaultBiosSettingsFileData,$flavor,$EGSettingsHashTable,$EG_Name)
{
    Try
    {
        $BiosSettingValidationFilePath="$location\BiosSettingValidation.xml"
        if(![System.IO.File]::Exists($BiosSettingValidationFilePath)){
            New-Item "$location\BiosSettingValidation.xml" -ItemType File | Out-Null
        }
        
        $SKUID=$DefaultBiosSettingsFileData.BIOSSetting.skuid
        [System.Xml.XmlDocument]$BiosSettingValidationXmlDocument=New-Object System.Xml.XmlDocument 
        #--------------------------------------------------------------------------------------
        #ValidationResults root
        #--------------------------------------------------------------------------------------
        [System.Xml.XmlElement]$xmlValidationResults= $BiosSettingValidationXmlDocument.CreateElement("ValidationResults") 
        $BiosSettingValidationXmlDocument.AppendChild($xmlValidationResults) | Out-Null
        $xmlValidationResults.SetAttribute("SKUID","$SKUID")
        $xmlValidationResults.SetAttribute("EG","$flavor")
        #--------------------------------------------------------------------------------------
        #EgBiosSettings tag
        #--------------------------------------------------------------------------------------
        $xmlEgBiosSettings= $BiosSettingValidationXmlDocument.CreateElement("EgBiosSettings")
        $xmlValidationResults.AppendChild($xmlEgBiosSettings) | Out-Null
        $xmlEgBiosSettings.SetAttribute("FileName","$EG_Name")
        #$xmlEgBiosSettings.SetAttribute("DateStamp","$ConfigTimeStamp")
        $DefaultBiosEGSettingsHashTable=@()

        $DefaultBiosEGSettingsHashTable = Validate-EGSettings $DefaultBiosSettingsFileData $EGSettingsHashTable $DefaultBiosEGSettingsHashTable $BiosSettingValidationXmlDocument
        
        Get-EmptyEGSettings $EGSettingsHashTable $DefaultBiosEGSettingsHashTable $BiosSettingValidationXmlDocument

        $BiosSettingValidationXmlDocument.Save($BiosSettingValidationFilePath);
        Write-Host "BiosSettings Valiation data is saved at : " -NoNewline; Write-Host "$location\BiosSettingValidation.xml" -ForegroundColor Green
    }
    catch
    {
        Write-Error "Error: Failed to create Bios settings validation xml" -ErrorAction Stop
    }
}
#--------------------------------------------------------------------------------------
#Validate the PSExcel module avaiable in systemdrive, if not Install PSExcel and create config file 
#--------------------------------------------------------------------------------------
Function Validate-PSExcel($EGSettingsFilePath,$config_Path)
{
    Try
    {
        $modulePath="$Env:systemdrive\Program Files\WindowsPowerShell\Modules"
        $psexcelCheck=Get-ChildItem $modulePath | Where-Object Name -eq 'PSExcel'
        
        if($psexcelCheck)
           { 
                Import-Module PSExcel
                Create-ConfigFile $EGSettingsFilePath $config_Path
           }
        else
           {
                Install-Module -Name 'PSExcel' -Force
                Import-Module PSExcel
                Create-ConfigFile $EGSettingsFilePath $config_Path
           }
    }
    catch
    {
        Write-Error "Error: PSExcel module installation failed." -ErrorAction Stop
    }
}

#--------------------------------------------------------------------------------------
#Get the tokens and it's values from the config file for the perticular generation and manufacture
#--------------------------------------------------------------------------------------
Function Get-TokenValues($generationFromConfig)
{
    $EGSettingsHashTable=@{}
    Try
    {
      foreach($EgtokenData in $generationFromConfig.Settings)
       {
            if($EgtokenData.Name.Trim() -match "M.2 vs SSATA") 
            {
                $EGSettingsHashTable.Add("M.2 vs SATA",$EgtokenData.ExpectedValue.Trim())
            }
            elseif($EgtokenData.Name.ToString().Trim().ToLower() -match "vt for directed i/o") 
            {
                $EGSettingsHashTable.Add("Intel VT for Directed I/O (VT-d)",$EgtokenData.ExpectedValue.Trim())
            }
            else
            {
               If(!$EGSettingsHashTable.ContainsKey($EgtokenData.Name.ToString()))
               {
                   $EGSettingsHashTable.Add($EgtokenData.Name.ToString().Trim(),$EgtokenData.ExpectedValue.ToString().Trim())
               }
            }
       }
       
       #$EG_Name=$generationFromConfig.EgBiosSettings.FileName
    }
    catch
    {
        Write-Error "Error: Config file corrupted" -ErrorAction Stop
    }
    return $EGSettingsHashTable
}

#--------------------------------------------------------------------------------------
#Remove unwanted tags from the defaultbiossettings.xml file
#--------------------------------------------------------------------------------------
Function Remove-UnwantedTags($DefaultBiosSettingsFile)
{
    try
    {
        $Deletetags="N/A", "User Name"
        Write-Host "Removed the tags which contains elements having valueDescription as : " -NoNewline; Write-Host "N/A, User Name" -ForegroundColor Green 
        $CurrentLocation=Get-Location
        $xmlData= [xml](Get-Content $DefaultBiosSettingsFile)
        $skuID=$xmlData.BIOSSetting.skuid
        $DefaultBiosSettingsFileNew="$CurrentLocation\$skuID.xml"
        Write-Host "New Bios Setting file is : " -NoNewline; Write-Host "$DefaultBiosSettingsFileNew" -ForegroundColor Green
        
        ($xmlData.BIOSSetting.Setting | Where-Object {$Deletetags -contains $_.valueDescription})| ForEach-Object {
            [void]$_.ParentNode.RemoveChild($_)
        }
        $xmlData.Save($DefaultBiosSettingsFileNew)
    }
    catch
    {
        Write-Error "Error: Failed to remove unwanted tags from default Bios settings xml"
    }
}

Write-Host "*********** Input Data ****************" -ForegroundColor Cyan
Write-Host "DefaultBiosSettings File Path : $DefaultBiosSettingsFile"
Write-Host "EG Settings File Path : $EGSettingsFilePath"
Write-Host "Generation : $Gen"
Write-Host "flavor : $flavor"
Write-Host "***************************************" -ForegroundColor Cyan

#--------------------------------------------------------------------------------------
#Get the XML data from the defaultbiossettings file
#--------------------------------------------------------------------------------------
Try
{
    $DefaultBiosSettingsFileData=[xml](Get-Content $DefaultBiosSettingsFile)
    Write-Host "DefaultBiosSettings data is read successfully."
 }
catch
{
     Write-Error "Error: XML parsing failed" -ErrorAction Stop
 }

#--------------------------------------------------------------------------------------
#Check whether the path is available or not
#--------------------------------------------------------------------------------------
if(Test-Path $EGSettingsFilePath)
  {
        Try
        {
            $location=Get-Location
            $config_Path="$location\Config.xml"
            Validate-PSExcel $EGSettingsFilePath $config_Path
            $config=[xml](Get-Content $config_Path)
            Write-Host "Config data is read successfully."
        }
        catch
        {
             Write-Error "Error: XML parsing failed" -ErrorAction Stop
        }

        $generationFromConfig=$config.BiosSettingsData.BiosSettings

        foreach($EGFromConfig in $generationFromConfig.EG)
        {
            #flavor check
            if($EGFromConfig.Name.toString().Trim().Contains($flavor.Trim()))
            {
                $EGSettingsHashTable=Get-TokenValues $EGFromConfig
                $EG_Name= $generationFromConfig.EgBiosSettings.FileName
            }
        }
        
        if([String]::IsNullOrEmpty($EG_Name)) { Write-Error "Error: Flavor $flavor is not available in the given generation : $gen " -ErrorAction Stop }
        else { Out-XML $location $DefaultBiosSettingsFileData $flavor $EGSettingsHashTable $EG_Name }
        
        $NotMatchedEGSettings=$false
        $ValidatedResults=[xml](Get-Content "$location\BiosSettingValidation.xml")
        foreach($settings in $ValidatedResults.ValidationResults.Settings)
        {
            if($settings.Result -eq "Not Matched")
            {
                $NotMatchedEGSettings=$true
                break
            }
        }

        #Replica of config file with the name of EG Settings and date time stamp
        if(!$NotMatchedEGSettings)
        {
            Write-Host "Would you like to create a replica for Config file.? (Default is No(N))" -ForegroundColor Yellow
            $MovingConfigFile=Read-Host "( y / n )"
            Switch($MovingConfigFile)
            {
                Y {
                    $fileName=$EGSettingsFilePath.Replace(".xlsx","").split('\')[-1]
                    $EGSettingsSheetCreatedDate =(Get-ChildItem $EGSettingsFilePath).CreationTime.ToString().Replace(':','-')
                    $destinationConfigPath= $location.Path+"\"+$fileName+"_"+$EGSettingsSheetCreatedDate+".xml"
                    Copy-Item -Path "$config_Path" -Destination $destinationConfigPath
                    Write-Host "Config file saved as $destinationConfigPath"
                }
                N { Write-Host "Replica of Config file is not made" }
                Default { Write-Host "Replica of Config file is not made" }
            }
        }

        Write-Host "Would you like to remove the unwanted tags from default Bios Settings.? (Default is No(N))" -ForegroundColor Yellow
        $removingUnwantedTags=Read-Host "( y / n )"
        Switch($removingUnwantedTags)
        {
            Y { Remove-UnwantedTags $DefaultBiosSettingsFile }
            N { Write-Host "Unwanted tags are not removed from default Bios Settings" }
            Default { Write-Host "Unwanted tags are not removed from default Bios Settings" }
        }
  }
else
  {
    Write-Host "Path $EGSettingsFilePath doesn't exists" -ForegroundColor Red
  }
