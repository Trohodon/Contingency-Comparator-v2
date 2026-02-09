<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">

  <Product
    Id="*"
    Name="Contingency Comparator"
    Language="1033"
    Version="1.0.0"
    Manufacturer="Dominion Energy"
    UpgradeCode="PUT-UPGRADE-GUID-HERE">

    <Package InstallerVersion="500" Compressed="yes" InstallScope="perMachine" />
    <MajorUpgrade DowngradeErrorMessage="A newer version is already installed." />
    <MediaTemplate />

    <Directory Id="TARGETDIR" Name="SourceDir">
      <Directory Id="ProgramFilesFolder">
        <Directory Id="INSTALLFOLDER" Name="Contingency Comparator">
          <Component Id="MainExe" Guid="PUT-COMPONENT-GUID-HERE">
            <File Source="..\dist\ContingencyComparator.exe" KeyPath="yes" />
          </Component>
        </Directory>
      </Directory>
    </Directory>

    <Feature Id="MainFeature" Title="Contingency Comparator" Level="1">
      <ComponentRef Id="MainExe" />
    </Feature>

  </Product>

</Wix>