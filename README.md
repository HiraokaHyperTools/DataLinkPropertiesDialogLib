# DataLinkPropertiesDialogLib

Display `Data Link Properties` dialog using `MSDASC.DataLinks.PromptNew` / `MSDASC.DataLinks.PromptEdit`:

![Data Link Properties](images/dataLinkProperties_providers.png)

## NuGet

```txt
PM> Install-Package DataLinkPropertiesDialogLib
```

## Usage

```cs
using DataLinkPropertiesDialogLib;

var newConnectionStringOrNull = DataLinkProperties.PromptNew();
```

```cs
using DataLinkPropertiesDialogLib;

var editedConnectionStringOrNull = DataLinkProperties.PromptEdit("Provider=MSDAOSP.1");
```

## Get assembly

Get definition like `MSDASC.DataLinks`:

```bat
tlbimp "C:\Program Files\Common Files\system\ole db\oledb32.dll"
```

Get definition like `ADODB.Connection`:

```bat
tlbimp "C:\Program Files\Common Files\system\ado\msado60.tlb"
```
