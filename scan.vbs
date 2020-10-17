Set CommonDialog = CreateObject("WIA.CommonDialog")
Set DeviceManager = CreateObject("WIA.DeviceManager")

Dim i, n
n = DeviceManager.DeviceInfos.Count
WScript.Echo "Number of Scanners found = " & n
For i = 1 to DeviceManager.DeviceInfos.Count
  WScript.Echo " Device " & i & ":" & vbTab & DeviceManager.DeviceInfos(i).Properties("Name").Value
Next

Set DevInfo = DeviceManager.DeviceInfos(1)
Set Device = DevInfo.Connect

Device.Items(1).Properties("6146").Value = 2 'colors
Device.Items(1).Properties("6147").Value = 300 'dots per inch/horizontal
Device.Items(1).Properties("6148").Value = 300 'dots per inch/vertical
Device.Items(1).Properties("6149").Value = 0 'x point where to start scan
Device.Items(1).Properties("6150").Value = 0 'y point where to start scan
Device.Items(1).Properties("6151").Value = 300 'horizontal exent DPI x inches wide
Device.Items(1).Properties("6152").Value = 300 'vertical extent DPI x inches tall
Device.Items(1).Properties("4104").Value = 8 'bits per pixel

Set img = CommonDialog.ShowTransfer(Device.Items(1), "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}", true)

img.SaveFile "img.bmp"
