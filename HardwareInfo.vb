Public Class HardwareInfo
    Public Property details As New Collections.SortedList
    Enum Catagories

        Win32_VideoController
        Win32_VideoSettings
        Win32_SoundDevice
        Win32_USBController
        Win32_USBControllerDevice
        Win32_VoltageProbe
        Win32_BIOS

    End Enum

    Sub New(searchcatagory As Catagories)
        Dim ccount = 0
        Dim scs As New Management.ManagementObjectSearcher("select * from " & System.Enum.GetName(GetType(Catagories), searchcatagory))

        For Each ob As Management.ManagementObject In scs.Get

            details.Add(ccount, New SortedList)


            For Each prop As Management.PropertyData In ob.Properties
                details(ccount).add(prop.Name, prop.Value)
            Next
            ccount += 1
        Next
    End Sub
End Class
