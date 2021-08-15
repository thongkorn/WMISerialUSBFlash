' / --------------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: http://www.facebook.com/g2gnet (For Thailand only)
' / Facebook: http://www.facebook.com/CommonIndy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' / Co-Developer: https://www.facebook.com/apirak.wongwai.5
' /
' / Purpose: Auto Detect and get Serial Number of USB Flash Drive with WMI (Windows Management Instrumentation).
' / Microsoft Visual Basic .NET (2010)
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------------------

'// Don't forget to Add Reference --> System.Management
Imports System.Management
Imports Microsoft.Win32
Imports System.Threading

Public Class frmWMISerialUSBFlash
    '// Win32_VolumeChangeEvent ::  https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-volumechangeevent
    Dim WithEvents pluggedInWatcher As ManagementEventWatcher
    Dim WithEvents pluggedOutWatcher As ManagementEventWatcher
    Dim pluggedInQuery As WqlEventQuery
    Dim pluggedOutQuery As WqlEventQuery
    Dim currentDrive As String = ""

    Private Sub frmWMISerialUSBFlash_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ''// ByPass Cross-thread operation.
        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
        ' Initialize ListView Control
        With lvwData
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
            .HideSelection = False
            .MultiSelect = False
            .Columns.Add("Drive", lvwData.Width \ 3 - 170)
            .Columns.Add("Model", lvwData.Width \ 3 + 115)
            .Columns.Add("Serial Number", lvwData.Width \ 3 + 50)
        End With

        ''// สร้างตัวตรวจจับการเสียบ/ถอด USB (EventType 1 = Configuration Changed, 2 = Device Arrival, 3 = Device Removal, 4 = Docking)
        Try
            ''// เริ่มตรวจจับการเสียบ USB
            pluggedInQuery = New WqlEventQuery
            pluggedInQuery.QueryString = "SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2"
            pluggedInWatcher = New ManagementEventWatcher(pluggedInQuery)
            pluggedInWatcher.Start()

            ''// เริ่มตรวจจับการถอด USB
            pluggedOutQuery = New WqlEventQuery
            pluggedOutQuery.QueryString = "SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 3"
            pluggedOutWatcher = New ManagementEventWatcher(pluggedOutQuery)
            pluggedOutWatcher.Start()

            '// ลูปหา USB Flash Drive ทั้งหมดที่กำลังถูกเสียบคาเครื่องอยู่
            Call ReadUsbFlashSerialNo()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub pluggedInWatcher_EventArrived(sender As Object, e As EventArrivedEventArgs) Handles pluggedInWatcher.EventArrived
        currentDrive = e.NewEvent.Properties("DriveName").Value.ToString

        '// ลูปหา USB Flash Drive ทั้งหมดที่กำลังถูกเสียบคาเครื่องอีกครั้ง
        Call ReadUsbFlashSerialNo()
    End Sub

    Private Sub pluggedOutWatcher_EventArrived(sender As Object, e As EventArrivedEventArgs) Handles pluggedOutWatcher.EventArrived
        currentDrive = e.NewEvent.Properties("DriveName").Value.ToString

        '// ลูปหา USB Flash Drive ทั้งหมดที่กำลังถูกเสียบคาเครื่องอีกครั้ง
        Call ReadUsbFlashSerialNo()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Procedure to read USB Flash Serial Number (Physical Serial Number , Not Volume) and return value all drives.
    Private Sub ReadUsbFlashSerialNo()
        Dim strTemp As String = ""
        Dim strArr() As String      '<-- คำตอบสุดท้าย (Physical Serial Number นั่นเอง)
        Dim MaxLen As Byte = 0      '<-- หาค่าความยาวสูงสุดของชุดข้อมูล ซึ่งจะเป็น Serial Number
        Dim idx As Byte = 0         '<-- เป็น Index ของ Array ซึ่งจะเก็บค่าคำตอบ Serial Number
        '// Start
        Dim thread As Thread = New Thread(Sub()
                                              Try
                                                  '// สร้าง ListViewItem
                                                  Dim LV As ListViewItem

                                                  '// Windows Query Language หรือ WQL "Win32_DiskDrive"
                                                  '// เลือกมาเฉพาะ InterfaceType = 'USB' และ MediaType = 'Removable Media'
                                                  '// เหตุผลที่เลือก MediaType = 'Removable Media' เพราะบางเครื่องอาจมี Card Reader อยู่
                                                  Dim WmiQuery As String = "SELECT Name,Model,PnPDeviceID FROM Win32_DiskDrive WHERE InterfaceType = 'USB' And MediaType = 'Removable Media'"
                                                  Dim WmiSearcher As New ManagementObjectSearcher(WmiQuery)

                                                  '// เคลียร์ค่าใน ListView
                                                  lvwData.Items.Clear()

                                                  '// ลูปหา USB Flash Drive ทั้งหมดที่กำลังถูกเสียบคาเครื่อง
                                                  For Each info As ManagementObject In WmiSearcher.Get()
                                                      '// รับค่า USB Storage จากการอ่านของ WMI
                                                      ' "USBSTOR\DISK&VEN_&PROD_&REV_1.00\7&15FF80F3&0&__&0"
                                                      strTemp = (String.Format("{0}", info("PnPDeviceID")))
                                                      '// แยกชุดข้อมูลออกจากกันด้วยเครื่องหมาย \
                                                      strArr = strTemp.Split("\")
                                                      ' strArr(0) = "USBSTOR"
                                                      ' strArr(1) = "DISK&VEN_&PROD_&REV_1.00"
                                                      ' strArr(2) = "7&15FF80F3&0&__&0"
                                                      '// เลือกเอาข้อมูลชุดสุดท้ายที่มี UBound สูงสุด จะได้ "7&15FF80F3&0&__&0" ซึ่งเป็นคำตอบแต่ยังไม่สุดท้าย
                                                      strTemp = strArr(UBound(strArr))

                                                      strArr = New String() {} ' implicit size, initialized to empty
                                                      '// แยก "7&15FF80F3&0&__&0" ด้วยเครื่องหมาย &
                                                      strArr = strTemp.Split("&")
                                                      ' strArr(0) = "7"               '<-- LBound
                                                      ' strArr(1) = "15FF80F3"
                                                      ' strArr(2) = "0"
                                                      ' strArr(3) = "__"
                                                      ' strArr(4) = "0"               '<-- UBound
                                                      '// ลูปตามจำนวน Index ของ Array โดยเริ่มจาก LBound = 0 ไปจนสิ้นสุดที่ UBound
                                                      '// เพื่อหาช่องเก็บข้อมูลยาวที่สุด
                                                      For Count = LBound(strArr) To UBound(strArr)
                                                          '// เปรียบเทียบค่าความยาวของข้อมูลในแต่ละ Array
                                                          If MaxLen < Len(strArr(Count)) Then
                                                              '// เก็บค่า Index ไว้เป็นคำตอบ (ตัวอย่างคือ Index = 1)
                                                              idx = Count
                                                              '// นำค่าความยาวที่มากกว่ามาเก็บไว้ ก่อนที่วนกลับไปเปรียบเทียบค่าใหม่
                                                              MaxLen = Len(strArr(Count))
                                                          End If
                                                      Next
                                                      '// นำคำตอบสุดท้ายมาแสดงผลใน ListView
                                                      '// Drive Letter
                                                      LV = lvwData.Items.Add(GetDriveLetterFromDisk(String.Format("{0}", info("Name"))))
                                                      '// Model
                                                      LV.SubItems.Add(Trim(info("Model")))
                                                      '// Physical Serial Number
                                                      LV.SubItems.Add(Replace(Trim(strArr(idx)), Chr(0), ""))
                                                  Next
                                              Catch err As ManagementException
                                                  MessageBox.Show("An error occurred while querying for WMI data: " & err.Message)
                                              End Try
                                          End Sub)
        thread.Start()
        thread.Join()
    End Sub

    ' / --------------------------------------------------------------------------------
    ' / Function to read Usb Flash Drive Letter.
    Public Function GetDriveLetterFromDisk(ByVal Name As String) As String
        Dim qPart, qDisk As ObjectQuery
        Dim mPart, mDisk As ManagementObjectSearcher
        Dim objPart, objDisk As ManagementObject
        Dim DriveLetter As String = ""
        ' WMI queries use the "\" as an escape charcter
        Name = Replace(Name, "\", "\\")
        ' First we map the Win32_DiskDrive instance with the association called
        ' Win32_DiskDriveToDiskPartition. Then we map the Win23_DiskPartion
        ' instance with the assocation called Win32_LogicalDiskToPartition
        qPart = New ObjectQuery("ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" & Name & """} WHERE AssocClass = Win32_DiskDriveToDiskPartition")
        mPart = New ManagementObjectSearcher(qPart)
        For Each objPart In mPart.Get()
            qDisk = New ObjectQuery("ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & objPart("DeviceID").ToString & """} WHERE AssocClass = Win32_LogicalDiskToPartition")
            mDisk = New ManagementObjectSearcher(qDisk)
            For Each objDisk In mDisk.Get()
                DriveLetter &= objDisk("Name").ToString
            Next
        Next
        Return DriveLetter
    End Function
End Class
