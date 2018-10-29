# Mazak-Machine-Dashboard
Live Dashboards
Public Sub CurrentParser()
    
    '''Overall Goal: Parse through online XML document'''
 
    ''Part 1: Collect the main Attributes of the machine and reading
  
    'Define main worksheet to input the parsed information
    Dim WS As Worksheet
    Set WS = Sheets.Add
    Set WS = ActiveSheet
    
    
    Dim intRand As Long
    intRand = Int((900000) * Rnd) + 100000
    
    'Refresh the page for the updated batch of Probe info
    WS.Cells.ClearContents
    
    'Define main attributes from the heading
    Dim machineName As String
    Dim machineIdentifier As String
    Dim creationTimeDate As String
    Dim creationTimeTime As String
    Dim sender As String
    Dim instanceID As String
    Dim nextSeq As String
    Dim firstSeq As String
    Dim lastSeq As String
    
    'Create variable to keep track of the row index when inputting data
    Dim row As Integer
    row = 1
    
    'Label the rows for the main Device information
    WS.Cells(row, 1) = "Machine Name"
    WS.Cells(row, 2) = "Unique Identifier"
    WS.Cells(row, 3) = "Creation Time: Date"
    WS.Cells(row, 4) = "Creation Time: Time"
    WS.Cells(row, 5) = "Sender"
    WS.Cells(row, 6) = "Instance ID"
    WS.Cells(row, 7) = "Next Sequence Number"
    WS.Cells(row, 8) = "First Sequence Number"
    WS.Cells(row, 9) = "Last Sequence Number"
    

    Dim machines(0) As String 
    machines(0) = "http://mtconnect.mazakcorp.com:5609/sample?" & intRand 'HAAS-VF2
    
    row = row + 1
    For j = 0 To (UBound(machines) - LBound(machines))
    
       
        Dim Req As New XMLHTTP
        Req.Open "GET", machines(j), False 'Remember to just use on NCSU campus since info's only available through their connection
        Req.send
        
        Dim xDoc As DOMDocument30
        Set xDoc = New DOMDocument30
        xDoc.LoadXML Req.responseText
        
        
        machineName = xDoc.SelectSingleNode("//Streams/DeviceStream").Attributes.getNamedItem("name").Text
        machineIdentifier = xDoc.SelectSingleNode("//Streams/DeviceStream").Attributes.getNamedItem("uuid").Text
        
        
        creationTimeDate = xDoc.SelectSingleNode("//MTConnectStreams").ChildNodes(0).Attributes.getNamedItem("creationTime").Text
        creationTimeDate = Left(creationTimeDate, 10)
        creationTimeTime = xDoc.SelectSingleNode("//MTConnectStreams").ChildNodes(0).Attributes.getNamedItem("creationTime").Text
        creationTimeTime = Mid(creationTimeTime, 12, 15)
        sender = xDoc.SelectSingleNode("//MTConnectStreams").ChildNodes(0).Attributes.getNamedItem("sender").Text
        instanceID = xDoc.SelectSingleNode("//MTConnectStreams").ChildNodes(0).Attributes.getNamedItem("instanceId").Text
        nextSeq = xDoc.SelectSingleNode("//MTConnectStreams").ChildNodes(0).Attributes.getNamedItem("nextSequence").Text
        firstSeq = xDoc.SelectSingleNode("//MTConnectStreams").ChildNodes(0).Attributes.getNamedItem("firstSequence").Text
        lastSeq = xDoc.SelectSingleNode("//MTConnectStreams").ChildNodes(0).Attributes.getNamedItem("lastSequence").Text
        
       
        WS.Cells(row, 1) = machineName
        WS.Cells(row, 2) = machineIdentifier
        WS.Cells(row, 3) = creationTimeDate
        WS.Cells(row, 4) = creationTimeTime
        WS.Cells(row, 5) = sender
        WS.Cells(row, 6) = instanceID
        WS.Cells(row, 7) = nextSeq
        WS.Cells(row, 8) = firstSeq
        WS.Cells(row, 9) = lastSeq
        
        row = row + 1
        
    Next j 'ending the first loop that collects the main info for the instance
        
    row = row + 1
  
    WS.Cells(row, 1) = "Machine Name"
    WS.Cells(row, 2) = "Unique Identifier"
    WS.Cells(row, 3) = "Instance ID"
    WS.Cells(row, 4) = "Component"
    WS.Cells(row, 5) = "Component Name"
    WS.Cells(row, 6) = "Component ID"
    WS.Cells(row, 7) = "Data Type"
    WS.Cells(row, 8) = "Data Item Type"
    WS.Cells(row, 9) = "Data Item ID"
    WS.Cells(row, 10) = "Time Stamp: Date"
    WS.Cells(row, 11) = "Time Stamp: Time"
    WS.Cells(row, 12) = "Name"
    WS.Cells(row, 13) = "Sequence"
    WS.Cells(row, 14) = "Type"
    WS.Cells(row, 15) = "SubType"
    WS.Cells(row, 16) = "Value"
    
    row = row + 1
        
    For j = 0 To (UBound(machines) - LBound(machines))

        Req.Open "GET", machines(j), False 'Remember to just use on NCSU campus since info's only available through their connection
        Req.send
        
        
        Set xDoc = New DOMDocument30
        xDoc.LoadXML Req.responseText
        
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim component As String
        Dim componentName As String
        Dim componentID As String
        Dim dataType As String
        Dim dataItemType
        Dim dataItem As String
        
        
        Set mainNode = xDoc.SelectNodes("//DeviceStream/ComponentStream")
        
        For x = 0 To mainNode.Length - 1
            For y = 0 To mainNode(x).ChildNodes.Length - 1
                For z = 0 To mainNode(x).ChildNodes(y).ChildNodes.Length - 1
                    'Input overall Component info
                    machineName = xDoc.SelectSingleNode("//Streams/DeviceStream").Attributes.getNamedItem("name").Text
                    machineIdentifier = xDoc.SelectSingleNode("//Streams/DeviceStream").Attributes.getNamedItem("uuid").Text
                    component = mainNode(x).Attributes.getNamedItem("component").Text
                    componentName = mainNode(x).Attributes.getNamedItem("name").Text
                    componentID = mainNode(x).Attributes.getNamedItem("componentId").Text
                    WS.Cells(row, 1) = machineName
                    WS.Cells(row, 2) = machineIdentifier
                    WS.Cells(row, 3) = instanceID
                    WS.Cells(row, 4) = component
                    WS.Cells(row, 5) = componentName
                    WS.Cells(row, 6) = componentID
                    
                    'Input specific data items
                    'Datatype
                    dataType = mainNode(x).ChildNodes(y).nodeName
                    WS.Cells(row, 7) = dataType
                    
                    'Data Item Type
                    dataItemType = mainNode(x).ChildNodes(y).ChildNodes(z).nodeName
                    WS.Cells(row, 8) = dataItemType
                    
                    'Data Item ID
                    If mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("dataItemId") Is Nothing = False Then
                        dataItem = mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("dataItemId").Text
                        WS.Cells(row, 9) = dataItem
                    End If
                    
                    'Time stamp - date
                    If mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("timestamp") Is Nothing = False Then
                        dataItem = mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("timestamp").Text
                        dataItem = Left(dataItem, 10)
                        WS.Cells(row, 10) = dataItem
                    End If
                    
                    
                    If mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("timestamp") Is Nothing = False Then
                        dataItem = mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("timestamp").Text
                        dataItem = Mid(dataItem, 12, 15)
                        WS.Cells(row, 11) = dataItem
                    End If
                    
                  
                    If mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("name") Is Nothing = False Then
                        dataItem = mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("name").Text
                        WS.Cells(row, 12) = dataItem
                    End If
                    
                    If mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("sequence") Is Nothing = False Then
                        dataItem = mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("sequence").Text
                        WS.Cells(row, 13) = dataItem
                    End If
                    
                    
                    If mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("type") Is Nothing = False Then
                        dataItem = mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("type").Text
                        WS.Cells(row, 14) = dataItem
                    End If
                    
                    
                    If mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("subType") Is Nothing = False Then
                        dataItem = mainNode(x).ChildNodes(y).ChildNodes(z).Attributes.getNamedItem("subType").Text
                        WS.Cells(row, 15) = dataItem
                    End If
                    
                    
                    If mainNode(x).ChildNodes(y).nodeName = "Samples" Then
                        dataItem = mainNode(x).ChildNodes(y).ChildNodes(z).Text
                        WS.Cells(row, 16) = dataItem
                    End If
                    
                    row = row + 1
                    
                Next z
            Next y
        Next x
    Next j 
    

    WS.Range("A1:I3").Name = "currentMainInfoRange" 'Adjust these ranges when adding more machines

    WS.Range("A5:P" & row - 1).Name = "currentDataItemsRange" 'Adjust these ranges when adding more machines
    
    'Autofit the columns
    WS.Columns("A:P").AutoFit

    Application.OnTime Now + TimeValue("00:10:5"), "CurrentParser"
       
End Sub



