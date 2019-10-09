Param (

    [pscredential]$NETcreds,
    [Parameter(Mandatory)]
    [string[]]$COMports,
    [Parameter(Mandatory)]
    [string]$serverIP,
    [Parameter(Mandatory)]
    [string]$file,
    [string[]]$confcmds,
    [switch]$skip = $false

)

# Function COM-ConnectConsole {

#     Param(

#         [Parameter(Mandatory)]
#         [System.IO.Ports.SerialPort]$COMport

#     )

#     End {

#         # $com = New
#         # $com.close()

#         $com.open()

#         $com.write([char]13)
#         $com.ReadExisting()
#         Write-Host $com.gettype()
#         return $com

#     }

# }

Function Config-ConsoleCOM {


    Param (

        # [Parameter(Mandatory)]
        [pscredential]$creds,
        [Parameter(Mandatory)]
        [string]$COM,
        [Parameter(Mandatory)]
        [string]$tftp_server_IP,
        [parameter(Mandatory)]
        [string]$sfilename,
        [string[]]$COM_cmds

    )

    Begin {

        $com_sess = New-Object System.IO.Ports.SerialPort($COM)
        $cmd = "`$com_sess.open()"

        (Invoke-Expression $cmd -ErrorVariable open_error) 2> $COM_verbose
        if(![string]::IsNullOrWhiteSpace($COM_verbose)){Write-Verbose $COM_verbose}

        if(![string]::IsNullOrWhiteSpace[$open_error]){

            Write-Verbose "Issues with COM port, making continuous attempts to open the port..."

        }

        while(![string]::IsNullOrWhiteSpace($open_error)){

            $com_sess.close()

            (Invoke-Expression $cmd -ErrorVariable open_error) 2> $null

        }
 
        $COM_collect = @{}

        $dfilename = $sfilename

        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.Password)
        $p_wd = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)

    }

    End {

        $init_output = @()

        $com_sess.Write([char]13 + [char]13)
        Start-Sleep -Milliseconds 250
        $output = $com_sess.ReadExisting()
        Write-Verbose $output

        if($output -like "*initial configuration*"){

            $com_sess.WriteLine("no")
            Start-Sleep -Milliseconds 250
            $output = $com_sess.ReadExisting()
            Write-Verbose $output

            $com_sess.Write([char]13 + [char]13)
            Start-Sleep -Milliseconds 250
            $output = $com_sess.ReadExisting()
            Write-Verbose $output

        }

        $com_sess.WriteLine("show version | include Cisco")
        Start-Sleep -Milliseconds 250
        $init_output = $com_sess.ReadExisting()
        Write-Verbose $init_output

        if($init_output -like "*Cisco IOS*" -or $init_output -like "*#*"){

            if($init_output -notlike "*#*"){

                $com_sess.WriteLine("enable")
                Start-Sleep -Milliseconds 250
                $output = $com_sess.ReadExisting()
                Write-Verbose $output
                
                if($output -like "*Password*"){

                    $com_sess.WriteLine($p_wd)
                    Start-Sleep -Milliseconds 250

                }

            }
    
            $com_sess.WriteLine("conf t")
            Start-Sleep -Milliseconds 250
            $output = $com_sess.ReadExisting()
            Write-Verbose $output
    
            $com_sess.WriteLine("int vlan 1")
            Start-Sleep -Milliseconds 250
            $output = $com_sess.ReadExisting()
            Write-Verbose $output
    
            $IP_end = $com_sess.PortName.replace("COM","")

            $com_sess.WriteLine("no shutdown")
            Start-Sleep -Milliseconds 250
            $output = $com_sess.ReadExisting()
            Write-Verbose $output

            $com_sess.WriteLine("ip address dhcp") # Use DHCP in atual network.
            # $com_sess.WriteLine("ip address 172.16.$($IP_end).10 255.255.255.0") # Statically set for testing. Set to DHCP for real scenario.
            Start-Sleep -Seconds 15
            $output = $com_sess.ReadExisting()
            Write-Verbose $output

    
            $com_sess.WriteLine("end")
            Start-Sleep -seconds 10
            $output = $com_sess.ReadExisting()
            Write-Verbose $output

            $com_sess.WriteLine("show ip interface brief | include Vlan1")
            Start-Sleep -Seconds 1
            $output = $com_sess.ReadExisting()
            Write-Verbose $output

            #$IP = $output.split("`n")[1].split("YES")[0].replace("Vlan1","").replace(" ","") 
            if(![string]::IsNullOrWhiteSpace($COM_verbose)){Write-Verbose $IP}

            if([string]::IsNullOrWhiteSpace($sfilename)){$tftp_output = "OK"}

            while($tftp_output -notlike "*OK*") {

                $com_sess.WriteLine("copy tftp: flash:")
                Start-Sleep -Milliseconds 250 
                $output = $com_sess.ReadExisting()
                Write-Verbose $output

                # $tftp_server_IP = "172.16.$($IP_end).5" # For testing purposes, static IP for both TFTP server and Cisco device are set static with the COM port number bbeing the subnet.
                # In real network, TFTP server IP should be provided by CLI argument, and Cisco device IP will be DHCP.

                $com_sess.WriteLine($tftp_server_IP)
                Start-Sleep -Milliseconds 250
                $output = $com_sess.ReadExisting()
                Write-Verbose $output

                $com_sess.WriteLine($sfilename)
                Start-Sleep -Milliseconds 250
                $output = $com_sess.ReadExisting()
                Write-Verbose $output

                $com_sess.WriteLine($dfilename)
                Start-Sleep -Milliseconds 250
                $filename_output = $com_sess.ReadExisting()
                Write-Verbose $filename_output

                $tftp_output = ""

                if($filename_output -like "*existing*"){

                    if(!$skip){
                        $ans = Read-Host "File already exists with same filename. Overwrite? ( [y]es / (n)o / (s)kip )"
                    } else {
                    
                        $ans = 's'

                    }

                    if($ans -eq 'y' -or $ans -eq ""){

                        $com_sess.WriteLine("")
                        Start-Sleep -Seconds 10 # EXTEND THIS TO ACCOMODATE FOR LARGER FILES
                        $tftp_proc = ""
                        Write-Verbose $tftp_proc
                        while($tftp_proc -notlike "*OK*") {
                            $tftp_proc = $com_sess.ReadExisting()

                            if($tftp_proc -like "*OK*"){

                                $tftp_output = "OK"
    
                            }
    
                        }

                    } elseif($ans -eq 's') {

                        $com_sess.WriteLine("n")
                        $tftp_output = "OK"

                    } else {

                        $com_sess.WriteLine("n")
                        $dfilename = Read-Host "Enter new filename"

                    }

                } else {

                    Start-Sleep -Seconds 20 # EXTEND THIS TO ACCOMODATE FOR LARGER FILES
                    $tftp_output = $com_sess.ReadExisting()

                }
                
            }
        
    
            $com_sess.WriteLine("dir " + $filename) # Use this for testing purposes.
            Start-Sleep -Milliseconds 250
            $output = $com_sess.ReadExisting()
            Write-Verbose $output

            $com_sess.WriteLine("conf t")
            Start-Sleep -Milliseconds 250
            Write-Verbose $com_sess.ReadExisting()
    
            $com_sess.WriteLine("boot system flash:/$($sfilename)") # Use this to boot system from the file. 
            Start-Sleep -Seconds 1
            Write-Verbose $com_sess.ReadExisting()
    
            $com_sess.WriteLine("end")
            Start-Sleep -Milliseconds 250
            $output = $com_sess.ReadExisting()
            Write-Verbose $output

            $COM_collect = @{"IP"=$IP;"COMPort"=$com_sess}

            $COM_cmds | % {

                $curr_cmd = $_
                $com_sess.WriteLine($curr_cmd)
                Start-Sleep -Milliseconds 250
                $output = $com_sess.ReadExisting()
                Write-Verbose $output
    

            }

            $com_sess.WriteLine("copy running-config startup-config" + [char]13)
            Start-Sleep -Seconds 5
            $output = $com_sess.ReadExisting()
            Write-Verbose $output

            $com_sess.WriteLine("write mem")
            Start-Sleep -Seconds 5 
            $output = $com_sess.ReadExisting()
            Write-Verbose $output

            $com_sess.Close()

            return $COM_collect

        }


    }

}


$COM_arr = @{}

$COM_ctr = 0

$ports_total = $COMports.Length

$ports_ctr = 0

$COMports | % {

    $c_port = $_

    # COM-ConnectConsole -COMport $com
    $COM_Name = $c_port
    
    $COM_arr.Add($COM_Name,(Config-ConsoleCOM -creds $NETcreds -COM $c_port -tftp_server_IP $serverIP -sfilename $file -COM_cmds $confcmds))

}

$COM_arr.Keys | % {

    $key = $_

    $COM_arr.$key.Keys | % {

        $info = $_

        $COM_arr[$key].$info

    }

}
