import sys
import pandas as pd
import clipboard as cb


def check_port(port):
    if port == 0:
        port_name = "ping"
        port_code = 1
    elif port == 21:
        port_name = "ftp"
        port_code = 2
    elif port == 22:
        port_name = "ssh"
        port_code = 3
    elif port == 23:
        port_name = "talent"
        port_code = 4
    elif port == 25:
        port_name = "smtp"
        port_code = 5
    elif port == "137" or port == "138" or port == "139" or port == "445":
        port_name = "netbios"
        port_code = 6
    elif port == 443:
        port_name = "https"
        port_code = 7
    elif port == 80:
        port_name = "http"
        port_code = 8
    elif port in range(3000, 3388) and port in range(3390, 3699):
        port_name = "sap"
        port_code = 9
    elif port == 3389:
        port_name = "terminal service"
        port_code = 10
    elif port == 610:
        port_name = "pop3"
        port_code = 11
    elif port == 995:
        port_name = "pop3 custom"
        port_code = 12
    else:
        port_name = "custom"
        port_code = 13
    return port_name, port_code


def split_ip_port(ip_port):
    s_ip, s_port = ip_port.split(':')
    return s_ip, s_port


def main():
    src_name = "Security TF PC"
    src_dest1 = ['1', '1', '1', '1']
    src_dest2 = ['2', '2', '2', '2']
    src_dest3 = ['3', '3', '3', '3']
    src_dest4 = ['4', '4', '4', '4']
    src_dest5 = ['5', '5', '5', '5']
    src_dest6 = ['6', '6', '6', '6']
    src_dests = [src_dest1, src_dest2, src_dest3, src_dest4, src_dest5, src_dest6]
    trg_code = "1"
    divider = "="*70

    print("\nFirewall Form Generator")
    trg_name = input(" - Enter the target service name: ")
    trg_dests = input(" - Enter IP and Port (separated by comma): ")
    trg_dests = trg_dests.split(",")
    print(divider)

    fw_file = trg_name+".xlsx"
    df = pd.DataFrame(columns=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])
    df.loc[0, 0] = "Source Name"
    df.loc[0, 1] = "Source IP"
    df.loc[0, 5] = "Target Name"
    df.loc[0, 6] = "Target IP"
    df.loc[0, 10] = "Target Code"
    df.loc[0, 11] = "Target Port Code"
    df.loc[0, 12] = "Target Port Number"

    # i, j, k = 0, 0, 1
    print(" - Table logs: ", end='| ')
    k = 1
    for i, src_dst_i in enumerate(src_dests):
        for j, trg_dst_j in enumerate(trg_dests):
            print(i, j , end=' | ')
            df.loc[k, 0] = src_name
            df.loc[k, 1:4] = src_dst_i
            df.loc[k, 5] = trg_name
            df.loc[k, 10] = trg_code
            trg_ip, trg_port = split_ip_port(trg_dst_j)
            trg_ip = str(trg_ip).split('.')
            port_name, port_code = check_port(int(trg_port))
            df.loc[k, 6:9] = trg_ip
            df.loc[k, 11] = port_code
            df.loc[k, 12] = trg_port
            k += 1
    # print('\n'+"="*40+'\n', df)
    df.to_excel(fw_file, index=False, header=False)
    print("\n - Generated file: "+fw_file+"\n"+divider)

    print("Information for Firewall Form")
    rs_1st = " - Source Name: "+src_name+"\n - Target Name: "+trg_name+"\n - IP and Port"
    for j, trg_dst_j in enumerate(trg_dests):
        trg_ip, trg_port = split_ip_port(trg_dst_j)
        rs_1st += "\n   ??? Target IP: "+trg_ip+"\n   ??? Target Port: "+trg_port
    cb.copy(rs_1st)
    print(rs_1st)
    input("Press Enter to next...")
    sys.stdout.write("\033[F")
    print(divider)

    print("Explain for Firewall Form")
    rs_2nd = " - Generate Firewall Form for "+trg_name+" service"
    cb.copy(rs_2nd)
    print(rs_2nd)
    input("Press Enter to exit...")
    sys.stdout.write("\033[F")
    print(divider)


if __name__ == "__main__":
    main() # e.g. Name: test, IP&Port: 1.2.3.4:80,2.3.4.5:443,3.4.5.6:8080
